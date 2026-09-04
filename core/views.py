import datetime
import json
import logging
import re
from functools import wraps
from io import BytesIO
from urllib.parse import quote

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, PermissionRequiredMixin
from django.contrib.auth.views import LoginView, LogoutView
from django.core.exceptions import PermissionDenied
from django.core.serializers.json import DjangoJSONEncoder
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse_lazy
from django.utils.decorators import method_decorator
from django.views import View
from django.views.decorators.csrf import csrf_protect, ensure_csrf_cookie
from django.views.decorators.http import require_POST
from django.views.generic import TemplateView, ListView, CreateView, UpdateView, DeleteView
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from groq import Groq, APIError, APIConnectionError, RateLimitError
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .models import (
    Logo, Sector, Trade, Level, TradeLevel, Trainer,
    Module, LearningOutcome, IndicativeContent, LessonPlan, AssessmentPlan,
    TrainerAccess, sync_trainer_login_account, TrainerLoginConflict,
)
from .forms import (
    LogoForm, SectorForm, TradeForm, LevelForm, TradeLevelForm, TrainerForm,
    ModuleForm, LearningOutcomeForm, IndicativeContentForm, StyledAuthenticationForm,
    LessonPlanForm, AssessmentPlanForm, ModulePickerForm, OutcomePickerForm,
    LearningOutcomeFormSet, IndicativeContentFormSet, TrainerAccessForm,
)

logger = logging.getLogger(__name__)


def get_client():
    """
    Builds a Groq API client from the key configured in .env / settings.
    Raises a clear, human-readable error if the key is missing or still
    set to a placeholder value, instead of letting a cryptic auth error
    bubble up from the Groq SDK.
    """
    api_key = (settings.GROQ_API_KEY or '').strip()
    if not api_key or api_key.startswith('gsk_your') or api_key == 'your-groq-api-key':
        raise RuntimeError(
            "GROQ_API_KEY is not configured. Add a valid key from "
            "https://console.groq.com/keys to your .env file."
        )
    return Groq(api_key=api_key)


# ---------------------------------------------------------------------------
# Shared data payload for the public "generator" pages
# ---------------------------------------------------------------------------
def _build_generator_payload(trainer=None):
    """
    Builds one JSON-serialisable snapshot of the curriculum structure
    (sectors -> trades -> trade/level links -> modules -> learning outcomes
    -> indicative contents, plus trainers and institution logos) so that the
    Scheme of Work / Lesson Plan / Assessment Plan generator pages can drive
    their cascading dropdowns and document preview entirely in the browser,
    without extra round trips to the server.

    `trainer` scopes the payload to a single trainer's own assigned modules
    (and, cascading from that, only the learning outcomes/indicative
    contents that belong to those modules) - this is what keeps a logged-in
    trainer from ever seeing or selecting a module they aren't assigned to.
    Pass None (the default) for the full, unscoped snapshot used by admins.
    """
    sectors = list(Sector.objects.values('id', 'sector_name'))

    trades = list(Trade.objects.values('id', 'trade_name', 'sector_id'))

    levels = list(Level.objects.values('id', 'class_level'))

    trade_levels = list(TradeLevel.objects.values('trade_id', 'level_id'))

    trainers = [
        {'id': t.id, 'name': t.full_name}
        for t in Trainer.objects.all()
    ]

    # Locked modules (is_active=False) are deliberately left out of this
    # payload so they never appear in the generator dropdowns - that's what
    # "blocks" their usage. When `trainer` is set, the queryset is further
    # narrowed to that trainer's own modules only, so a trainer's browser
    # never even receives data for modules assigned to someone else.
    modules_qs = Module.objects.select_related('trainer').filter(is_active=True)
    if trainer is not None:
        modules_qs = modules_qs.filter(trainer=trainer)

    modules = [
        {
            'id': m.id,
            'mod_code': m.mod_code,
            'mod_name': m.mod_name,
            'trade_id': m.trade_id,
            'level_id': m.level_id,
            'trainer_id': m.trainer_id,
            'trainer_name': m.trainer.full_name if m.trainer_id else '',
            'learning_hours': m.learning_hours,
            'term': m.term,
            # Drives the Scheme of Work generator's term/week fields so
            # they're loaded from this module's own record instead of a
            # hardcoded assumption (e.g. always 3 terms of 12 weeks).
            'num_terms': m.num_terms,
            'term_weeks': m.get_term_weeks_list(),
        }
        for m in modules_qs
    ]
    module_ids = {m['id'] for m in modules}

    # Outcomes/contents are cascaded from the (possibly trainer-scoped)
    # module list above, so a trainer's payload never leaks outcomes or
    # indicative content belonging to a module they aren't assigned to.
    outcomes = list(
        LearningOutcome.objects
        .filter(module_id__in=module_ids)
        .values('id', 'module_id', 'outcome_text', 'learning_hours')
    )
    outcome_ids = {o['id'] for o in outcomes}

    contents = list(
        IndicativeContent.objects
        .filter(outcome_id__in=outcome_ids)
        .values('id', 'outcome_id', 'indic_name')
    )

    logos = [

        {'id': l.id, 'name': l.name, 'image': l.image}
        for l in Logo.objects.all()
    ]

    return {
        'sectors': sectors,
        'trades': trades,
        'levels': levels,
        'trade_levels': trade_levels,
        'trainers': trainers,
        'modules': modules,
        'outcomes': outcomes,
        'contents': contents,
        'logos': logos,
    }


def _dump_generator_payload(trainer=None):
    """
    JSON-encodes the generator payload (optionally scoped to `trainer`, see
    _build_generator_payload) and neutralises any literal '</' sequence
    (which could otherwise prematurely close the surrounding <script> tag
    if it appeared inside free-text curriculum content).
    """
    raw = json.dumps(_build_generator_payload(trainer), cls=DjangoJSONEncoder)
    return raw.replace('</', '<\\/')


# ---------------------------------------------------------------------------
# Shared defensive coercion helpers for AI (Groq) JSON responses.
#
# Models don't always respect "return a single string" / "return an array"
# instructions perfectly, so every field pulled out of an AI response goes
# through one of these instead of being trusted as-is - used by both the
# Scheme of Work and Lesson Plan AI endpoints below.
# ---------------------------------------------------------------------------
def _split_items(value):
    """Break a value (list or string) into a clean list of item strings,
    regardless of whether the model used commas, semicolons, newlines,
    or returned a proper array."""
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str):
        parts = re.split(r"\s*(?:\n|;|,)\s*", value)
        return [p.strip() for p in parts if p.strip()]
    return []


def _coerce_to_csv_line(value, max_items=None):
    """
    Defensive normalizer for fields that must render as ONE comma-separated
    line (no bullets, no numbering, no JSON arrays), even if the model
    slips and returns an array instead of a string. Also enforces an
    optional cap on the number of items.
    """
    items = _split_items(value)
    if max_items:
        items = items[:max_items]
    return ", ".join(items)


def _coerce_to_list(value, min_items=None, max_items=None, fallback=None):
    """
    Defensive normalizer for fields that must render as a bullet LIST
    (an array of short item strings), even if the model returns a single
    comma/semicolon/newline-separated string instead of a JSON array.
    Pads with `fallback` items (if provided) to satisfy min_items, and
    truncates to max_items.
    """
    items = _split_items(value)
    if max_items:
        items = items[:max_items]
    if min_items and fallback:
        i = 0
        while len(items) < min_items and i < len(fallback):
            if fallback[i] not in items:
                items.append(fallback[i])
            i += 1
    return items


def _keywords(text):
    """Lower-cased, stopword-free significant words from a short phrase -
    used for a lightweight, best-effort overlap check between fields
    (never a full semantic check, just enough to catch a field that
    drifted completely away from another)."""
    words = re.findall(r"[a-zA-Z]+", (text or "").lower())
    stopwords = {
        "a", "an", "the", "of", "for", "and", "or", "to", "in", "on",
        "with", "this", "that", "is", "are", "by", "as", "at", "from",
        "into", "their", "its", "be", "will", "can",
    }
    return {w for w in words if len(w) > 2 and w not in stopwords}


def _ensure_objectives_cover_steps(objectives, step_titles, max_items=5):
    """
    Best-effort guard so Objectives never silently drift away from what
    the Development/Body steps (i.e. the Range) actually cover: if a step
    title shares no keyword with ANY objective, append one plain,
    grounded fallback objective for that sub-topic. Never removes or
    rewrites what the model returned - only pads gaps, up to max_items.
    """
    result = list(objectives)
    covered_words = set()
    for o in result:
        covered_words |= _keywords(o)

    for title in step_titles:
        if len(result) >= max_items:
            break
        title_words = _keywords(title)
        if not title_words or (title_words & covered_words):
            continue
        clean_title = title.strip().rstrip(".")
        if not clean_title:
            continue
        result.append(f"Perform {clean_title}.")
        covered_words |= title_words

    return result


def _ensure_min_steps(steps, outcome_text, min_count=3):
    """
    Guarantees the Lesson Plan generator always has at least `min_count`
    Development/Body steps, even if the AI returns fewer (e.g. because the
    outcome only has one or two indicative content items). Pads with
    generic-but-usable progressive steps grounded in the outcome text,
    which the trainer can edit before printing.
    """
    filler_bank = [
        {
            "title": "Guided practice",
            "trainer_activity": [
                "Circulates to support trainees applying the outcome hands-on",
                "Checks understanding and corrects common errors",
            ],
            "learner_activity": [
                "Practise the skill individually or in pairs under guidance",
                "Ask questions where unclear",
            ],
            "resources": "Tools & equipment, task sheet",
        },
        {
            "title": "Consolidation and review",
            "trainer_activity": [
                "Reviews key points against the learning outcome",
                "Poses follow-up questions to confirm understanding",
            ],
            "learner_activity": [
                "Summarise what was learned in their own words",
                "Respond to review questions",
            ],
            "resources": "Whiteboard, notebook",
        },
        {
            "title": "Application check",
            "trainer_activity": [
                "Sets a short task applying the outcome to a new scenario",
                "Observes and gives targeted feedback",
            ],
            "learner_activity": [
                "Attempt the task independently",
                "Reflect on feedback received",
            ],
            "resources": "Task sheet, relevant tools/materials",
        },
    ]

    padded = list(steps)
    i = 0
    while len(padded) < min_count and i < len(filler_bank):
        padded.append(filler_bank[i])
        i += 1
    return padded


# ---------------------------------------------------------------------------
# Public landing page
# ---------------------------------------------------------------------------
class LandingView(TemplateView):
    """
    Public entry point for the whole system. Shown before login and presents
    the three main modules as large, clickable tiles.
    """
    template_name = 'core/landing.html'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update({
            'module_count': Module.objects.count(),
            'lesson_plan_count': LessonPlan.objects.count(),
            'assessment_plan_count': AssessmentPlan.objects.count(),
        })
        return context


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
class AdminLoginView(LoginView):
    template_name = 'registration/login.html'
    authentication_form = StyledAuthenticationForm
    redirect_authenticated_user = True


class AdminLogoutView(LogoutView):
    next_page = reverse_lazy('login')


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
class DashboardView(LoginRequiredMixin, TemplateView):
    """
    Shared Dashboard for both admins and trainers. Trainers get the same
    URL/template, just a smaller, module-scoped slice of the context - see
    the trainer branch below and the {% if request.user.trainer_profile %}
    checks in dashboard.html that decide what actually renders.
    """
    template_name = 'core/dashboard.html'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        trainer = getattr(self.request.user, 'trainer_profile', None)

        if trainer:
            # Trainer view: counts scoped to their own assigned modules only.
            outcome_qs = LearningOutcome.objects.filter(module__trainer=trainer)
            content_qs = IndicativeContent.objects.filter(outcome__module__trainer=trainer)
            context.update({
                'is_trainer': True,
                'module_count': Module.objects.filter(trainer=trainer).count(),
                'outcome_count': outcome_qs.count(),
                'content_count': content_qs.count(),
                'recent_outcomes': outcome_qs.select_related('module').order_by('-id')[:5],
            })
            return context

        module_count = Module.objects.count()
        outcome_count = LearningOutcome.objects.count()

        # Trainers currently blocked from generating (free generation used
        # up and no active paid access) - surfaced on the Dashboard so an
        # Admin notices without having to open the Trainer Access screen.
        # has_access is a Python property (depends on today's date), so
        # this is evaluated in Python rather than as a queryset filter.
        pending_access = [
            access for access in
            TrainerAccess.objects.select_related('trainer').all()
            if not access.has_access
        ]

        context.update({
            'is_trainer': False,
            'sector_count': Sector.objects.count(),
            'trade_count': Trade.objects.count(),
            'level_count': Level.objects.count(),
            'module_count': module_count,
            'trainer_count': Trainer.objects.count(),
            'outcome_count': outcome_count,
            'content_count': IndicativeContent.objects.count(),
            'logo_count': Logo.objects.count(),
            'lesson_plan_count': LessonPlan.objects.count(),
            'assessment_plan_count': AssessmentPlan.objects.count(),
            'latest_logo': Logo.objects.order_by('-id').first(),
            # Recently touched records so an admin can jump straight back
            # into what they were working on instead of hunting through
            # the sidebar.
            'recent_modules': Module.objects.select_related('trade', 'level').order_by('-id')[:5],
            'recent_outcomes': LearningOutcome.objects.select_related('module').order_by('-id')[:5],
            # Nudges the "Quick actions" panel towards bulk-adding outcomes
            # first (a module has none yet) vs. contents once outcomes exist.
            'has_modules': module_count > 0,
            'has_outcomes': outcome_count > 0,
            # Trainer generator-access snapshot for the "Trainer Access"
            # KPI card + quick action.
            'pending_access_count': len(pending_access),
            'pending_access_trainers': [a.trainer for a in pending_access][:5],
        })
        return context


# ---------------------------------------------------------------------------
# Trainer access control
#
# Trainers log in through the same login page/User model as admins (see
# Trainer.user), but must only reach Learning Outcomes / Indicative Contents,
# scoped to modules assigned to them. Everything else stays admin-only.
#
# This mixin is the single gate: it's wired into the four generic CRUD base
# classes below, so every model built on top of them (Sectors, Trades,
# Levels, Trade Levels, Modules, Trainers, Logos, Lesson Plans, and the
# Learning Outcome / Indicative Content create-edit-delete screens) is
# admin-only by default. A view opts trainers in explicitly by setting
# trainer_allowed = True.
# ---------------------------------------------------------------------------
class TrainerAccessMixin:
    trainer_allowed = False

    def dispatch(self, request, *args, **kwargs):
        # None for admins/staff (no linked Trainer row) - only set for a
        # logged-in trainer account, and used below to scope querysets.
        self.trainer_profile = getattr(request.user, 'trainer_profile', None)
        if self.trainer_profile and not self.trainer_allowed:
            raise PermissionDenied("Trainers do not have access to this section.")
        return super().dispatch(request, *args, **kwargs)


# ---------------------------------------------------------------------------
# Generic CRUD scaffolding
# ---------------------------------------------------------------------------
class BaseListView(TrainerAccessMixin, LoginRequiredMixin, ListView):
    """
    Generic list view. Subclasses set:
      - model, headers (list[str]), columns (list[str] dotted attribute paths)
      - title, create_url_name, edit_url_name, delete_url_name

    Optionally, subclasses can turn on trainer grouping so the list renders
    as a collapsed-by-default accordion of "Trainer -> their records" instead
    of one flat table. This keeps Modules / Learning Outcomes / Indicative
    Contents organised under the trainer they belong to, and the CRUD table
    for a given trainer only appears once that trainer's group is expanded:
      - group_by_trainer = True
      - trainer_accessor = dotted path from the row object to its Trainer
        (e.g. 'trainer', 'module.trainer', 'outcome.module.trainer')

    A second, nested grouping level can be layered on top of trainer
    grouping so each trainer's records are further collapsed by Module
    (Trainer -> Module -> records), rendered as cards rather than a table
    once the module is expanded:
      - group_by_module = True (requires group_by_trainer = True)
      - module_accessor = dotted path from the row object to its Module
        (e.g. 'module', 'outcome.module')
    """
    template_name = 'core/crud_list.html'
    paginate_by = 25
    headers = []
    columns = []
    title = ''
    create_url_name = ''
    edit_url_name = ''
    delete_url_name = ''
    empty_message = 'No records found.'
    thumbnail_field = None  # e.g. 'image' -> renders a preview thumbnail column

    group_by_trainer = False
    trainer_accessor = 'trainer'
    unassigned_group_label = 'Unassigned'

    group_by_module = False
    module_accessor = 'module'
    unassigned_module_label = 'Unassigned module'

    # Optional secondary "Add multiple" button in the topbar, for models
    # that support the bulk-create workflow (Learning Outcomes / Indicative
    # Contents) alongside the regular single-record "Add" button.
    bulk_create_url_name = ''
    bulk_create_label = ''

    # Optional per-row quick-action link into the *child* model's
    # bulk-create screen, pre-scoped to this row so an admin can go
    # straight from "Modules" to "add outcomes for this module" (or from
    # "Learning Outcomes" to "add contents for this outcome") without
    # re-picking the parent. quick_add_param is the query string key the
    # bulk-create view reads (e.g. 'module' or 'outcome'); the row's own pk
    # is passed as its value.
    quick_add_url_name = ''
    quick_add_label = ''
    quick_add_param = ''

    def get_column_value(self, obj, col):
        value = obj
        for part in col.split('.'):
            if value is None:
                return ''
            value = getattr(value, part, '')
            if callable(value):
                value = value()
        return value

    def get_related_for_obj(self, obj, accessor):
        value = obj
        for part in accessor.split('.'):
            if value is None:
                return None
            value = getattr(value, part, None)
        return value

    def get_trainer_for_obj(self, obj):
        return self.get_related_for_obj(obj, self.trainer_accessor)

    def get_module_for_obj(self, obj):
        return self.get_related_for_obj(obj, self.module_accessor)

    def build_trainer_groups(self, rows):
        """Buckets already-built rows by trainer, preserving each row's
        original order within its group, and sorts groups alphabetically
        by trainer name with the "Unassigned" bucket always last. When
        group_by_module is enabled, each trainer group is further split
        into nested module groups instead of a flat 'rows' list."""
        groups_by_id = {}
        order = []
        for row in rows:
            key = row['trainer_id'] or 0
            if key not in groups_by_id:
                groups_by_id[key] = {
                    'trainer_id': row['trainer_id'],
                    'trainer_name': row['trainer_name'],
                    'rows': [],
                }
                order.append(key)
            groups_by_id[key]['rows'].append(row)

        group_list = [groups_by_id[key] for key in order]
        group_list.sort(key=lambda g: (g['trainer_id'] is None, g['trainer_name'].lower()))

        if self.group_by_module:
            for group in group_list:
                group['modules'] = self.build_module_groups(group.pop('rows'))

        return group_list

    def build_module_groups(self, rows):
        """Buckets a trainer's rows by module, preserving order, and sorts
        alphabetically by module label with any unassigned bucket last."""
        groups_by_id = {}
        order = []
        for row in rows:
            key = row['module_id'] or 0
            if key not in groups_by_id:
                groups_by_id[key] = {
                    'module_id': row['module_id'],
                    'module_label': row['module_label'],
                    'rows': [],
                }
                order.append(key)
            groups_by_id[key]['rows'].append(row)

        group_list = [groups_by_id[key] for key in order]
        group_list.sort(key=lambda g: (g['module_id'] is None, g['module_label'].lower()))
        return group_list

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        rows = []
        for obj in context['object_list']:
            cells = [self.get_column_value(obj, col) for col in self.columns]
            row = {'pk': obj.pk, 'cells': cells, 'fields': list(zip(self.headers, cells))}
            if self.thumbnail_field:
                row['thumbnail'] = self.get_column_value(obj, self.thumbnail_field)
            if self.group_by_trainer:
                trainer = self.get_trainer_for_obj(obj)
                row['trainer_id'] = trainer.pk if trainer else None
                row['trainer_name'] = trainer.full_name if trainer else self.unassigned_group_label
            if self.group_by_module:
                module = self.get_module_for_obj(obj)
                row['module_id'] = module.pk if module else None
                row['module_label'] = str(module) if module else self.unassigned_module_label
            rows.append(row)

        context.update({
            'headers': self.headers,
            'rows': rows,
            'title': self.title,
            'create_url_name': self.create_url_name,
            'edit_url_name': self.edit_url_name,
            'delete_url_name': self.delete_url_name,
            'empty_message': self.empty_message,
            'has_thumbnail': bool(self.thumbnail_field),
            'group_by_trainer': self.group_by_trainer,
            'group_by_module': self.group_by_module,
            'bulk_create_url_name': self.bulk_create_url_name,
            'bulk_create_label': self.bulk_create_label,
            'quick_add_url_name': self.quick_add_url_name,
            'quick_add_label': self.quick_add_label,
            'quick_add_param': self.quick_add_param,
        })

        if self.group_by_trainer:
            context['trainer_groups'] = self.build_trainer_groups(rows)

        return context


class BaseCreateView(TrainerAccessMixin, LoginRequiredMixin, CreateView):
    template_name = 'core/crud_form.html'
    title = ''
    list_url_name = ''

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update({
            'title': f'Add {self.title}',
            'list_url_name': self.list_url_name,
        })
        return context

    def get_success_url(self):
        return reverse_lazy(self.list_url_name)


class BaseUpdateView(TrainerAccessMixin, LoginRequiredMixin, UpdateView):
    template_name = 'core/crud_form.html'
    title = ''
    list_url_name = ''

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update({
            'title': f'Edit {self.title}',
            'list_url_name': self.list_url_name,
        })
        return context

    def get_success_url(self):
        return reverse_lazy(self.list_url_name)


class BaseDeleteView(TrainerAccessMixin, LoginRequiredMixin, DeleteView):
    template_name = 'core/crud_delete.html'
    title = ''
    list_url_name = ''

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update({
            'title': self.title,
            'list_url_name': self.list_url_name,
        })
        return context

    def get_success_url(self):
        return reverse_lazy(self.list_url_name)


# ---------------------------------------------------------------------------
# Sector
# ---------------------------------------------------------------------------
class SectorListView(BaseListView):
    model = Sector
    headers = ['Sector Name']
    columns = ['sector_name']
    title = 'Sectors'
    create_url_name = 'sector-create'
    edit_url_name = 'sector-edit'
    delete_url_name = 'sector-delete'
    empty_message = 'No sectors yet. Click "Add Sector" to create one.'


class SectorCreateView(BaseCreateView):
    model = Sector
    form_class = SectorForm
    title = 'Sector'
    list_url_name = 'sector-list'


class SectorUpdateView(BaseUpdateView):
    model = Sector
    form_class = SectorForm
    title = 'Sector'
    list_url_name = 'sector-list'


class SectorDeleteView(BaseDeleteView):
    model = Sector
    title = 'Sector'
    list_url_name = 'sector-list'


# ---------------------------------------------------------------------------
# Trade
# ---------------------------------------------------------------------------
class TradeListView(BaseListView):
    model = Trade
    headers = ['Trade Name', 'Sector']
    columns = ['trade_name', 'sector.sector_name']
    title = 'Trades'
    create_url_name = 'trade-create'
    edit_url_name = 'trade-edit'
    delete_url_name = 'trade-delete'
    empty_message = 'No trades yet. Click "Add Trade" to create one.'

    def get_queryset(self):
        return super().get_queryset().select_related('sector')


class TradeCreateView(BaseCreateView):
    model = Trade
    form_class = TradeForm
    title = 'Trade'
    list_url_name = 'trade-list'


class TradeUpdateView(BaseUpdateView):
    model = Trade
    form_class = TradeForm
    title = 'Trade'
    list_url_name = 'trade-list'


class TradeDeleteView(BaseDeleteView):
    model = Trade
    title = 'Trade'
    list_url_name = 'trade-list'


# ---------------------------------------------------------------------------
# Level
# ---------------------------------------------------------------------------
class LevelListView(BaseListView):
    model = Level
    headers = ['Class Level']
    columns = ['class_level']
    title = 'Levels'
    create_url_name = 'level-create'
    edit_url_name = 'level-edit'
    delete_url_name = 'level-delete'
    empty_message = 'No levels yet. Click "Add Level" to create one.'


class LevelCreateView(BaseCreateView):
    model = Level
    form_class = LevelForm
    title = 'Level'
    list_url_name = 'level-list'


class LevelUpdateView(BaseUpdateView):
    model = Level
    form_class = LevelForm
    title = 'Level'
    list_url_name = 'level-list'


class LevelDeleteView(BaseDeleteView):
    model = Level
    title = 'Level'
    list_url_name = 'level-list'


# ---------------------------------------------------------------------------
# TradeLevel
# ---------------------------------------------------------------------------
class TradeLevelListView(BaseListView):
    model = TradeLevel
    headers = ['Trade', 'Level']
    columns = ['trade.trade_name', 'level.class_level']
    title = 'Trade Levels'
    create_url_name = 'tradelevel-create'
    edit_url_name = 'tradelevel-edit'
    delete_url_name = 'tradelevel-delete'
    empty_message = 'No trade-level links yet. Click "Add Trade Level" to create one.'

    def get_queryset(self):
        return super().get_queryset().select_related('trade', 'level')


class TradeLevelCreateView(BaseCreateView):
    model = TradeLevel
    form_class = TradeLevelForm
    title = 'Trade Level'
    list_url_name = 'tradelevel-list'


class TradeLevelUpdateView(BaseUpdateView):
    model = TradeLevel
    form_class = TradeLevelForm
    title = 'Trade Level'
    list_url_name = 'tradelevel-list'


class TradeLevelDeleteView(BaseDeleteView):
    model = TradeLevel
    title = 'Trade Level'
    list_url_name = 'tradelevel-list'


# ---------------------------------------------------------------------------
# Trainer
# ---------------------------------------------------------------------------
class TrainerListView(BaseListView):
    model = Trainer
    headers = ['First Name', 'Last Name', 'Username']
    columns = ['fname', 'lname', 'username']
    title = 'Trainers'
    create_url_name = 'trainer-create'
    edit_url_name = 'trainer-edit'
    delete_url_name = 'trainer-delete'
    empty_message = 'No trainers yet. Click "Add Trainer" to create one.'

    # Drives the extra "Login account" column + its manage actions in
    # crud_list.html: create a missing shadow account, or enable/disable
    # an existing one - all without leaving the Dashboard for raw SQL.
    show_login_management = True
    login_toggle_url_name = 'trainer-login-toggle'
    login_repair_url_name = 'trainer-login-repair'

    def get_queryset(self):
        return super().get_queryset().select_related('user')

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        # Stamp each row with its trainer's own login-account state so the
        # template can render an Active/Inactive/No account badge plus the
        # matching action button, the same way ModuleListView stamps
        # is_active for the lock/unlock badge above.
        trainers_by_pk = {obj.pk: obj for obj in context['object_list']}
        for row in context['rows']:
            trainer = trainers_by_pk.get(row['pk'])
            user = trainer.user if trainer else None
            if user is None:
                row['login_status'] = 'missing'
            elif user.is_active:
                row['login_status'] = 'active'
            else:
                row['login_status'] = 'inactive'
            row['login_last_login'] = user.last_login if user else None

        context.update({
            'show_login_management': self.show_login_management,
            'login_toggle_url_name': self.login_toggle_url_name,
            'login_repair_url_name': self.login_repair_url_name,
        })
        return context


class TrainerCreateView(BaseCreateView):
    model = Trainer
    form_class = TrainerForm
    title = 'Trainer'
    list_url_name = 'trainer-list'


class TrainerUpdateView(BaseUpdateView):
    model = Trainer
    form_class = TrainerForm
    title = 'Trainer'
    list_url_name = 'trainer-list'


class TrainerLoginToggleView(TrainerAccessMixin, LoginRequiredMixin, View):
    """
    Handles the Enable/Disable buttons in the Trainers list's "Login
    account" column. Flips the trainer's shadow auth.User.is_active, which
    TrainerBackend checks on every login attempt - so this blocks/restores
    a trainer's ability to sign in without touching their password or
    deleting anything. Admin-only (default TrainerAccessMixin behaviour);
    POST-only so it can't be triggered by a GET/crawler/back-button replay.
    """
    @method_decorator(require_POST)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, pk):
        trainer = get_object_or_404(Trainer, pk=pk)
        if trainer.user is None:
            messages.error(
                request,
                f'"{trainer.full_name}" has no login account yet - use '
                '"Create login" first.'
            )
            return redirect('trainer-list')

        trainer.user.is_active = not trainer.user.is_active
        trainer.user.save(update_fields=['is_active'])

        state = 'enabled' if trainer.user.is_active else 'disabled'
        messages.success(request, f'Login for "{trainer.full_name}" is now {state}.')
        return redirect('trainer-list')


class TrainerLoginRepairView(TrainerAccessMixin, LoginRequiredMixin, View):
    """
    Handles the "Create login" button in the Trainers list for a trainer
    that has no linked auth.User yet (e.g. one created before Trainer.user
    existed, or left behind by a previous manual database fix). Runs the
    same account creation/relinking logic TrainerForm uses on every save
    (see models.sync_trainer_login_account), so an admin never has to open
    a database console to get a trainer signed-in again.
    """
    @method_decorator(require_POST)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, pk):
        trainer = get_object_or_404(Trainer, pk=pk)
        try:
            sync_trainer_login_account(trainer)
        except TrainerLoginConflict as exc:
            messages.error(request, str(exc))
        else:
            messages.success(
                request,
                f'Login account created for "{trainer.full_name}".'
            )
        return redirect('trainer-list')


class TrainerDeleteView(BaseDeleteView):
    model = Trainer
    title = 'Trainer'
    list_url_name = 'trainer-list'

    def form_valid(self, form):
        """
        Deleting a Trainer only detaches trainer.user (on_delete=SET_NULL)
        - it does NOT remove the shadow auth.User login account that went
        with it, which would otherwise sit in auth_user forever, holding
        that username and blocking it from ever being reused by a future
        trainer (see models.sync_trainer_login_account). That shadow account
        has no purpose without its Trainer (unusable password, no
        staff/superuser rights, login only ever happens via
        core_trainer.password_hash), so it's safe to remove here too.
        """
        user = self.object.user
        if user is not None and not user.is_staff and not user.is_superuser:
            user.delete()
        return super().form_valid(form)


# ---------------------------------------------------------------------------
# Trainer Access (grant/revoke generator access)
#
# Previously the ONLY way to flip a trainer's paid access on/off was the
# separate Django /admin site (TrainerAccessAdmin). This section adds that
# same capability straight into the Admin Dashboard, admin-only (trainers
# never see this - TrainerAccessMixin's default trainer_allowed = False
# already blocks them, same as every other admin-only screen).
# ---------------------------------------------------------------------------
ACCESS_GRANT_DAYS = 90


class TrainerAccessListView(TrainerAccessMixin, LoginRequiredMixin, ListView):
    """
    One screen listing every trainer's generator access status (free
    generation used?, paid?, expiry, payment note) with one-click
    Grant/Revoke actions plus a link into the fine-grained edit form.
    """
    model = TrainerAccess
    template_name = 'core/trainer_access_list.html'
    context_object_name = 'access_rows'
    paginate_by = 25

    def get_queryset(self):
        return (
            TrainerAccess.objects.select_related('trainer')
            .order_by('trainer__lname', 'trainer__fname')
        )

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = 'Trainer Access'
        context['grant_days'] = ACCESS_GRANT_DAYS
        return context


class TrainerAccessGrantView(TrainerAccessMixin, LoginRequiredMixin, View):
    """
    Handles the "Grant Access" button: marks the trainer Paid and opens a
    fresh 90-day window from today, every time it's clicked - regardless
    of any previous paid_until date - so an Admin can use this to both
    grant NEW access and renew/extend an existing (even still-active)
    one. Equivalent to (and replaces the need for) manually editing
    TrainerAccess in the Django /admin site.
    """
    @method_decorator(require_POST)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, pk):
        access = get_object_or_404(TrainerAccess.objects.select_related('trainer'), pk=pk)
        access.is_paid = True
        access.paid_until = datetime.date.today() + datetime.timedelta(days=ACCESS_GRANT_DAYS)
        access.save(update_fields=['is_paid', 'paid_until', 'updated_at'])

        messages.success(
            request,
            f'Access granted for {access.trainer.full_name} until '
            f'{access.paid_until.strftime("%d %b %Y")}.'
        )
        return redirect('traineraccess-list')


class TrainerAccessRevokeView(TrainerAccessMixin, LoginRequiredMixin, View):
    """
    Handles the "Revoke Access" button: turns paid access off immediately.
    The trainer keeps their existing free-generation status untouched
    (this only removes PAID access), and can be re-granted at any time.
    """
    @method_decorator(require_POST)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, pk):
        access = get_object_or_404(TrainerAccess.objects.select_related('trainer'), pk=pk)
        access.is_paid = False
        access.save(update_fields=['is_paid', 'updated_at'])

        messages.success(request, f'Paid access revoked for {access.trainer.full_name}.')
        return redirect('traineraccess-list')


class TrainerAccessUpdateView(BaseUpdateView):
    """
    Fine-grained edit form (custom paid_until date, payment reference, or
    resetting the one-time free generation) for cases the one-click
    Grant/Revoke buttons don't cover.
    """
    model = TrainerAccess
    form_class = TrainerAccessForm
    title = 'Trainer Access'
    list_url_name = 'traineraccess-list'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['title'] = f'Edit Access — {self.object.trainer.full_name}'
        return context


# ---------------------------------------------------------------------------
# Module
# ---------------------------------------------------------------------------
class ModuleListView(BaseListView):
    model = Module
    # "Trainer" is dropped from the columns here because the list is grouped
    # by trainer below - the group header already carries that information.
    headers = ['Code', 'Name', 'Trade', 'Level', 'Hours', 'Term']
    columns = ['mod_code', 'mod_name', 'trade.trade_name', 'level.class_level', 'learning_hours', 'term']
    title = 'Modules'
    create_url_name = 'module-create'
    edit_url_name = 'module-edit'
    delete_url_name = 'module-delete'
    empty_message = 'No modules yet. Click "Add Module" to create one.'

    group_by_trainer = True
    trainer_accessor = 'trainer'

    # Jump straight from a module row to bulk-adding its Learning Outcomes.
    quick_add_url_name = 'outcome-bulk-create'
    quick_add_label = 'Add outcomes'
    quick_add_param = 'module'

    # Renders the Lock/Unlock status badge + Turn On/Turn Off button on
    # each row. The button itself is only shown to users who hold the
    # 'core.toggle_module_status' permission - see get_context_data below.
    toggle_url_name = 'module-toggle-active'

    def get_queryset(self):
        return super().get_queryset().select_related('trade', 'level', 'trainer')

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        # Stamp each row with its module's own lock state so the template
        # can render a Locked/Active badge and the correct button label.
        objects_by_pk = {obj.pk: obj for obj in context['object_list']}
        for row in context['rows']:
            module = objects_by_pk.get(row['pk'])
            row['is_active'] = module.is_active if module else True

        context.update({
            'toggle_url_name': self.toggle_url_name,
            # Rights check happens here, once, instead of per-row in the
            # template: only users with this permission ever see the
            # Turn On/Turn Off button at all.
            'can_toggle_module': self.request.user.has_perm('core.toggle_module_status'),
        })
        return context


class ModuleCreateView(BaseCreateView):
    model = Module
    form_class = ModuleForm
    title = 'Module'
    list_url_name = 'module-list'


class ModuleUpdateView(BaseUpdateView):
    model = Module
    form_class = ModuleForm
    title = 'Module'
    list_url_name = 'module-list'


class ModuleDeleteView(BaseDeleteView):
    model = Module
    title = 'Module'
    list_url_name = 'module-list'


class ModuleToggleActiveView(TrainerAccessMixin, LoginRequiredMixin, PermissionRequiredMixin, View):
    """
    Handles the Turn On / Turn Off buttons on the Modules list.

    Rights check: only users who hold the 'core.toggle_module_status'
    permission may flip a module's lock state. Anyone else - even a
    logged-in admin without that permission - gets a 403 (raise_exception)
    instead of silently redirecting to login, since they *are* authenticated,
    they just aren't allowed to do this particular action.

    POST-only (guarded further by @require_POST via dispatch) so the toggle
    can never be triggered by a GET request/crawler/back-button replay.
    """
    permission_required = 'core.toggle_module_status'
    raise_exception = True

    @method_decorator(require_POST)
    def dispatch(self, request, *args, **kwargs):
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, pk):
        module = get_object_or_404(Module, pk=pk)
        module.is_active = not module.is_active
        module.save(update_fields=['is_active'])

        state = 'unlocked' if module.is_active else 'locked'
        messages.success(
            request,
            f'Module "{module.mod_code} - {module.mod_name}" is now {state}.'
        )
        return redirect('module-list')


# ---------------------------------------------------------------------------
# LearningOutcome
# ---------------------------------------------------------------------------
class LearningOutcomeListView(BaseListView):
    model = LearningOutcome
    # "Module" is dropped from the columns here because the list is grouped
    # Trainer -> Module below - the module group header already carries it.
    headers = ['Outcome', 'Hours']
    columns = ['outcome_text', 'learning_hours']
    title = 'Learning Outcomes'
    create_url_name = 'outcome-create'
    edit_url_name = 'outcome-edit'
    delete_url_name = 'outcome-delete'
    empty_message = 'No learning outcomes yet. Click "Add Learning Outcome" to create one.'

    group_by_trainer = True
    trainer_accessor = 'module.trainer'
    group_by_module = True
    module_accessor = 'module'

    # Topbar shortcut into the bulk-create screen…
    bulk_create_url_name = 'outcome-bulk-create'
    bulk_create_label = 'Add multiple'
    # …and, per-row, straight on to bulk-adding this outcome's contents.
    quick_add_url_name = 'content-bulk-create'
    quick_add_label = 'Add contents'
    quick_add_param = 'outcome'

    # Trainers can view this screen - one of the two things they're
    # allowed to see. The queryset below scopes it to their own modules.
    trainer_allowed = True

    def get_queryset(self):
        qs = super().get_queryset().select_related('module__trainer')
        if self.trainer_profile:
            qs = qs.filter(module__trainer=self.trainer_profile)
        return qs


class LearningOutcomeCreateView(BaseCreateView):
    model = LearningOutcome
    form_class = LearningOutcomeForm
    title = 'Learning Outcome'
    list_url_name = 'outcome-list'


class LearningOutcomeUpdateView(BaseUpdateView):
    model = LearningOutcome
    form_class = LearningOutcomeForm
    title = 'Learning Outcome'
    list_url_name = 'outcome-list'


class LearningOutcomeDeleteView(BaseDeleteView):
    model = LearningOutcome
    title = 'Learning Outcome'
    list_url_name = 'outcome-list'


# ---------------------------------------------------------------------------
# IndicativeContent
# ---------------------------------------------------------------------------
class IndicativeContentListView(BaseListView):
    model = IndicativeContent
    headers = ['Outcome', 'Indicative Content']
    columns = ['outcome.outcome_text', 'indic_name']
    title = 'Indicative Contents'
    create_url_name = 'content-create'
    edit_url_name = 'content-edit'
    delete_url_name = 'content-delete'
    empty_message = 'No indicative contents yet. Click "Add Indicative Content" to create one.'

    group_by_trainer = True
    trainer_accessor = 'outcome.module.trainer'
    group_by_module = True
    module_accessor = 'outcome.module'

    bulk_create_url_name = 'content-bulk-create'
    bulk_create_label = 'Add multiple'

    # Trainers can view this screen - the second of the two things they're
    # allowed to see. The queryset below scopes it to their own modules.
    trainer_allowed = True

    def get_queryset(self):
        qs = super().get_queryset().select_related('outcome__module__trainer')
        if self.trainer_profile:
            qs = qs.filter(outcome__module__trainer=self.trainer_profile)
        return qs


class IndicativeContentCreateView(BaseCreateView):
    model = IndicativeContent
    form_class = IndicativeContentForm
    title = 'Indicative Content'
    list_url_name = 'content-list'


class IndicativeContentUpdateView(BaseUpdateView):
    model = IndicativeContent
    form_class = IndicativeContentForm
    title = 'Indicative Content'
    list_url_name = 'content-list'


class IndicativeContentDeleteView(BaseDeleteView):
    model = IndicativeContent
    title = 'Indicative Content'
    list_url_name = 'content-list'


# ---------------------------------------------------------------------------
# Bulk-create: "pick a parent once, add several children"
#
# A module usually needs several Learning Outcomes, and each outcome usually
# needs several Indicative Contents. Creating those one-at-a-time through the
# single-record form means re-selecting the same parent over and over. These
# two views let an admin pick the parent once and add a whole batch of
# children in a single submit.
# ---------------------------------------------------------------------------
class LearningOutcomeBulkCreateView(TrainerAccessMixin, LoginRequiredMixin, View):
    template_name = 'core/bulk_outcome_form.html'
    picker_form_class = ModulePickerForm
    formset_class = LearningOutcomeFormSet
    formset_prefix = 'outcomes'

    def get(self, request):
        selected_module = None
        module_id = request.GET.get('module')
        if module_id:
            selected_module = Module.objects.filter(pk=module_id).select_related('trade', 'level').first()

        picker_form = self.picker_form_class(
            initial={'module': selected_module.pk} if selected_module else None
        )
        formset = self.formset_class(prefix=self.formset_prefix)
        return render(request, self.template_name, self._context(picker_form, formset, selected_module))

    def post(self, request):
        picker_form = self.picker_form_class(request.POST)
        formset = self.formset_class(request.POST, prefix=self.formset_prefix)

        # If the module was locked (pre-selected via a quick-add link), keep
        # showing it locked on redisplay even if the rest of the form fails.
        selected_module = None
        if request.POST.get('lock_module') == '1':
            selected_module = Module.objects.filter(
                pk=request.POST.get('module')
            ).select_related('trade', 'level').first()

        if picker_form.is_valid() and formset.is_valid():
            module = picker_form.cleaned_data['module']
            created = 0
            for form in formset:
                if not form.has_changed() or form.cleaned_data.get('DELETE'):
                    continue
                outcome_text = (form.cleaned_data.get('outcome_text') or '').strip()
                if not outcome_text:
                    continue
                LearningOutcome.objects.create(
                    module=module,
                    outcome_text=outcome_text,
                    learning_hours=form.cleaned_data.get('learning_hours') or 0,
                )
                created += 1

            if created:
                messages.success(
                    request,
                    f'Added {created} learning outcome{"s" if created != 1 else ""} to "{module}".'
                )
                return redirect('outcome-list')
            messages.warning(request, 'Add at least one learning outcome before saving.')

        return render(request, self.template_name, self._context(picker_form, formset, selected_module))

    def _context(self, picker_form, formset, selected_module):
        return {
            'picker_form': picker_form,
            'formset': formset,
            'selected_module': selected_module,
            'title': 'Add Multiple Learning Outcomes',
            'parent_label': 'module',
        }


class IndicativeContentBulkCreateView(TrainerAccessMixin, LoginRequiredMixin, View):
    template_name = 'core/bulk_content_form.html'
    picker_form_class = OutcomePickerForm
    formset_class = IndicativeContentFormSet
    formset_prefix = 'contents'

    def get(self, request):
        selected_outcome = None
        outcome_id = request.GET.get('outcome')
        if outcome_id:
            selected_outcome = LearningOutcome.objects.filter(
                pk=outcome_id
            ).select_related('module').first()

        picker_form = self.picker_form_class(
            initial={'outcome': selected_outcome.pk} if selected_outcome else None
        )
        formset = self.formset_class(prefix=self.formset_prefix)
        return render(request, self.template_name, self._context(picker_form, formset, selected_outcome))

    def post(self, request):
        picker_form = self.picker_form_class(request.POST)
        formset = self.formset_class(request.POST, prefix=self.formset_prefix)

        selected_outcome = None
        if request.POST.get('lock_outcome') == '1':
            selected_outcome = LearningOutcome.objects.filter(
                pk=request.POST.get('outcome')
            ).select_related('module').first()

        if picker_form.is_valid() and formset.is_valid():
            outcome = picker_form.cleaned_data['outcome']
            created = 0
            for form in formset:
                if not form.has_changed() or form.cleaned_data.get('DELETE'):
                    continue
                indic_name = (form.cleaned_data.get('indic_name') or '').strip()
                if not indic_name:
                    continue
                IndicativeContent.objects.create(outcome=outcome, indic_name=indic_name)
                created += 1

            if created:
                messages.success(
                    request,
                    f'Added {created} indicative content{"s" if created != 1 else ""} to "{outcome}".'
                )
                return redirect('content-list')
            messages.warning(request, 'Add at least one indicative content before saving.')

        return render(request, self.template_name, self._context(picker_form, formset, selected_outcome))

    def _context(self, picker_form, formset, selected_outcome):
        return {
            'picker_form': picker_form,
            'formset': formset,
            'selected_outcome': selected_outcome,
            'title': 'Add Multiple Indicative Contents',
            'parent_label': 'outcome',
        }


# ---------------------------------------------------------------------------
# Logo
# ---------------------------------------------------------------------------
class LogoListView(BaseListView):
    model = Logo
    headers = ['Name']
    columns = ['name']
    thumbnail_field = 'image'
    title = 'Logos'
    create_url_name = 'logo-create'
    edit_url_name = 'logo-edit'
    delete_url_name = 'logo-delete'
    empty_message = 'No logos yet. Click "Add Logo" to create one.'


class LogoCreateView(BaseCreateView):
    model = Logo
    form_class = LogoForm
    title = 'Logo'
    list_url_name = 'logo-list'


class LogoUpdateView(BaseUpdateView):
    model = Logo
    form_class = LogoForm
    title = 'Logo'
    list_url_name = 'logo-list'


class LogoDeleteView(BaseDeleteView):
    model = Logo
    title = 'Logo'
    list_url_name = 'logo-list'


# ---------------------------------------------------------------------------
# Lesson Plans
# ---------------------------------------------------------------------------
class LessonPlanListView(BaseListView):
    model = LessonPlan
    headers = ['Title', 'Module', 'Trainer', 'Week', 'Date']
    columns = ['title', 'module.mod_code', 'trainer.full_name', 'week', 'lesson_date']
    title = 'Lesson Plans'
    create_url_name = 'lessonplan-create'
    edit_url_name = 'lessonplan-edit'
    delete_url_name = 'lessonplan-delete'
    empty_message = 'No lesson plans yet. Click "Add Lesson Plan" to create one.'

    def get_queryset(self):
        return super().get_queryset().select_related('module', 'trainer')


class LessonPlanCreateView(BaseCreateView):
    model = LessonPlan
    form_class = LessonPlanForm
    title = 'Lesson Plan'
    list_url_name = 'lessonplan-list'


class LessonPlanUpdateView(BaseUpdateView):
    model = LessonPlan
    form_class = LessonPlanForm
    title = 'Lesson Plan'
    list_url_name = 'lessonplan-list'


class LessonPlanDeleteView(BaseDeleteView):
    model = LessonPlan
    title = 'Lesson Plan'
    list_url_name = 'lessonplan-list'


# ---------------------------------------------------------------------------
# Assessment Plan
# ---------------------------------------------------------------------------
class AssessmentPlanListView(BaseListView):
    model = AssessmentPlan
    headers = ['Module', 'Trainer', 'Type', 'Date', 'Candidates']
    columns = ['module.mod_code', 'trainer.full_name', 'get_assessment_type_display', 'assessment_date', 'num_candidates']
    title = 'Assessment Plans'
    create_url_name = 'assessmentplan-create'
    edit_url_name = 'assessmentplan-edit'
    delete_url_name = 'assessmentplan-delete'
    empty_message = 'No assessment plans yet. Click "Add Assessment Plan" to create one.'

    def get_queryset(self):
        return super().get_queryset().select_related('module', 'trainer')


class AssessmentPlanCreateView(BaseCreateView):
    model = AssessmentPlan
    form_class = AssessmentPlanForm
    title = 'Assessment Plan'
    list_url_name = 'assessmentplan-list'


class AssessmentPlanUpdateView(BaseUpdateView):
    model = AssessmentPlan
    form_class = AssessmentPlanForm
    title = 'Assessment Plan'
    list_url_name = 'assessmentplan-list'


class AssessmentPlanDeleteView(BaseDeleteView):
    model = AssessmentPlan
    title = 'Assessment Plan'
    list_url_name = 'assessmentplan-list'


# ---------------------------------------------------------------------------
# Public generator pages (Scheme of Work / Lesson Plan / Assessment Plan)
# ---------------------------------------------------------------------------
def _generator_header_context(request):
    """
    Shared header context for the three login-gated generator pages below.

    - `tms_data_json`: the curriculum snapshot driving the page's cascading
      dropdowns, scoped to the logged-in trainer's own assigned modules
      only (admins/staff with no linked Trainer profile still get the full,
      unscoped snapshot - they aren't "assigned" a subset of modules).
    - `is_trainer` / `assigned_module_count`: drive the page header's
      signed-in summary (who's logged in and how much data/access they
      have), replacing the old static "Public tool - no login required"
      copy now that these pages require login.
    """
    trainer = getattr(request.user, 'trainer_profile', None)
    if trainer is not None:
        module_count = Module.objects.filter(trainer=trainer, is_active=True).count()
    else:
        module_count = Module.objects.filter(is_active=True).count()

    return {
        'tms_data_json': _dump_generator_payload(trainer),
        'is_trainer': trainer is not None,
        'assigned_module_count': module_count,
    }


@method_decorator(ensure_csrf_cookie, name='dispatch')
class SchemeOfWorkUserView(LoginRequiredMixin, TemplateView):
    """
    Generator page, now gated behind login (LoginRequiredMixin redirects
    anonymous visitors to LOGIN_URL) since generating costs a trainer's
    free/paid access. ensure_csrf_cookie is still required here because this
    page has no {% csrf_token %} in its template (it's not a Django <form>
    POST) - without it, Django never sets the csrftoken cookie, so the
    page's own JS fetch() calls (AI-generate and the access check) have no
    token to send and every request is rejected with 403 CSRF cookie not
    set, before it ever reaches Groq or the access check.
    """
    template_name = "scheme_of_work_user.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update(_generator_header_context(self.request))
        return context


@method_decorator(ensure_csrf_cookie, name='dispatch')
class LessonPlanUserView(LoginRequiredMixin, TemplateView):
    """Same reasoning as SchemeOfWorkUserView above."""
    template_name = "lesson_plan_user.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update(_generator_header_context(self.request))
        return context


@method_decorator(ensure_csrf_cookie, name='dispatch')
class AssessmentPlanUserView(LoginRequiredMixin, TemplateView):
    """Same reasoning as SchemeOfWorkUserView above: no {% csrf_token %} form on
    this page, so ensure_csrf_cookie is required for the JS fetch() to work."""
    template_name = "assessment_plan_user.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update(_generator_header_context(self.request))
        return context


def login_required_json(view_func):
    """
    Same auth check as django.contrib.auth.decorators.login_required, but
    for JSON/fetch()-only endpoints: instead of redirecting an
    unauthenticated request to LOGIN_URL (which a fetch() caller can't
    usefully follow), it returns a 403 JSON body - the same shape used by
    check_generation_access below.
    """
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return JsonResponse({"error": "Login required."}, status=403)
        return view_func(request, *args, **kwargs)
    return _wrapped


# ---------------------------------------------------------------------------
# "Download as Word/Excel" export endpoints for the three generator pages
# above.
#
# These take the SAME data already rendered on screen - the browser reads
# it straight out of the DOM (see TMSGen.collectExportPayload in
# generator.js) and POSTs it here as JSON, rather than the server
# re-deriving the document from scratch. That keeps the export always in
# sync with whatever a trainer has typed/edited into the on-screen cells,
# without duplicating the Scheme/Session-Plan/Assessment generation logic
# a second time on the backend.
#
# Expected JSON body shape (all keys optional/best-effort):
#   {
#     "title": "...", "subtitle": "...",
#     "meta": [[label, value], ...],
#     "tables": [ [ {"section": bool, "cells": [
#         {"text": "...", "colspan": 1, "header": bool}, ... ]}, ... ], ... ],
#     "signoff": [{"label": "...", "name": "..."}, ...],
#     "footer": "..."
#   }
# ---------------------------------------------------------------------------
def _parse_export_payload(request):
    """Shared JSON parsing/validation for the three export endpoints below."""
    try:
        payload = json.loads(request.body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None, JsonResponse({"error": "Invalid JSON payload"}, status=400)

    if not isinstance(payload, dict):
        return None, JsonResponse({"error": "Invalid payload"}, status=400)

    tables = payload.get("tables")
    if not isinstance(tables, list) or not tables:
        return None, JsonResponse({"error": "No document content to export"}, status=400)

    return payload, None


def _content_disposition(filename):
    """
    Builds an RFC 5987/6266-compliant Content-Disposition header value for
    a downloaded file. Django's HttpResponse header assignment falls back
    to RFC 2047 ("=?utf-8?q?...?=") for non-ASCII header values, which is
    an *email* header encoding and isn't reliably understood by browsers
    for Content-Disposition - so titles containing an em dash, accents,
    etc. (common in module/trade names) could produce a mangled/garbled
    downloaded filename. This instead sends a plain ASCII fallback
    filename plus the proper filename*=UTF-8''<percent-encoded> form,
    which every modern browser prefers when present.
    """
    ascii_fallback = filename.encode("ascii", "ignore").decode("ascii").strip() or "Document"
    encoded = quote(filename)
    return f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded}"


def _docx_response(document, filename):
    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response["Content-Disposition"] = _content_disposition(filename)
    return response


def _build_export_docx(payload):
    """
    Builds a simple, clean Word document from a collectExportPayload()
    JSON structure: title heading, optional subtitle/meta, one Word table
    per captured HTML table (merging cells for any colspan, and bolding
    "section"/subhead rows), then the sign-off lines and footer text.
    """
    document = Document()

    for section in document.sections:
        section.top_margin = Pt(36)
        section.bottom_margin = Pt(36)
        section.left_margin = Pt(36)
        section.right_margin = Pt(36)

    title = (payload.get("title") or "Document").strip()
    heading = document.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = (payload.get("subtitle") or "").strip()
    if subtitle:
        p = document.add_paragraph(subtitle)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = payload.get("meta") or []
    if meta:
        meta_table = document.add_table(rows=0, cols=2)
        meta_table.style = "Table Grid"
        for label, value in meta:
            row = meta_table.add_row().cells
            row[0].paragraphs[0].add_run(str(label)).bold = True
            row[1].text = str(value)
        document.add_paragraph()

    for table_rows in payload.get("tables") or []:
        if not table_rows:
            continue
        n_cols = max((sum(c.get("colspan", 1) for c in r.get("cells", [])) for r in table_rows), default=1)
        n_cols = max(n_cols, 1)

        doc_table = document.add_table(rows=0, cols=n_cols)
        doc_table.style = "Table Grid"

        for row in table_rows:
            cells = row.get("cells", [])
            doc_row = doc_table.add_row().cells
            col_index = 0
            is_section = bool(row.get("section"))
            for cell in cells:
                if col_index >= n_cols:
                    break
                colspan = max(int(cell.get("colspan", 1) or 1), 1)
                end_index = min(col_index + colspan - 1, n_cols - 1)

                target = doc_row[col_index]
                if end_index > col_index:
                    target = target.merge(doc_row[end_index])

                run = target.paragraphs[0].add_run(str(cell.get("text", "")))
                if cell.get("header") or is_section:
                    run.bold = True

                col_index = end_index + 1
        document.add_paragraph()

    signoff = payload.get("signoff") or []
    if signoff:
        for line in signoff:
            label = (line.get("label") or "").strip()
            name = (line.get("name") or "").strip()
            if not (label or name):
                continue
            p = document.add_paragraph()
            p.add_run(f"{label} ").bold = True
            p.add_run(name)

    footer = (payload.get("footer") or "").strip()
    if footer:
        p = document.add_paragraph(footer)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.italic = True
            run.font.size = Pt(9)

    return document


@require_POST
@csrf_protect
@login_required_json
def export_scheme_of_work_docx(request):
    """"Download as Word" for the Scheme of Work generator page."""
    payload, error_response = _parse_export_payload(request)
    if error_response:
        return error_response
    document = _build_export_docx(payload)
    filename = f"{_export_filename(payload.get('title'), 'Scheme of Work')}.docx"
    return _docx_response(document, filename)


@require_POST
@csrf_protect
@login_required_json
def export_lesson_plan_docx(request):
    """"Download as Word" for the Lesson Plan / Session Plan generator page."""
    payload, error_response = _parse_export_payload(request)
    if error_response:
        return error_response
    document = _build_export_docx(payload)
    filename = f"{_export_filename(payload.get('title'), 'Lesson Plan')}.docx"
    return _docx_response(document, filename)


@require_POST
@csrf_protect
@login_required_json
def export_assessment_plan_xlsx(request):
    """"Download as Excel" for the Assessment Plan generator page."""
    payload, error_response = _parse_export_payload(request)
    if error_response:
        return error_response

    tables = payload.get("tables") or []
    table_rows = tables[0] if tables else []

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Assessment Plan"

    n_cols = max((sum(c.get("colspan", 1) for c in r.get("cells", [])) for r in table_rows), default=1)
    n_cols = max(n_cols, 1)

    current_row = 1

    title = (payload.get("title") or "Assessment Plan").strip()
    sheet.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=14)
    sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=n_cols)
    current_row += 1

    subtitle = (payload.get("subtitle") or "").strip()
    if subtitle:
        sheet.cell(row=current_row, column=1, value=subtitle)
        sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=n_cols)
        current_row += 1

    for label, value in (payload.get("meta") or []):
        sheet.cell(row=current_row, column=1, value=str(label)).font = Font(bold=True)
        sheet.cell(row=current_row, column=2, value=str(value))
        current_row += 1

    current_row += 1  # blank spacer row before the table

    for row in table_rows:
        col_index = 1
        for cell in row.get("cells", []):
            colspan = max(int(cell.get("colspan", 1) or 1), 1)
            xl_cell = sheet.cell(row=current_row, column=col_index, value=str(cell.get("text", "")))
            xl_cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.get("header") or row.get("section"):
                xl_cell.font = Font(bold=True)
            if colspan > 1:
                end_col = min(col_index + colspan - 1, n_cols)
                sheet.merge_cells(start_row=current_row, start_column=col_index, end_row=current_row, end_column=end_col)
                col_index = end_col + 1
            else:
                col_index += 1
        current_row += 1

    current_row += 1
    for line in (payload.get("signoff") or []):
        label = (line.get("label") or "").strip()
        name = (line.get("name") or "").strip()
        if not (label or name):
            continue
        sheet.cell(row=current_row, column=1, value=label).font = Font(bold=True)
        sheet.cell(row=current_row, column=2, value=name)
        current_row += 1

    for col in range(1, n_cols + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 22

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    filename = f"{_export_filename(payload.get('title'), 'Assessment Plan')}.xlsx"
    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = _content_disposition(filename)
    return response


def _export_filename(title, fallback):
    """Mirrors the frontend's TMSGen.filenameSafe so downloads get a sane name
    even if the client didn't already sanitise the title it sent."""
    clean = re.sub(r'[\\/:*?"<>|]', "-", str(title or "")).strip()
    clean = re.sub(r"\s+", " ", clean)
    if len(clean) > 120:
        clean = clean[:120].strip()
    return clean or fallback


# ---------------------------------------------------------------------------
# Trainer payment/access gating for the three generator pages above.
# ---------------------------------------------------------------------------
# Trainer-facing copy shown once the free generation has been used and the
# trainer has no active paid access. Manual/offline MTN MoMo flow - an
# Admin verifies the payment and flips `is_paid` on for the trainer from
# the Django admin (see TrainerAccessAdmin in admin.py).
GENERATION_ACCESS_BLOCKED_MESSAGE = (
    "Free Generation Used\n\n"
    "You have used your 1 free generation. To generate another Scheme, "
    "Lesson Plan, or Assessment Plan, please make a payment of 1000 FRW via "
    "MTN MoMo.\n\n"
    "Payment Number/Code: 0787306250/444322\n"
    "Amount: 1000 FRW\n\n"
    "After payment, please contact Jean Marie Vianney on WhatsApp. An "
    "administrator will verify your payment and activate your generation "
    "access.\n\n"
    "Note: Your access will be activated after payment verification."
)


@require_POST
@csrf_protect
def check_generation_access(request):
    """
    Called by the Scheme of Work / Lesson Plan / Assessment Plan generator
    pages before they render a document. Consumes the trainer's 1 free
    generation the first time it's used; after that, requires active paid
    access (granted manually by an Admin - see TrainerAccess/TrainerAccessAdmin).
    """
    if not request.user.is_authenticated:
        return JsonResponse({"error": "Login required."}, status=403)

    trainer = getattr(request.user, 'trainer_profile', None)
    if trainer is None:
        return JsonResponse(
            {"error": "This account has no linked trainer profile."}, status=403
        )

    access, _ = TrainerAccess.objects.get_or_create(trainer=trainer)

    if access.has_access:
        if not access.free_generation_used:
            access.free_generation_used = True
            access.save()
        return JsonResponse({"allowed": True})

    return JsonResponse({"allowed": False, "message": GENERATION_ACCESS_BLOCKED_MESSAGE})

# ---------------------------------------------------------------------------
# AI (Groq) endpoint used by the Scheme of Work generator page
# ---------------------------------------------------------------------------
@require_POST
@csrf_protect
@login_required_json
def generate_scheme_ai_content(request):
    """
    AI endpoint for the Scheme of Work generator page.

    SECURITY: previously had no auth check at all, so anyone with the URL
    could call it and burn the Groq API key/quota. Now gated behind
    @login_required_json (a 403-JSON flavour of login_required - see
    above - since this is a fetch()-only endpoint, not an HTML form).

    Receives the module/trade/level context plus the list of learning
    outcomes (with their indicative content) selected in the browser, and
    asks Groq to draft, for EACH outcome: learning activities, resources,
    learning place, and evidence of formative assessment. The Groq API key
    never leaves the server - the browser only ever talks to this endpoint.

    Field shapes returned per outcome (ALL plain single-line strings -
    no bullets, no numbering, no JSON arrays):
    - learning_activities: a SINGLE comma-separated line naming at most 3
      facilitation methodologies/techniques (e.g. "Demonstration, Group
      discussion, Practical exercise") - not a list of steps, and not a
      restatement of the outcome or its indicative content.
    - resources: a SINGLE comma-separated line of the tools/equipment/
      materials needed to deliver THAT row's outcome and indicative
      content specifically.
    - learning_place: a single string (e.g. "Workshop").
    - evidence: a SINGLE comma-separated line naming the assessment
      type(s) used to confirm formative learning for that outcome, drawn
      from: Written assessment, Oral assessment, Practical assessment,
      Assignment. Normally just ONE type per outcome.
    """
    # --------------------------------------------------
    # 1. Parse request safely
    # --------------------------------------------------
    try:
        payload = json.loads(request.body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return JsonResponse({"error": "Invalid JSON payload"}, status=400)

    outcomes = payload.get("outcomes")
    if not isinstance(outcomes, list) or not outcomes:
        return JsonResponse({"error": "Invalid or missing 'outcomes' field"}, status=400)

    # Trim/validate each outcome so we never forward garbage to the model,
    # and so we know exactly which outcome_ids we expect back.
    clean_outcomes = []
    expected_ids = []
    for o in outcomes:
        if not isinstance(o, dict) or "outcome_id" not in o:
            continue
        clean_outcomes.append({
            "outcome_id": o.get("outcome_id"),
            "learning_unit": str(o.get("learning_unit", ""))[:500],
            "learning_hours": o.get("learning_hours"),
            "indicative_content": [str(c)[:300] for c in (o.get("indicative_content") or [])][:20],
        })
        expected_ids.append(o.get("outcome_id"))

    if not clean_outcomes:
        return JsonResponse({"error": "No valid outcomes supplied"}, status=400)

    context_bits = []
    if payload.get("sector"):
        context_bits.append(f"Sector: {payload['sector']}")
    if payload.get("trade"):
        context_bits.append(f"Trade/Occupation: {payload['trade']}")
    if payload.get("level"):
        context_bits.append(f"Level: {payload['level']}")
    if payload.get("module_code") or payload.get("module_name"):
        context_bits.append(f"Module: {payload.get('module_code', '')} - {payload.get('module_name', '')}")
    if payload.get("hours_per_week"):
        context_bits.append(f"Hours per week: {payload['hours_per_week']}")
    if payload.get("weeks_per_term"):
        wpt = payload["weeks_per_term"]
        wpt_str = ", ".join(str(w) for w in wpt) if isinstance(wpt, list) else str(wpt)
        context_bits.append(f"Weeks per term: {wpt_str}")
    context_block = "\n".join(context_bits) or "No additional module context supplied."

    # --------------------------------------------------
    # 2. Build a structured, strict-JSON prompt
    # --------------------------------------------------
    system_prompt = (
        "You are an expert TVET (Technical and Vocational Education and Training) "
        "curriculum designer who writes practical, workshop-ready Schemes of Work. "
        "You always respond with a single valid JSON object and nothing else - "
        "no markdown fences, no commentary, no explanations."
    )

    user_prompt = f"""Using the module context below, generate practical training content for
EACH learning outcome listed in INPUT DATA.

MODULE CONTEXT:
{context_block}

For every outcome, produce FOUR fields, and every one of them must be a
SINGLE plain string on one line - NOT a JSON array, NOT bullet points, NOT
numbered steps. Where an item has more than one part, separate the parts
with ", " on that same line.

- "learning_activities": AT MOST 3 facilitation methodologies/teaching
  techniques the trainer would use to deliver that specific outcome (e.g.
  "Demonstration, Group discussion, Practical exercise" or "Role play, Q&A").
  Name ONLY the methodology/technique - do NOT describe, restate, or
  summarize the learning outcome or its indicative content, and do NOT
  write full sentences or steps. 1-3 items only.

- "resources": the specific tools, equipment, or materials required to
  deliver THIS outcome's own indicative content (e.g. "Digital multimeter,
  Safety goggles, Wiring diagram handout"). Base this strictly on what this
  row's outcome and indicative content need - do not reuse a generic list
  across rows, and do not pad with unrelated items.

- "learning_place": where the learning happens (e.g. "Classroom", "ICT Lab",
  "Workshop", "Field visit"). A single value, not a combination.

- "evidence": the assessment type(s) used to confirm formative learning
  happened for THIS outcome, chosen from: "Written assessment",
  "Oral assessment", "Practical assessment", "Assignment". Normally name
  ONLY ONE type per outcome - pick whichever best fits how this specific
  outcome and indicative content would actually be assessed. Only name a
  second type if genuinely both apply (e.g. "Practical assessment, Oral
  assessment").

Rules:
- ALL FOUR fields ("learning_activities", "resources", "learning_place",
  "evidence") must be plain strings with items separated by ", " where
  applicable - NEVER a JSON array, and NEVER containing bullet characters,
  dashes, or numbering.
- Ground every answer in the outcome's own text and its indicative content - do not invent unrelated topics.
- Use concise, real workshop/training language, not vague filler.
- Return EVERY outcome_id from the input, in the same order, exactly once.
- Respond with STRICT JSON ONLY, matching exactly this shape:
{{
  "results": [
    {{
      "outcome_id": <number>,
      "learning_activities": "<comma-separated string, max 3 items>",
      "resources": "<comma-separated string>",
      "learning_place": "<string>",
      "evidence": "<comma-separated string, usually 1 item>"
    }}
  ]
}}

INPUT DATA:
{json.dumps(clean_outcomes, indent=2)}
""".strip()

    # --------------------------------------------------
    # 3. Call Groq
    # --------------------------------------------------
    try:
        client = get_client()
    except RuntimeError as e:
        logger.error("Groq client not configured: %s", e)
        return JsonResponse({"error": str(e)}, status=500)

    model_name = getattr(settings, "GROQ_MODEL", "openai/gpt-oss-120b")

    try:
        completion = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.4,
            max_completion_tokens=4096,
            response_format={"type": "json_object"},
        )
    except RateLimitError:
        logger.warning("Groq rate limit hit")
        return JsonResponse(
            {"error": "The AI service is rate-limited right now. Please try again in a moment."},
            status=429,
        )
    except APIConnectionError as e:
        logger.error("Could not reach Groq: %s", e)
        return JsonResponse(
            {"error": "Could not reach the AI service. Check your internet connection and try again."},
            status=502,
        )
    except APIError as e:
        # Covers invalid/expired API keys, decommissioned models, bad requests, etc.
        logger.error("Groq API error (%s): %s", getattr(e, "status_code", "?"), e)
        message = str(e)
        if getattr(e, "status_code", None) == 401:
            message = "The Groq API key is invalid or missing. Check GROQ_API_KEY in your .env file."
        elif "decommissioned" in message.lower():
            message = (
                f"The model '{model_name}' is no longer available on Groq. "
                "Update GROQ_MODEL in your .env file to a current model "
                "(see https://console.groq.com/docs/models)."
            )
        return JsonResponse({"error": message}, status=502)
    except Exception as e:
        logger.exception("Unexpected error calling Groq")
        return JsonResponse({"error": "Internal server error", "detail": str(e)}, status=500)

    raw_content = (completion.choices[0].message.content or "").strip()

    # --------------------------------------------------
    # 4. Parse the AI response safely (strip stray ```json fences just in case)
    # --------------------------------------------------
    cleaned = raw_content
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        if cleaned.lower().startswith("json"):
            cleaned = cleaned[4:]
        cleaned = cleaned.strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        logger.error("AI returned invalid JSON: %s", raw_content)
        return JsonResponse(
            {"error": "AI returned invalid JSON", "raw_output": raw_content},
            status=502,
        )

    # --------------------------------------------------
    # 5. Validate structure
    # --------------------------------------------------
    results = data.get("results") if isinstance(data, dict) else None
    if not isinstance(results, list):
        logger.error("AI response missing 'results' list: %s", data)
        return JsonResponse(
            {"error": "Invalid AI response structure", "raw_output": data},
            status=502,
        )

    # Coerce outcome_id back to int where possible so the frontend's
    # lookup-by-id always matches, even if the model returned a numeric string.
    for r in results:
        if not isinstance(r, dict):
            continue

        if "outcome_id" in r:
            try:
                r["outcome_id"] = int(r["outcome_id"])
            except (TypeError, ValueError):
                pass

        r["learning_activities"] = _coerce_to_csv_line(r.get("learning_activities"), max_items=3)
        r["resources"] = _coerce_to_csv_line(r.get("resources"))
        r["evidence"] = _coerce_to_csv_line(r.get("evidence"), max_items=2)

        # learning_place stays a plain string
        if "learning_place" in r and not isinstance(r["learning_place"], str):
            r["learning_place"] = str(r["learning_place"])

    # --------------------------------------------------
    # 6. Return success
    # --------------------------------------------------
    return JsonResponse({"results": results})


# ---------------------------------------------------------------------------
# AI (Groq) endpoint used by the Lesson Plan (Session Plan) generator page
# ---------------------------------------------------------------------------
MIN_LESSON_PLAN_STEPS = 3
MAX_LESSON_PLAN_STEPS = 6


@require_POST
@csrf_protect
@login_required_json
def generate_lesson_plan_ai_content(request):
    """
    AI endpoint for the Lesson Plan (Session Plan) generator page.

    SECURITY: gated behind @login_required_json - see the note on
    generate_scheme_ai_content above.

    Receives the module/trade/level context plus the single learning
    outcome (with its indicative content) selected in the browser, and
    asks Groq to draft:
      - objectives: a JSON array of 2-5 clear, measurable, learner-focused
        session objectives, synthesised from the learning outcome, the
        Range and the indicative content - NOT a copy of the indicative
        content bullets. Server-side, padded (best-effort keyword check)
        so every step/Range sub-topic is covered by at least one
        objective.
      - facilitation_techniques: a SINGLE comma-separated line naming at
        most 3 facilitation methodologies/techniques (same shape/rules as
        the Scheme of Work's "learning_activities" field) - not a
        restatement of the outcome or its indicative content.
      - introduction: the session's opening, as trainer/learner activity
        bullet lists plus resources. The app always fixes the opening
        sequence itself - greet each other, then trainer takes/confirms
        attendance, then (unless this is a new module: Learning Outcome 1
        + Indicative Content 1.1) learners briefly recall the previous
        class - the AI only drafts the topic-specific activities after
        that, not generic "welcomes trainees" filler.
      - steps: the Development/Body of the session, as a JSON array of
        AT LEAST 3 steps (padded server-side if the model returns fewer),
        each with a short sub-topic "title" plus separate "trainer_activity"
        and "learner_activity" BULLET LISTS (JSON arrays of short plain
        phrases - not sentences glued together), and a "resources" line
        specific to that step.
      - reflection: a draft ANSWER (not the bare question) for each of the
        two standard post-session reflection prompts - "went_well" and
        "change_next_time" - grounded in this specific outcome/content, for
        the trainer to revise once the session has actually been taught.
      - range: a SHORT, comma/semicolon-separated line of sub-topics
        (very short phrases, not sentences) scoping this session. Derived
        server-side from the final "steps" titles (in order), so the
        Range and the Development/Body steps always match exactly.
      - conclusion_summary: a short recap for the Conclusion's Summary
        line that names the same sub-topics as the Range/steps, tying
        them back to the learning outcome.
      - evaluation: a short draft paragraph, written from the trainer's
        point of view AS IF the session has just been delivered, covering
        how the session went overall, how the class/trainees responded,
        and the trainer's own evaluation of it - for the trainer to revise
        once the session has actually been taught.
      - assessment: a JSON object choosing ONE of "Assessment" or
        "Assignment" (never both) with matching drafted content that
        directly tests the actual objectives above.
      - appendices: a short list of the specific handouts, checklists or
        materials attached to this session (plain text, one per line).
      - questions: a JSON array of EXACTLY 3 short assessment questions
        drawn from the Range, learning outcome, indicative content and the
        Development/Body steps of this session.

    The Groq API key never leaves the server - the browser only ever talks
    to this endpoint.
    """
    # --------------------------------------------------
    # 1. Parse request safely
    # --------------------------------------------------
    try:
        payload = json.loads(request.body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return JsonResponse({"error": "Invalid JSON payload"}, status=400)

    outcome_text = str(payload.get("outcome_text", ""))[:500]
    if not outcome_text.strip():
        return JsonResponse({"error": "Missing or empty 'outcome_text' field"}, status=400)

    indicative_content = [
        str(c)[:300] for c in (payload.get("indicative_content") or [])
    ][:20]

    # A "new module" session is Learning Outcome 1 + Indicative Content 1.1
    # (i.e. the very first outcome of the module, generating the session
    # that covers its first indicative content point). The browser works
    # this out from ordering/id and tells us, since that context lives in
    # the frontend's already-loaded outcomes/contents data.
    is_new_module = bool(payload.get("is_new_module"))

    context_bits = []
    if payload.get("sector"):
        context_bits.append(f"Sector: {payload['sector']}")
    if payload.get("trade"):
        context_bits.append(f"Trade/Occupation: {payload['trade']}")
    if payload.get("level"):
        context_bits.append(f"Level: {payload['level']}")
    if payload.get("module_code") or payload.get("module_name"):
        context_bits.append(f"Module: {payload.get('module_code', '')} - {payload.get('module_name', '')}")
    if payload.get("topic"):
        context_bits.append(f"Session topic: {payload['topic']}")
    if payload.get("range"):
        context_bits.append(f"Range (scope for this session): {payload['range']}")
    if payload.get("total_minutes"):
        context_bits.append(f"Total session duration: {payload['total_minutes']} minutes")
    context_block = "\n".join(context_bits) or "No additional session context supplied."

    # --------------------------------------------------
    # 2. Build a structured, strict-JSON prompt
    # --------------------------------------------------
    system_prompt = (
        "You are an expert TVET (Technical and Vocational Education and Training) "
        "curriculum designer who writes practical, workshop-ready session/lesson "
        "plans. Write every field in simple, clear English with common, everyday "
        "words - avoid complex, technical or academic vocabulary wherever a "
        "plainer word says the same thing. You always respond with a single "
        "valid JSON object and nothing else - no markdown fences, no "
        "commentary, no explanations."
    )

    user_prompt = f"""Using the session context below, draft the facilitation technique(s) and the
Development/Body of a competence-based training session plan for THIS
learning outcome.

SESSION CONTEXT:
{context_block}

LEARNING OUTCOME:
{outcome_text}

INDICATIVE CONTENT:
{json.dumps(indicative_content, indent=2)}

Produce a JSON object with exactly three top-level fields:

- "objectives": a JSON array of 2-5 clear, measurable, learner-focused
  session objectives - NOT a copy or rewording of the indicative content
  bullets, and NOT a restatement of the learning outcome itself. Each
  objective must:
  - Start with a single observable, measurable action verb (e.g.
    "Identify", "Describe", "Demonstrate", "Install", "Configure",
    "Troubleshoot", "Assemble", "Calculate", "Differentiate", "Apply") -
    never vague verbs like "understand", "know", or "learn".
  - Be phrased as a direct action statement starting with the verb - do
    NOT prefix it with "By the end of the session, trainees will be able
    to" or any similar lead-in (e.g. write "Configure a basic network
    switch", not "By the end of the session, trainees will be able to
    configure a basic network switch").
  - Be scoped to what THIS session actually covers - i.e. grounded in,
    but synthesised from, the learning outcome, the Range, and the
    indicative content above, not a line-by-line restatement of any one
    of them.
  - Be short (one sentence each) and independently achievable/assessable.
  Taken together, the objectives must cover every sub-topic you plan to
  put in "steps" below (i.e. every Range item) - do not draft an
  objective for a sub-topic you never teach in a step, and do not teach a
  step whose sub-topic no objective covers. Plan the steps/Range and the
  objectives together so the two line up.

- "facilitation_techniques": AT MOST 3 facilitation methodologies/teaching
  techniques the trainer would use across this session (e.g.
  "Demonstration, Group discussion, Practical exercise"). A SINGLE plain
  string on one line, items separated by ", " - NOT a JSON array, NOT
  bullet points. Name ONLY the methodology/technique, not a description.

- "introduction": a JSON object for the session's Introduction (before the
  Development/Body), with exactly three fields, specific to THIS topic and
  outcome - NOT generic filler like "welcomes trainees":
  - IMPORTANT: the app ALWAYS opens every session with the trainer and
    learners greeting each other, then the trainer taking/confirming
    attendance (roll call) - these two activities are added automatically,
    so do NOT draft greeting/welcoming or attendance/roll-call activities
    yourself, they would just be duplicated and discarded.
  - {"This session opens a brand-new module (Learning Outcome 1, Indicative Content 1.1), so do NOT draft any activity about recalling or discussing a previous class either - there isn't one yet." if is_new_module else "After attendance, the app also automatically adds an activity where learners briefly recall/discuss the previous class - so do NOT draft that yourself either, it would be duplicated."}
  - "trainer_activity": a JSON array of 2-3 SHORT, concrete, plain-text
    bullet phrases (imperative mood) for what the trainer does AFTER
    greeting/attendance{"" if is_new_module else "/recall"} to open THIS topic - e.g. review prior
    learning that actually leads into this outcome, and name the real
    sub-topics trainees are about to cover.
  - "learner_activity": a JSON array of 1-2 SHORT, concrete, plain-text
    bullet phrases for what trainees do during that same part of the
    introduction, in the same style.
  - "resources": the specific items needed for the introduction - a
    SINGLE comma-separated plain string, not an array.

- "steps": a JSON array covering the Development/Body of the session,
  broken into logical, progressive sub-topics grounded in the indicative
  content above. The step titles, IN ORDER, must double as the session's
  Range/sub-topic list, so keep them tightly matched to the Range you
  would draft for this same outcome and indicative content - no step
  should introduce a sub-topic that isn't part of the Range, and no Range
  sub-topic should be left without a matching step. Requirements:
  - MINIMUM {MIN_LESSON_PLAN_STEPS} steps, MAXIMUM {MAX_LESSON_PLAN_STEPS} steps, ALWAYS.
  - If fewer than {MIN_LESSON_PLAN_STEPS} indicative content items are given, split
    the content into progressive stages instead (e.g. explanation and
    demonstration, guided/hands-on practice, independent practice or
    verification) so the step count still meets the minimum.
  - If there are more indicative content items than {MAX_LESSON_PLAN_STEPS}, group
    closely related items together into a single step.
  - Each step is an object with FOUR fields:
    - "title": a short sub-topic label (a few words, NOT a full sentence,
      NOT prefixed with "Step" or a number - the app numbers steps itself,
      and this exact label is reused verbatim as one Range/sub-topic entry).
    - "trainer_activity": a JSON array of 2-4 SHORT, concrete, plain-text
      bullet phrases describing what the trainer does for this step
      (imperative mood, e.g. "Demonstrates the wiring sequence"). Each
      item is its own array entry - NOT one long sentence, NOT numbered,
      NOT starting with a dash or bullet character.
    - "learner_activity": a JSON array of 2-3 SHORT, concrete, plain-text
      bullet phrases describing what the trainee does for this step, in
      the same style as trainer_activity.
    - "resources": the specific tools, equipment, or materials this step
      needs - a SINGLE comma-separated plain string, not an array.

- "reflection": a JSON object with exactly two fields, each a short DRAFT
  ANSWER (1-2 plain sentences) written from the trainer's point of view,
  grounded in this specific outcome/content - NOT the question itself,
  and NOT generic filler. Since the session hasn't been delivered yet,
  phrase both as realistic, specific anticipations the trainer can revise
  after actually teaching it:
  - "went_well": what is likely to land well in this session, and why
    (e.g. which activity, demonstration, or hands-on moment).
  - "change_next_time": one concrete thing worth watching or adjusting
    (e.g. pacing, a step that may need more practice time, a resource
    that may be in short supply).

- "range": a SINGLE comma/semicolon-separated line of very short sub-topic
  phrases (a few words each, NOT full sentences) that scope this session.
  This MUST be the exact same sub-topics as the "steps" titles above, in
  the same order (the app will use your step titles as the authoritative
  Range if this differs) - matching the shape/rules of
  "facilitation_techniques" (one plain string, not an array, not bullet
  points).

- "conclusion_summary": a short draft paragraph (1-2 plain sentences) for
  the Conclusion's Summary line, written from the trainer's point of view
  AS IF the session has just been delivered. It must explicitly recap the
  SAME sub-topics covered in the "steps"/Range above and tie them back to
  the learning outcome - e.g. name the actual sub-topics taught, not a
  generic "recaps the key points" line.

- "evaluation": a short draft paragraph (2-3 plain sentences), written from
  the trainer's point of view AS IF the session has just been delivered -
  covering how the session went overall, how the class/trainees responded
  or performed, and how the trainer would rate/evaluate the session,
  grounded in this outcome/content. Write it as a real post-session
  evaluation note (e.g. "The session went well - most trainees could ...
  The class was engaged during ... Overall, the session met its
  objectives."), NOT a generic instruction telling the trainer what to
  write, and NOT phrased as a prediction about the future.

- "assessment": a JSON object with exactly two fields, deciding automatically
  between a formative check during the session ("Assessment") or a
  follow-up task after it ("Assignment") for THIS outcome - choose ONLY
  ONE, never both:
  - "type": either the exact string "Assessment" or the exact string
    "Assignment".
  - "content": 1-2 plain sentences describing that single chosen
    assessment or assignment, specific to this outcome/content. It must
    directly test whether trainees can now do what the "objectives"
    above say they can do - name or clearly reflect the actual action(s)
    from the objectives (e.g. if an objective is "Configure a basic
    network switch", the assessment must check that trainees can
    configure a switch, not something unrelated).

- "appendices": a short list (1-3 items) of the specific handouts,
  checklists, worksheets or diagrams that would be attached to THIS
  session, grounded in its content - a SINGLE string, one item per line
  (separate items with "\\n"), NOT a JSON array. If nothing specific
  applies, return an empty string.

- "questions": a JSON array of EXACTLY 3 short assessment questions for
  trainees, based on the Range, the learning outcome, the indicative
  content and the Development/Body steps of this session - plain
  questions a trainer could ask or set as a short quiz, not multi-part
  essay prompts.

Rules:
- Ground every objective and every step in the outcome's own text, the
  Range, and the indicative content - do not invent unrelated topics, and
  do not simply copy or lightly reword any indicative content line as an
  objective.
- Keep Range, steps, and the Conclusion summary tightly consistent with
  each other and with the Learning Outcome and Indicative Content: the
  Range is exactly the ordered list of step titles, and the Conclusion
  summary must name those same sub-topics - never introduce a sub-topic
  in one field that doesn't also appear in the others.
- Keep Objectives, steps/Range, and the Assessment/Assignment consistent
  with each other too: every step's sub-topic must be covered by at least
  one objective, and the Assessment/Assignment content must test the
  actual objectives - do not draft any of these three in isolation.
- Use concise, real workshop/training language, not vague filler, and keep
  every field in simple, everyday English - short words and short
  sentences over technical or academic phrasing.
- Respond with STRICT JSON ONLY, matching exactly this shape:
{{
  "objectives": ["<objective 1>", "<objective 2>"],
  "facilitation_techniques": "<comma-separated string, max 3 items>",
  "introduction": {{
    "trainer_activity": ["<bullet 1>", "<bullet 2>"],
    "learner_activity": ["<bullet 1>", "<bullet 2>"],
    "resources": "<comma-separated string>"
  }},
  "steps": [
    {{
      "title": "<short sub-topic label>",
      "trainer_activity": ["<bullet 1>", "<bullet 2>"],
      "learner_activity": ["<bullet 1>", "<bullet 2>"],
      "resources": "<comma-separated string>"
    }}
  ],
  "reflection": {{
    "went_well": "<1-2 sentence draft answer>",
    "change_next_time": "<1-2 sentence draft answer>"
  }},
  "range": "<comma-separated short sub-topics - same as step titles>",
  "conclusion_summary": "<1-2 sentence recap naming the same sub-topics>",
  "evaluation": "<2-3 sentence post-session evaluation, written as if just delivered>",
  "assessment": {{
    "type": "Assessment or Assignment - pick one",
    "content": "<1-2 sentence description testing the actual objectives>"
  }},
  "appendices": "<one item per line, or empty string>",
  "questions": ["<question 1>", "<question 2>", "<question 3>"]
}}
""".strip()

    # --------------------------------------------------
    # 3. Call Groq
    # --------------------------------------------------
    try:
        client = get_client()
    except RuntimeError as e:
        logger.error("Groq client not configured: %s", e)
        return JsonResponse({"error": str(e)}, status=500)

    model_name = getattr(settings, "GROQ_MODEL", "openai/gpt-oss-120b")

    try:
        completion = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.4,
            max_completion_tokens=4096,
            response_format={"type": "json_object"},
        )
    except RateLimitError:
        logger.warning("Groq rate limit hit")
        return JsonResponse(
            {"error": "The AI service is rate-limited right now. Please try again in a moment."},
            status=429,
        )
    except APIConnectionError as e:
        logger.error("Could not reach Groq: %s", e)
        return JsonResponse(
            {"error": "Could not reach the AI service. Check your internet connection and try again."},
            status=502,
        )
    except APIError as e:
        logger.error("Groq API error (%s): %s", getattr(e, "status_code", "?"), e)
        message = str(e)
        if getattr(e, "status_code", None) == 401:
            message = "The Groq API key is invalid or missing. Check GROQ_API_KEY in your .env file."
        elif "decommissioned" in message.lower():
            message = (
                f"The model '{model_name}' is no longer available on Groq. "
                "Update GROQ_MODEL in your .env file to a current model "
                "(see https://console.groq.com/docs/models)."
            )
        return JsonResponse({"error": message}, status=502)
    except Exception as e:
        logger.exception("Unexpected error calling Groq")
        return JsonResponse({"error": "Internal server error", "detail": str(e)}, status=500)

    raw_content = (completion.choices[0].message.content or "").strip()

    # --------------------------------------------------
    # 4. Parse the AI response safely (strip stray ```json fences just in case)
    # --------------------------------------------------
    cleaned = raw_content
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        if cleaned.lower().startswith("json"):
            cleaned = cleaned[4:]
        cleaned = cleaned.strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        logger.error("AI returned invalid JSON: %s", raw_content)
        return JsonResponse(
            {"error": "AI returned invalid JSON", "raw_output": raw_content},
            status=502,
        )

    if not isinstance(data, dict):
        logger.error("AI response is not a JSON object: %s", data)
        return JsonResponse(
            {"error": "Invalid AI response structure", "raw_output": data},
            status=502,
        )

    # --------------------------------------------------
    # 5. Validate & defensively coerce structure
    # --------------------------------------------------
    # Objectives must never fall back to a copy of the indicative content -
    # if the model returns nothing usable, fall back to a single generic
    # objective built from the outcome text only.
    short_outcome = outcome_text.strip()
    if ":" in short_outcome:
        short_outcome = short_outcome.split(":", 1)[1].strip()
    objectives = _coerce_to_list(
        data.get("objectives"),
        min_items=2,
        max_items=5,
        fallback=[
            f"Apply {short_outcome[:120]}." if short_outcome
            else "Perform the key steps covered in this session.",
            "Explain the key steps covered in this session in their own words.",
        ],
    )

    facilitation_techniques = _coerce_to_csv_line(
        data.get("facilitation_techniques"), max_items=3
    )

    # Introduction - structurally fixed so every session opens the same,
    # reliable way regardless of what the model returns:
    #   1. trainer and learners greet each other
    #   2. trainer takes/confirms attendance (roll call)
    #   3. UNLESS this is a new module (Learning Outcome 1 + Indicative
    #      Content 1.1), learners briefly recall/discuss the previous class
    # Everything after that is AI-drafted so it references the actual
    # topic instead of generic "welcomes trainees" filler. Greeting/
    # attendance/recall duplicates the model may still draft on its own
    # are stripped out below, since those slots are already covered.
    INTRO_FIXED_TRAINER = [
        "Greets the learners as they arrive",
        "Takes attendance and confirms the roll call",
    ]
    INTRO_FIXED_LEARNER = [
        "Greet the trainer and each other",
        "Respond to the roll call",
    ]
    INTRO_RECALL_TRAINER = "Invites learners to briefly recall what was covered in the previous class"
    INTRO_RECALL_LEARNER = "Briefly recall and share what was covered in the previous class"

    fixed_trainer = list(INTRO_FIXED_TRAINER)
    fixed_learner = list(INTRO_FIXED_LEARNER)
    if not is_new_module:
        fixed_trainer.append(INTRO_RECALL_TRAINER)
        fixed_learner.append(INTRO_RECALL_LEARNER)

    _STRUCTURAL_TRIGGERS = [
        "greet", "welcome",
        "attendance", "roll call", "roll-call",
        "previous class", "previous lesson", "prior class", "prior lesson",
        "last class", "last lesson", "recall",
    ]

    def _drop_structural(items):
        """Remove any AI-drafted bullet that duplicates the fixed
        greeting/attendance/recall activities enforced above."""
        kept = []
        for item in items:
            lowered = item.lower()
            if any(trigger in lowered for trigger in _STRUCTURAL_TRIGGERS):
                continue
            kept.append(item)
        return kept

    raw_introduction = data.get("introduction")
    raw_introduction = raw_introduction if isinstance(raw_introduction, dict) else {}

    topic_trainer = _coerce_to_list(
        _drop_structural(_split_items(raw_introduction.get("trainer_activity"))),
        min_items=1,
        max_items=3,
        fallback=[
            f"Reviews prior learning relevant to \"{short_outcome[:80]}\"",
            "Introduces the topic and session objectives",
            "Motivates learners for the session",
        ],
    )
    topic_learner = _coerce_to_list(
        _drop_structural(_split_items(raw_introduction.get("learner_activity"))),
        min_items=1,
        max_items=2,
        fallback=[
            "Listen and ask clarifying questions",
            "Share what they already know about the topic",
        ],
    )

    introduction = {
        "trainer_activity": fixed_trainer + topic_trainer,
        "learner_activity": fixed_learner + topic_learner,
        "resources": _coerce_to_csv_line(raw_introduction.get("resources")) or
        "Attendance register, whiteboard, projector",
    }

    raw_steps = data.get("steps")
    steps = []
    if isinstance(raw_steps, list):
        for s in raw_steps:
            if not isinstance(s, dict):
                continue
            steps.append({
                "title": str(s.get("title", "")).strip()[:120],
                "trainer_activity": _coerce_to_list(
                    s.get("trainer_activity"),
                    min_items=2,
                    max_items=4,
                    fallback=["Guides trainees through this step"],
                ),
                "learner_activity": _coerce_to_list(
                    s.get("learner_activity"),
                    min_items=2,
                    max_items=3,
                    fallback=["Participate in the activity under guidance"],
                ),
                "resources": _coerce_to_csv_line(s.get("resources")),
            })

    # Always guarantee at least MIN_LESSON_PLAN_STEPS steps, and never more
    # than MAX_LESSON_PLAN_STEPS, regardless of what the model returned.
    steps = _ensure_min_steps(steps, outcome_text, min_count=MIN_LESSON_PLAN_STEPS)
    steps = steps[:MAX_LESSON_PLAN_STEPS]

    raw_reflection = data.get("reflection")
    raw_reflection = raw_reflection if isinstance(raw_reflection, dict) else {}
    went_well = str(raw_reflection.get("went_well", "")).strip()[:400]
    change_next_time = str(raw_reflection.get("change_next_time", "")).strip()[:400]
    reflection = {
        "went_well": went_well or (
            f"Trainees engaging hands-on with \"{outcome_text.strip()[:80]}\" through the "
            "planned demonstration and guided practice."
        ),
        "change_next_time": change_next_time or (
            "Watch the pacing of the guided-practice step and be ready to extend it "
            "if trainees need more time to master the skill."
        ),
    }

    # Range/sub-topics - derived directly from the FINAL step titles (in
    # order) rather than trusting a separately-drafted "range" string, so
    # the Range field and the Development/Body steps always match exactly.
    step_titles = [s["title"] for s in steps if s.get("title")]
    range_text = "; ".join(step_titles)
    if not range_text:
        range_text = _coerce_to_csv_line(data.get("range"), max_items=6)
    if not range_text:
        range_text = ", ".join(indicative_content[:5])

    # Best-effort check: pad Objectives with any step/Range sub-topic that
    # no objective touches, so Objectives never silently drift away from
    # what the Development/Body steps actually cover.
    objectives = _ensure_objectives_cover_steps(objectives, step_titles, max_items=5)

    # Conclusion summary - recaps the SAME sub-topics as Range/steps above,
    # so the Conclusion clearly ties back to what was actually taught.
    conclusion_summary = str(data.get("conclusion_summary", "")).strip()[:400]
    if not conclusion_summary:
        if step_titles:
            topics_list = ", ".join(step_titles[:-1]) + (
                " and " + step_titles[-1] if len(step_titles) > 1 else step_titles[0]
            )
            conclusion_summary = (
                f"Recaps {topics_list}, checking these are understood against "
                f"\"{short_outcome[:100]}\"."
            )
        else:
            conclusion_summary = (
                f"Recaps the key points covered against \"{short_outcome[:100]}\" and "
                "checks overall understanding."
            )

    # Evaluation of the session - a post-session style note (how the
    # session went, how the class was, how the trainer rates it), never
    # left blank.
    evaluation = str(data.get("evaluation", "")).strip()[:600]
    if not evaluation:
        evaluation = (
            f"The session on \"{short_outcome[:100]}\" went well overall. The class stayed "
            "engaged through the guided and independent practice steps, and most trainees "
            "were able to follow along and take part. Overall, the session met its objectives."
        )

    # Assessment/Assignment - the model must pick exactly one; if it
    # doesn't, default to "Assessment" (an in-session formative check)
    # since that's always safe to run without extra trainee time.
    raw_assessment = data.get("assessment")
    raw_assessment = raw_assessment if isinstance(raw_assessment, dict) else {}
    assessment_type = str(raw_assessment.get("type", "")).strip().lower()
    if assessment_type not in ("assessment", "assignment"):
        assessment_type = "assessment"
    assessment_type = assessment_type.capitalize()
    primary_objective = objectives[0].strip().rstrip(".") if objectives else short_outcome[:100]
    assessment_content = str(raw_assessment.get("content", "")).strip()[:400]
    if not assessment_content:
        if assessment_type == "Assignment":
            assessment_content = (
                f"Give trainees a short follow-up task to {primary_objective[:1].lower()}"
                f"{primary_objective[1:]}, to complete and hand in after the session."
            )
        else:
            assessment_content = (
                f"Ask trainees a few quick questions and watch their work to check they can "
                f"{primary_objective[:1].lower()}{primary_objective[1:]}."
            )

    # Best-effort check: if the assessment content shares no keyword with
    # ANY objective, it likely drifted away from what the objectives
    # actually promise - append a short clause tying it back to the
    # primary objective rather than leaving it unrelated.
    if objectives:
        objective_words = set()
        for o in objectives:
            objective_words |= _keywords(o)
        if not (_keywords(assessment_content) & objective_words):
            assessment_content = (
                assessment_content.rstrip(".") +
                f", to check trainees can {primary_objective[:1].lower()}{primary_objective[1:]}."
            )[:400]

    assessment = {"type": assessment_type, "content": assessment_content}

    # Appendices - short plain list, empty string is a valid "nothing
    # applies" answer so we don't force filler text onto every session.
    appendices = str(data.get("appendices", "")).strip()[:500]

    # Questions - always exactly 3, padded with simple generic-but-usable
    # questions grounded in the outcome text if the model returns fewer.
    questions = _coerce_to_list(
        data.get("questions"),
        min_items=3,
        max_items=3,
        fallback=[
            f"What is {short_outcome[:80]}?" if short_outcome else "What did you learn in this session?",
            "List the main steps covered in this session.",
            "Why is this skill useful in real work?",
        ],
    )

    # --------------------------------------------------
    # 6. Return success
    # --------------------------------------------------
    return JsonResponse({
        "objectives": objectives,
        "facilitation_techniques": facilitation_techniques,
        "introduction": introduction,
        "steps": steps,
        "reflection": reflection,
        "range": range_text,
        "conclusion_summary": conclusion_summary,
        "evaluation": evaluation,
        "assessment": assessment,
        "appendices": appendices,
        "questions": questions,
    })


# ---------------------------------------------------------------------------
# AI (Groq) endpoint used by the Assessment Plan generator page
# ---------------------------------------------------------------------------
@require_POST
@csrf_protect
@login_required_json
def generate_assessment_plan_ai_content(request):
    """
    AI endpoint for the Assessment Plan generator page.

    SECURITY: gated behind @login_required_json - see the note on
    generate_scheme_ai_content above.

    Receives the module/trade/level context plus the list of learning
    outcomes (with their indicative content) selected in the browser, and
    asks Groq to draft, for EACH outcome: the best-fit assessment type,
    the resources needed to conduct that assessment, and one practical
    observation note. The Groq API key never leaves the server - the
    browser only ever talks to this endpoint.

    Field shapes returned per outcome (ALL plain single-line strings -
    no bullets, no numbering, no JSON arrays):
    - assessment_type: normally ONE assessment method, chosen from:
      Written assessment, Oral assessment, Practical assessment,
      Assignment.
    - resources: a SINGLE comma-separated line of the tools/equipment/
      materials/documents needed to conduct the assessment for that row's
      outcome specifically.
    - observation: one short practical note about administering that
      specific assessment.

    Deliberately NOT generated here (these stay blank/editable in the UI -
    they're administrative/scheduling facts only a human trainer knows):
    module_type, trainer, num_candidates, num_invigilators,
    assessment_date, place, publication_date.
    """
    # --------------------------------------------------
    # 1. Parse request safely
    # --------------------------------------------------
    try:
        payload = json.loads(request.body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return JsonResponse({"error": "Invalid JSON payload"}, status=400)

    outcomes = payload.get("outcomes")
    if not isinstance(outcomes, list) or not outcomes:
        return JsonResponse({"error": "Invalid or missing 'outcomes' field"}, status=400)

    # Trim/validate each outcome so we never forward garbage to the model,
    # and so we know exactly which outcome_ids we expect back.
    clean_outcomes = []
    expected_ids = []
    for o in outcomes:
        if not isinstance(o, dict) or "outcome_id" not in o:
            continue
        clean_outcomes.append({
            "outcome_id": o.get("outcome_id"),
            "learning_unit": str(o.get("learning_unit", ""))[:500],
            "indicative_content": [str(c)[:300] for c in (o.get("indicative_content") or [])][:20],
        })
        expected_ids.append(o.get("outcome_id"))

    if not clean_outcomes:
        return JsonResponse({"error": "No valid outcomes supplied"}, status=400)

    context_bits = []
    if payload.get("sector"):
        context_bits.append(f"Sector: {payload['sector']}")
    if payload.get("trade"):
        context_bits.append(f"Trade/Occupation: {payload['trade']}")
    if payload.get("level"):
        context_bits.append(f"Level: {payload['level']}")
    if payload.get("module_code") or payload.get("module_name"):
        context_bits.append(f"Module: {payload.get('module_code', '')} - {payload.get('module_name', '')}")
    context_block = "\n".join(context_bits) or "No additional module context supplied."

    # --------------------------------------------------
    # 2. Build a structured, strict-JSON prompt
    # --------------------------------------------------
    system_prompt = (
        "You are an expert TVET (Technical and Vocational Education and Training) "
        "assessment planner who writes practical, workshop-ready assessment plans. "
        "You always respond with a single valid JSON object and nothing else - no "
        "markdown fences, no commentary, no explanations."
    )

    user_prompt = f"""Using the module context below, generate assessment planning content for
EACH learning outcome listed in INPUT DATA.

MODULE CONTEXT:
{context_block}

For every outcome, produce THREE fields, and every one of them must be a
SINGLE plain string on one line - NOT a JSON array, NOT bullet points.

- "assessment_type": the single best-fit assessment method for confirming
  this outcome was achieved, chosen from: "Written assessment",
  "Oral assessment", "Practical assessment", "Assignment". Only name a
  second type if genuinely both apply (comma-separated).
- "resources": a SINGLE comma-separated line of the tools/equipment/
  materials/documents needed to CONDUCT the assessment for this specific
  outcome (e.g. "Assessment rubric, Practical work station, Marking guide").
  Base this strictly on what this outcome's own content needs.
- "observation": one short practical note about administering this specific
  assessment - e.g. a prerequisite, timing consideration, or special
  condition. Keep it to one plain sentence, no filler.
Rules:
- Ground every answer in the outcome's own text and its indicative content.
- Do NOT invent module type, trainer names, candidate counts, invigilator
  counts, dates, or venue - those are not part of your output.
- Return EVERY outcome_id from the input, in the same order, exactly once.
- Respond with STRICT JSON ONLY, matching exactly this shape:
{{
  "results": [
    {{
      "outcome_id": <number>,
      "assessment_type": "<string>",
      "resources": "<comma-separated string>",
      "observation": "<single sentence string>"
    }}
  ]
}}
INPUT DATA:
{json.dumps(clean_outcomes, indent=2)}
""".strip()

    # --------------------------------------------------
    # 3. Call Groq
    # --------------------------------------------------
    try:
        client = get_client()
    except RuntimeError as e:
        logger.error("Groq client not configured: %s", e)
        return JsonResponse({"error": str(e)}, status=500)

    model_name = getattr(settings, "GROQ_MODEL", "openai/gpt-oss-120b")

    try:
        completion = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.4,
            max_completion_tokens=4096,
            response_format={"type": "json_object"},
        )
    except RateLimitError:
        logger.warning("Groq rate limit hit")
        return JsonResponse(
            {"error": "The AI service is rate-limited right now. Please try again in a moment."},
            status=429,
        )
    except APIConnectionError as e:
        logger.error("Could not reach Groq: %s", e)
        return JsonResponse(
            {"error": "Could not reach the AI service. Check your internet connection and try again."},
            status=502,
        )
    except APIError as e:
        # Covers invalid/expired API keys, decommissioned models, bad requests, etc.
        logger.error("Groq API error (%s): %s", getattr(e, "status_code", "?"), e)
        message = str(e)
        if getattr(e, "status_code", None) == 401:
            message = "The Groq API key is invalid or missing. Check GROQ_API_KEY in your .env file."
        elif "decommissioned" in message.lower():
            message = (
                f"The model '{model_name}' is no longer available on Groq. "
                "Update GROQ_MODEL in your .env file to a current model "
                "(see https://console.groq.com/docs/models)."
            )
        return JsonResponse({"error": message}, status=502)
    except Exception as e:
        logger.exception("Unexpected error calling Groq")
        return JsonResponse({"error": "Internal server error", "detail": str(e)}, status=500)

    raw_content = (completion.choices[0].message.content or "").strip()

    # --------------------------------------------------
    # 4. Parse the AI response safely (strip stray ```json fences just in case)
    # --------------------------------------------------
    cleaned = raw_content
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        if cleaned.lower().startswith("json"):
            cleaned = cleaned[4:]
        cleaned = cleaned.strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        logger.error("AI returned invalid JSON: %s", raw_content)
        return JsonResponse(
            {"error": "AI returned invalid JSON", "raw_output": raw_content},
            status=502,
        )

    # --------------------------------------------------
    # 5. Validate structure
    # --------------------------------------------------
    results = data.get("results") if isinstance(data, dict) else None
    if not isinstance(results, list):
        logger.error("AI response missing 'results' list: %s", data)
        return JsonResponse(
            {"error": "Invalid AI response structure", "raw_output": data},
            status=502,
        )

    # Coerce outcome_id back to int where possible so the frontend's
    # lookup-by-id always matches, even if the model returned a numeric string.
    for r in results:
        if not isinstance(r, dict):
            continue

        if "outcome_id" in r:
            try:
                r["outcome_id"] = int(r["outcome_id"])
            except (TypeError, ValueError):
                pass

        r["assessment_type"] = _coerce_to_csv_line(r.get("assessment_type"), max_items=2)
        r["resources"] = _coerce_to_csv_line(r.get("resources"))

        # observation stays a plain single-sentence string
        if "observation" in r and not isinstance(r["observation"], str):
            r["observation"] = str(r["observation"])

    # --------------------------------------------------
    # 6. Return success
    # --------------------------------------------------
    return JsonResponse({"results": results})