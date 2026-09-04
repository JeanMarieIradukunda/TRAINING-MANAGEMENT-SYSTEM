import base64

from django import forms
from django.contrib.auth import get_user_model
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.hashers import make_password

from .models import (
    Logo, Sector, Trade, Level, TradeLevel, Trainer,
    Module, LearningOutcome, IndicativeContent, LessonPlan, AssessmentPlan,
    TrainerAccess, is_reusable_shadow_account, sync_trainer_login_account,
    TrainerLoginConflict,
)

MAX_LOGO_SIZE_BYTES = 2 * 1024 * 1024  # 2 MB


def encode_image_to_base64(uploaded_file):
    """
    Reads an uploaded image file and returns it as a data URI string
    (e.g. "data:image/png;base64,iVBORw0KG...") suitable for storing
    directly in Logo.image (a TextField) and for use as an <img src="">.
    """
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    encoded = base64.b64encode(raw_bytes).decode('utf-8')
    content_type = getattr(uploaded_file, 'content_type', None) or 'image/png'
    return f'data:{content_type};base64,{encoded}'


class StyledAuthenticationForm(AuthenticationForm):
    """Django's built-in login form, dressed up with Bootstrap classes."""

    username = forms.CharField(
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter your username',
            'autofocus': True,
        })
    )
    password = forms.CharField(
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter your password',
        })
    )


class BootstrapModelForm(forms.ModelForm):
    """Adds Bootstrap 5 form-control / form-select classes to every field."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field_name, field in self.fields.items():
            widget = field.widget
            if isinstance(widget, (forms.CheckboxInput,)):
                widget.attrs.setdefault('class', 'form-check-input')
            elif isinstance(widget, (forms.Select, forms.SelectMultiple)):
                widget.attrs.setdefault('class', 'form-select')
            else:
                widget.attrs.setdefault('class', 'form-control')
            widget.attrs.setdefault('placeholder', field.label)


class LogoForm(BootstrapModelForm):
    logo_file = forms.ImageField(
        label='Logo Image',
        required=False,
        widget=forms.ClearableFileInput(attrs={
            'class': 'form-control',
            'accept': 'image/png, image/jpeg, image/webp, image/svg+xml',
        }),
        help_text='PNG, JPG, WEBP or SVG, up to 2 MB. Stored securely and shown as a live preview.',
    )

    class Meta:
        model = Logo
        # NOTE: 'logo_file' is intentionally NOT listed here — it's not a
        # model field. It's declared above and Django automatically
        # includes it on the form regardless of Meta.fields.
        fields = ['name']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if not self.instance.pk:
            self.fields['logo_file'].required = True

    def clean_logo_file(self):
        logo_file = self.cleaned_data.get('logo_file')
        if logo_file and logo_file.size > MAX_LOGO_SIZE_BYTES:
            raise forms.ValidationError('Image must be smaller than 2 MB.')
        return logo_file

    def save(self, commit=True):
        logo = super().save(commit=False)
        logo_file = self.cleaned_data.get('logo_file')
        if logo_file:
            logo.image = encode_image_to_base64(logo_file)
        if commit:
            logo.save()
        return logo


class SectorForm(BootstrapModelForm):
    class Meta:
        model = Sector
        fields = ['sector_name']


class TradeForm(BootstrapModelForm):
    class Meta:
        model = Trade
        fields = ['sector', 'trade_name']


class LevelForm(BootstrapModelForm):
    class Meta:
        model = Level
        fields = ['class_level']


class TradeLevelForm(BootstrapModelForm):
    class Meta:
        model = TradeLevel
        fields = ['trade', 'level']


class TrainerForm(BootstrapModelForm):
    password = forms.CharField(
        label='Password',
        widget=forms.PasswordInput(attrs={'class': 'form-control', 'placeholder': 'Password'}),
        required=False,
        help_text='Leave blank to keep the current password when editing.'
    )

    class Meta:
        model = Trainer
        fields = ['fname', 'lname', 'username']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if not self.instance.pk:
            self.fields['password'].required = True

    def clean_username(self):
        """
        Trainer.username has its own unique constraint, but the shadow
        auth.User account created in sync_trainer_login_account() (see
        core/models.py) shares that same value and has its OWN unique
        constraint on auth_user. Deleting a Trainer doesn't remove its
        shadow User row (see TrainerDeleteView), so without this check,
        re-using a username that still belongs to an existing User -- e.g.
        a leftover shadow account, or a genuinely different login --
        surfaces as a raw IntegrityError from the database instead of a
        normal form error.

        A leftover shadow account (no staff/superuser rights, no usable
        password, not linked to another Trainer) is safe to silently
        reuse/relink -- see is_reusable_shadow_account(). Anything else
        sharing the username is rejected here with a clear message.
        """
        username = self.cleaned_data['username']
        User = get_user_model()
        existing_user = User.objects.filter(username=username).first()

        if existing_user is not None:
            same_account = self.instance.pk and self.instance.user_id == existing_user.pk
            if not same_account and not is_reusable_shadow_account(existing_user):
                raise forms.ValidationError(
                    'This username is already in use by another account. '
                    'Please choose a different username.'
                )
        return username

    def save(self, commit=True):
        trainer = super().save(commit=False)
        raw_password = self.cleaned_data.get('password')
        if raw_password:
            trainer.password_hash = make_password(raw_password)
        if commit:
            trainer.save()
            try:
                sync_trainer_login_account(trainer)
            except TrainerLoginConflict:
                # clean_username() above already guards against this in the
                # normal form-submit flow, so this only fires on a genuine
                # race (another save landed on the same username between
                # validation and this save). The Trainer record itself is
                # still saved correctly; just leave its login account as-is
                # rather than crashing - an admin can retry from the
                # Trainers list, which surfaces a "Create login" action for
                # any trainer without one.
                pass
        return trainer


class TrainerAccessForm(BootstrapModelForm):
    """
    Lets an Admin fine-tune a trainer's generator access straight from the
    Dashboard (Trainer Access screen), instead of needing the separate
    Django /admin site: flip Paid on/off, adjust the paid-until date by
    hand (e.g. to shorten/extend a window), jot a payment reference, or
    reset the one-time free generation so a trainer gets it again.
    """
    class Meta:
        model = TrainerAccess
        fields = ['is_paid', 'paid_until', 'payment_reference', 'free_generation_used']
        widgets = {
            'paid_until': forms.DateInput(attrs={'type': 'date'}),
        }
        labels = {
            'is_paid': 'Paid access active',
            'paid_until': 'Access expires on',
            'payment_reference': 'Payment reference / note',
            'free_generation_used': 'Free generation already used',
        }
        help_texts = {
            'paid_until': (
                'Leave blank while Paid access is checked to auto-set a fresh '
                '90-day window on save. Set a date yourself to grant a custom '
                'window, or to shorten/extend an existing one.'
            ),
        }


class ModuleForm(BootstrapModelForm):
    class Meta:
        model = Module
        # Trainer first: recording a Module starts with "who teaches
        # it", then the curriculum placement (trade/level), then the
        # module's own identity.
        fields = ['trainer', 'trade', 'level', 'mod_code', 'mod_name', 'learning_hours', 'term']


class LearningOutcomeForm(BootstrapModelForm):
    class Meta:
        model = LearningOutcome
        fields = ['module', 'outcome_text', 'learning_hours']
        widgets = {
            'outcome_text': forms.Textarea(attrs={'rows': 3}),
        }


class IndicativeContentForm(BootstrapModelForm):
    class Meta:
        model = IndicativeContent
        fields = ['outcome', 'indic_name']
        widgets = {
            'indic_name': forms.Textarea(attrs={'rows': 3}),
        }


class LessonPlanForm(BootstrapModelForm):
    class Meta:
        model = LessonPlan
        fields = ['module', 'trainer', 'title', 'week', 'lesson_date', 'objectives', 'activities', 'resources']
        widgets = {
            'lesson_date': forms.DateInput(attrs={'type': 'date'}),
            'objectives': forms.Textarea(attrs={'rows': 3}),
            'activities': forms.Textarea(attrs={'rows': 3}),
            'resources': forms.Textarea(attrs={'rows': 2}),
        }


class AssessmentPlanForm(BootstrapModelForm):
    class Meta:
        model = AssessmentPlan
        fields = [
            'module', 'learning_outcome', 'trainer', 'module_type', 'assessment_type',
            'num_candidates', 'num_invigilators', 'assessment_date', 'resources',
            'place', 'publication_date', 'observation',
        ]
        widgets = {
            'assessment_date': forms.DateInput(attrs={'type': 'date'}),
            'publication_date': forms.DateInput(attrs={'type': 'date'}),
            'resources': forms.Textarea(attrs={'rows': 2}),
            'observation': forms.Textarea(attrs={'rows': 3}),
        }


# ---------------------------------------------------------------------------
# Bulk-create workflow: "pick a parent once, then add several children"
#
# Learning Outcomes belong to a Module, and Indicative Contents belong to a
# Learning Outcome. Admins almost always add several of these at a time for
# the *same* parent (e.g. 5 outcomes for one module), so the single-record
# BootstrapModelForm above is repetitive for that job. These picker forms +
# formsets power a dedicated "add multiple" screen: choose the parent once,
# then fill in as many child rows as needed and save them all together.
# ---------------------------------------------------------------------------
class ModulePickerForm(forms.Form):
    """Step 1 of the Learning Outcome bulk-create screen: which module are
    these outcomes for. Rendered as a locked/read-only summary instead of
    this dropdown whenever the module was already chosen (e.g. the admin
    arrived via a module's "Add outcomes" quick link)."""
    module = forms.ModelChoiceField(
        queryset=Module.objects.select_related('trade', 'level').order_by('mod_code'),
        label='Module',
        empty_label='Select a module…',
        widget=forms.Select(attrs={'class': 'form-select'}),
    )


class OutcomePickerForm(forms.Form):
    """Step 1 of the Indicative Content bulk-create screen: which learning
    outcome these contents belong to."""
    outcome = forms.ModelChoiceField(
        queryset=LearningOutcome.objects.select_related('module').order_by('module__mod_code', 'id'),
        label='Learning Outcome',
        empty_label='Select a learning outcome…',
        widget=forms.Select(attrs={'class': 'form-select'}),
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['outcome'].label_from_instance = (
            lambda obj: f"{obj.module.mod_code} — {obj.outcome_text[:70]}"
        )


class LearningOutcomeQuickForm(BootstrapModelForm):
    """One row of the Learning Outcome bulk-create formset. Deliberately
    excludes 'module' — that's chosen once via ModulePickerForm and applied
    to every row on save."""
    class Meta:
        model = LearningOutcome
        fields = ['outcome_text', 'learning_hours']
        widgets = {
            'outcome_text': forms.Textarea(attrs={'rows': 2}),
        }


class IndicativeContentQuickForm(BootstrapModelForm):
    """One row of the Indicative Content bulk-create formset. Excludes
    'outcome' — that's chosen once via OutcomePickerForm."""
    class Meta:
        model = IndicativeContent
        fields = ['indic_name']
        widgets = {
            'indic_name': forms.Textarea(attrs={'rows': 2}),
        }


LearningOutcomeFormSet = forms.formset_factory(LearningOutcomeQuickForm, extra=3, can_delete=True)
IndicativeContentFormSet = forms.formset_factory(IndicativeContentQuickForm, extra=4, can_delete=True)
