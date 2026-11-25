"""
Exam Builder
============

A fourth generator page (alongside Scheme of Work, Lesson Plan and Assessment
Plan) where a trainer fills in input fields - header, instructions, sections,
question types and the questions themselves - and the exam is assembled in the
look of the school's "SCHOOL COMPREHENSIVE ASSESSMENT" Word template
(core/exam_assets/exam_template.docx): letterhead logo at the right of the
header block, bold SECTOR / TRADE / MODULE / title lines, a dashed rule, bold
"SECTION A: ... /60 - Marks" headings and "Q1." numbered questions.

How the pieces fit
------------------
* ExamBuilderView renders exam_builder_user.html with the SAME trainer-scoped
  curriculum snapshot the other generators use (_generator_header_context), so
  a trainer can only ever pick their own assigned modules, and the logos come
  from the system-stored Logo records.
* The browser turns everything the trainer typed into ONE flat "document
  model" (header lines, logos, ordered blocks). That single model drives both
  the live on-screen/printed preview and the Word download, so the two can
  never drift apart.
* export_exam_docx turns that model into a .docx built ON TOP OF the uploaded
  template (its theme fonts, page setup and styles), and archives a copy for
  the Dean of Studies like the other generators do.
* exam_ai_draft optionally drafts questions from the selected learning
  outcomes (Groq). If the AI is unreachable it degrades gracefully instead of
  failing, so the builder always works.
"""
import json
import logging
import os
import re
from io import BytesIO

from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import JsonResponse
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_protect, ensure_csrf_cookie
from django.views.decorators.http import require_POST
from django.views.generic import TemplateView

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Emu
from groq import APIConnectionError, APIError, RateLimitError

from . import views as core_views

logger = logging.getLogger(__name__)

EXAM_DOC_TYPE = "exam"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "exam_assets", "exam_template.docx")

# Page presets (inches). Margins copied from the uploaded template:
# 0.75" left/right, 1" top/bottom.
PAGE_SIZES = {
    "letter": (8.5, 11.0),
    "a4": (8.27, 11.69),
}
MARGIN_LR = 0.75
MARGIN_TB = 1.0

DOT = "\u2026"  # the "…" the template uses for blanks


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------
@method_decorator(ensure_csrf_cookie, name="dispatch")
class ExamBuilderView(LoginRequiredMixin, TemplateView):
    """
    Login-gated like the other generator pages. ensure_csrf_cookie is needed
    because the page talks to the server with fetch() (access check, AI draft,
    Word export) rather than through a Django <form>.
    """
    template_name = "exam_builder_user.html"

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context.update(core_views._generator_header_context(self.request))
        return context


# ---------------------------------------------------------------------------
# AI drafting of questions
# ---------------------------------------------------------------------------
VALID_TYPES = ("mcq", "tf", "fill", "match", "open")
MAX_QUESTIONS_PER_REQUEST = 40


def _clean(value, limit=600):
    return re.sub(r"\s+", " ", str(value or "")).strip()[:limit]


def _blank_question(kind):
    base = {"text": ""}
    if kind == "mcq":
        base.update({"options": ["", "", "", ""], "answer": 0})
    elif kind == "tf":
        base.update({"answer": True})
    elif kind == "match":
        base.update({"pairs": [{"l": "", "r": ""} for _ in range(4)]})
    elif kind == "open":
        base.update({"answerText": ""})
    return base


def _normalise_ai_question(kind, raw):
    """Coerces one question object returned by the model into the shape the
    browser editor uses. Returns None if it is unusable."""
    if not isinstance(raw, dict):
        return None
    if kind == "mcq":
        options = raw.get("options")
        if not isinstance(options, list):
            return None
        options = [_clean(o, 200) for o in options][:4]
        text = _clean(raw.get("question") or raw.get("text"))
        if not text or len(options) < 2:
            return None
        while len(options) < 4:
            options.append("")
        answer = raw.get("answer")
        idx = 0
        if isinstance(answer, int) and 0 <= answer < 4:
            idx = answer
        elif isinstance(answer, str) and answer.strip():
            letter = answer.strip()[0].upper()
            if letter in "ABCD":
                idx = "ABCD".index(letter)
        return {"text": text, "options": options, "answer": idx}
    if kind == "tf":
        text = _clean(raw.get("statement") or raw.get("question") or raw.get("text"))
        if not text:
            return None
        answer = raw.get("answer")
        if isinstance(answer, str):
            answer = answer.strip().lower() in ("true", "t", "yes")
        return {"text": text, "answer": bool(answer)}
    if kind == "fill":
        text = _clean(raw.get("sentence") or raw.get("question") or raw.get("text"))
        if not text or "[" not in text:
            return None
        return {"text": text}
    if kind == "match":
        pairs = raw.get("pairs")
        if not isinstance(pairs, list):
            return None
        clean_pairs = []
        for p in pairs[:8]:
            if isinstance(p, dict):
                left = _clean(p.get("left") or p.get("l"), 200)
                right = _clean(p.get("right") or p.get("r"), 200)
                if left and right:
                    clean_pairs.append({"l": left, "r": right})
        if len(clean_pairs) < 2:
            return None
        text = _clean(raw.get("instruction") or raw.get("question") or raw.get("text")) \
            or "Match each item in Column A with its correct answer in Column B."
        return {"text": text, "pairs": clean_pairs}
    if kind == "open":
        text = _clean(raw.get("question") or raw.get("text"))
        if not text:
            return None
        return {"text": text, "answerText": _clean(raw.get("answer") or raw.get("answerText"), 1200)}
    return None


def _fallback_questions(requests_, contents):
    """Used when the AI is unreachable. Open questions can be built honestly
    from the indicative content; every other type is returned blank for the
    trainer to complete (inventing facts offline would be worse than a blank)."""
    verbs = ["Explain", "Describe", "Outline", "Discuss", "State and explain"]
    out = {}
    pool = [c for c in contents if c]
    for req in requests_:
        kind, count = req["type"], req["count"]
        items = []
        for i in range(count):
            q = _blank_question(kind)
            if kind == "open" and pool:
                q["text"] = f"{verbs[i % len(verbs)]} {pool[i % len(pool)].rstrip('.')}."
            items.append(q)
        out[kind] = items
    return out


@require_POST
@csrf_protect
@core_views.login_required_json
def exam_ai_draft(request):
    """
    Drafts questions for ONE section from the learning outcomes the trainer
    selected. Request JSON:
      context   : {sector, trade, level, module_code, module_name}
      outcomes  : [{text, contents: [..]}]
      requests  : [{type: mcq|tf|fill|match|open, count: n}]
      difficulty: easy | medium | hard
    Response JSON: {questions: {<type>: [question, ...]}, source: ai|fallback, notice}
    Always answers 200 with usable (possibly blank) questions when the request
    itself is valid, so the builder never dead-ends on an AI outage.
    """
    try:
        payload = json.loads(request.body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return JsonResponse({"error": "Invalid JSON payload"}, status=400)

    raw_requests = payload.get("requests")
    requests_ = []
    total = 0
    if isinstance(raw_requests, list):
        for r in raw_requests:
            if not isinstance(r, dict) or r.get("type") not in VALID_TYPES:
                continue
            try:
                count = int(r.get("count") or 0)
            except (TypeError, ValueError):
                count = 0
            count = max(0, min(count, MAX_QUESTIONS_PER_REQUEST - total))
            if count:
                requests_.append({"type": r["type"], "count": count})
                total += count
    if not requests_:
        return JsonResponse({"error": "Choose at least one question type with a number of questions."}, status=400)

    outcomes = []
    contents_flat = []
    for o in (payload.get("outcomes") or [])[:12]:
        if not isinstance(o, dict):
            continue
        text = _clean(o.get("text"), 400)
        contents = [_clean(c, 250) for c in (o.get("contents") or [])][:20]
        if text:
            outcomes.append({"outcome": text, "indicative_content": contents})
            contents_flat.extend(contents)
    if not outcomes:
        return JsonResponse({"error": "Select at least one learning outcome first."}, status=400)

    ctx = payload.get("context") or {}
    context_bits = []
    for key, label in (("sector", "Sector"), ("trade", "Trade/Occupation"), ("level", "Level")):
        if ctx.get(key):
            context_bits.append(f"{label}: {_clean(ctx[key], 120)}")
    if ctx.get("module_code") or ctx.get("module_name"):
        context_bits.append(f"Module: {_clean(ctx.get('module_code'), 40)} - {_clean(ctx.get('module_name'), 150)}")
    context_block = "\n".join(context_bits) or "No additional context supplied."
    difficulty = payload.get("difficulty") if payload.get("difficulty") in ("easy", "medium", "hard") else "medium"

    spec_lines = []
    shape_lines = []
    shapes = {
        "mcq": '"mcq": [{"question": "...", "options": ["...", "...", "...", "..."], "answer": "B"}]',
        "tf": '"tf": [{"statement": "...", "answer": true}]',
        "fill": '"fill": [{"sentence": "A [router] connects different networks."}]',
        "match": '"match": [{"instruction": "...", "pairs": [{"left": "...", "right": "..."}]}]',
        "open": '"open": [{"question": "...", "answer": "short model answer / marking points"}]',
    }
    rules = {
        "mcq": "exactly 4 plausible options, exactly one correct; \"answer\" is the letter A, B, C or D; vary the position of the correct letter",
        "tf": "a clear statement that is definitely true or definitely false; roughly half true, half false",
        "fill": "one sentence with the missing key word(s) wrapped in [square brackets]; never bracket more than 2 words",
        "match": "each item has 4 to 5 pairs; every right-hand answer is short and unambiguous",
        "open": "a question needing a written answer of 2-6 sentences; \"answer\" lists the expected marking points",
    }
    for r in requests_:
        spec_lines.append(f"- {r['count']} \"{r['type']}\" question(s): {rules[r['type']]}")
        shape_lines.append(shapes[r["type"]])

    system_prompt = (
        "You are an experienced TVET (Technical and Vocational Education and Training) examiner in Rwanda who "
        "writes clear, fair, curriculum-aligned examination questions in simple English. You always respond with a "
        "single valid JSON object and nothing else - no markdown fences, no commentary."
    )
    user_prompt = f"""Write examination questions for the module below, based ONLY on the learning outcomes and
indicative content provided. Difficulty: {difficulty}. Test understanding of the content, do not invent topics
outside it, and never repeat a question.

MODULE CONTEXT:
{context_block}

LEARNING OUTCOMES AND INDICATIVE CONTENT:
{json.dumps(outcomes, indent=2)}

PRODUCE EXACTLY:
{chr(10).join(spec_lines)}

Spread the questions across the learning outcomes. Respond with STRICT JSON ONLY in exactly this shape
(include only the keys listed, each an array of exactly the requested size):
{{
  {("," + chr(10) + "  ").join(shape_lines)}
}}
"""

    notice = ""
    source = "ai"
    ai_data = None
    try:
        client = core_views.get_client()
        model_name = getattr(core_views.settings, "GROQ_MODEL", "openai/gpt-oss-120b")
        ai_data, _raw = core_views._groq_json_completion(
            client, model_name, system_prompt, user_prompt, 3800, temperature=0.5
        )
    except RateLimitError:
        notice = "The AI service is busy right now - blank questions were added for you to fill in."
    except (APIConnectionError, APIError, RuntimeError, ValueError) as exc:
        logger.warning("Exam AI draft unavailable: %s", exc)
        notice = "The AI service could not be reached - blank questions were added for you to fill in."
    except Exception:
        logger.exception("Unexpected error drafting exam questions")
        notice = "The AI service had a problem - blank questions were added for you to fill in."

    fallback = _fallback_questions(requests_, contents_flat)
    result = {}
    if not isinstance(ai_data, dict):
        source = "fallback"
        result = fallback
    else:
        for r in requests_:
            kind, count = r["type"], r["count"]
            good = []
            raw_list = ai_data.get(kind)
            if isinstance(raw_list, list):
                for item in raw_list:
                    q = _normalise_ai_question(kind, item)
                    if q:
                        good.append(q)
            good = good[:count]
            # pad with fallback/blank questions so the count the trainer asked for is honoured
            pad = fallback.get(kind, [])
            while len(good) < count:
                good.append(pad[len(good)] if len(good) < len(pad) else _blank_question(kind))
            result[kind] = good

    return JsonResponse({"questions": result, "source": source, "notice": notice})


# ---------------------------------------------------------------------------
# Word export  (built on top of the uploaded template)
# ---------------------------------------------------------------------------
def _open_base_document(page):
    """Opens the uploaded template (so its theme fonts/styles/page setup are
    inherited) and empties its body, ready to be rebuilt."""
    try:
        doc = Document(TEMPLATE_PATH)
    except Exception:
        logger.warning("Exam template missing/unreadable, using a plain document", exc_info=True)
        doc = Document()
    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    # the template's letterhead picture is no longer referenced - drop the orphan relationship
    for rid, rel in list(doc.part.rels.items()):
        if rel.reltype.endswith("/image"):
            try:
                doc.part.drop_rel(rid)
            except Exception:
                pass
    width, height = PAGE_SIZES.get(page, PAGE_SIZES["letter"])
    section = doc.sections[0]
    section.page_width = Inches(width)
    section.page_height = Inches(height)
    section.left_margin = section.right_margin = Inches(MARGIN_LR)
    section.top_margin = section.bottom_margin = Inches(MARGIN_TB)
    return doc, width - 2 * MARGIN_LR


def _set_cell_borders(table, color="808080", size=4):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def _no_borders(table):
    tbl = table._tbl
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tbl.tblPr.append(borders)


def _dashed_rule(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "dashed")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pbdr.append(bottom)
    pPr.append(pbdr)


def _keep_with_next(paragraph, on=True):
    paragraph.paragraph_format.keep_with_next = on


def _add_text(paragraph, text, bold=False, italic=False, size=None):
    """Adds text, turning '\n' into real line breaks."""
    parts = str(text or "").split("\n")
    for i, part in enumerate(parts):
        run = paragraph.add_run(part)
        run.bold = bold or None
        run.italic = italic or None
        if size:
            run.font.size = Pt(size)
        if i < len(parts) - 1:
            run.add_break()


def _picture_size(blob, max_h_in, max_w_in):
    """Height-limited size that also respects a maximum width."""
    try:
        from docx.image.image import Image as DocxImage
        img = DocxImage.from_blob(blob)
        w_px, h_px = img.px_width or 1, img.px_height or 1
    except Exception:
        return Inches(max_h_in), None
    h = max_h_in
    w = h * (w_px / h_px)
    if w > max_w_in:
        w = max_w_in
        h = w * (h_px / w_px)
    return Inches(h), Inches(w)


def _add_logos(container, logos, position, max_total_w_in, max_h_in=1.0):
    """Places the selected logos side by side inside `container` (a cell or the
    document) in one paragraph."""
    blobs = []
    for logo in logos:
        blob = core_views._resolve_logo_bytes((logo or {}).get("src"))
        if blob:
            blobs.append(blob)
    if not blobs:
        return False
    per_w = max_total_w_in / len(blobs)
    paragraph = container.add_paragraph() if hasattr(container, "add_paragraph") else container
    paragraph.alignment = {
        "right": WD_ALIGN_PARAGRAPH.RIGHT, "left": WD_ALIGN_PARAGRAPH.LEFT,
    }.get(position, WD_ALIGN_PARAGRAPH.CENTER)
    paragraph.paragraph_format.space_after = Pt(4)
    placed = 0
    for blob in blobs:
        try:
            height, width = _picture_size(blob, max_h_in, per_w)
            run = paragraph.add_run()
            run.add_picture(BytesIO(blob), height=height)
            if len(blobs) > 1:
                run.add_text("  ")
            placed += 1
        except Exception:
            logger.warning("Skipping a logo python-docx cannot embed", exc_info=True)
    return placed > 0


def _remove_first_empty_paragraph(cell):
    first = cell.paragraphs[0]
    if not first.text and not first.runs:
        first._element.getparent().remove(first._element)


def _build_header(doc, model, text_width):
    header = model.get("header") or {}
    lines = [str(l.get("text", "")) for l in (header.get("lines") or []) if isinstance(l, dict) and l.get("text")]
    logos = header.get("logos") or []
    position = header.get("logoPosition") if header.get("logoPosition") in ("left", "right", "center") else "right"

    have_logos = False
    if logos:
        probe = [core_views._resolve_logo_bytes((l or {}).get("src")) for l in logos]
        have_logos = any(probe)

    if have_logos and position in ("left", "right"):
        logo_w = min(3.2, text_width * 0.46) if len(logos) == 1 else min(3.6, text_width * 0.5)
        table = doc.add_table(rows=1, cols=2)
        _no_borders(table)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        text_cell_idx, logo_cell_idx = (0, 1) if position == "right" else (1, 0)
        widths = {text_cell_idx: text_width - logo_w, logo_cell_idx: logo_w}
        for idx, w in widths.items():
            table.columns[idx].width = Inches(w)
            table.rows[0].cells[idx].width = Inches(w)
        text_cell = table.rows[0].cells[text_cell_idx]
        logo_cell = table.rows[0].cells[logo_cell_idx]
        _remove_first_empty_paragraph(text_cell)
        for line in lines:
            p = text_cell.add_paragraph()
            _add_text(p, line, bold=True)
        if not lines:
            text_cell.add_paragraph()
        _remove_first_empty_paragraph(logo_cell)
        _add_logos(logo_cell, logos, position, logo_w - 0.15)
        doc.add_paragraph().paragraph_format.space_after = Pt(0)
    else:
        if have_logos:  # centred row above the header lines
            _add_logos(doc, logos, "center", text_width)
        for line in lines:
            p = doc.add_paragraph()
            _add_text(p, line, bold=True)


def _question_paragraph(doc, block, text_width):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(text_width), WD_TAB_ALIGNMENT.RIGHT)
    num = str(block.get("num") or "")
    run = p.add_run(num + " ")
    run.bold = True
    text = str(block.get("text") or "")
    _add_text(p, text)
    marks = block.get("marks")
    if marks not in (None, "", 0) and block.get("showMarks", True):
        run = p.add_run(f"\t[{marks} {'Mark' if str(marks) == '1' else 'Marks'}]")
        run.bold = True
    return p


def _render_question(doc, block, text_width, show_lines):
    kind = block.get("kind")
    if kind == "blank":
        p = doc.add_paragraph()
        r = p.add_run(f"{block.get('num', '')} {DOT * 6}?")
        r.bold = True
        return
    p = _question_paragraph(doc, block, text_width)
    _keep_with_next(p, kind in ("mcq", "match"))
    if kind == "mcq":
        options = block.get("options") or []
        for i, opt in enumerate(options):
            op = doc.add_paragraph()
            op.paragraph_format.left_indent = Inches(0.4)
            op.paragraph_format.space_after = Pt(2)
            _add_text(op, opt)
            _keep_with_next(op, i < len(options) - 1)
        doc.paragraphs[-1].paragraph_format.space_after = Pt(8)
    elif kind == "match":
        left = block.get("left") or []
        right = block.get("right") or []
        rows = max(len(left), len(right))
        table = doc.add_table(rows=rows + 1, cols=2)
        _set_cell_borders(table)
        table.autofit = False
        half = text_width * 0.48
        for ci in (0, 1):
            table.columns[ci].width = Inches(half)
        head = table.rows[0].cells
        for ci, label in enumerate(("COLUMN A", "COLUMN B")):
            head[ci].width = Inches(half)
            hp = head[ci].paragraphs[0]
            hp.paragraph_format.space_after = Pt(2)
            hr = hp.add_run(label)
            hr.bold = True
        for i in range(rows):
            cells = table.rows[i + 1].cells
            for ci, items in enumerate((left, right)):
                cells[ci].width = Inches(half)
                cp = cells[ci].paragraphs[0]
                cp.paragraph_format.space_after = Pt(2)
                if i < len(items):
                    _add_text(cp, items[i])
        # keep the whole table on one page: every row but the last stays with the next
        for ri, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            cant = OxmlElement("w:cantSplit")
            trPr.append(cant)
            if ri < len(table.rows) - 1:
                for cell in row.cells:
                    for par in cell.paragraphs:
                        par.paragraph_format.keep_with_next = True
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(4)
    elif kind == "open":
        count = int(block.get("lines") or 0) if show_lines else 0
        for _ in range(max(0, min(count, 20))):
            lp = doc.add_paragraph(DOT * 60)
            lp.paragraph_format.space_after = Pt(2)
            lp.paragraph_format.left_indent = Inches(0.0)
        if count:
            doc.paragraphs[-1].paragraph_format.space_after = Pt(8)


def _build_exam_docx(model):
    page = model.get("page") if model.get("page") in PAGE_SIZES else "letter"
    doc, text_width = _open_base_document(page)
    try:
        base_size = float(model.get("fontSize") or 11)
        if 9 <= base_size <= 14:
            doc.styles["Normal"].font.size = Pt(base_size)
    except (TypeError, ValueError):
        pass

    _build_header(doc, model, text_width)
    if model.get("rule", True):
        rule = doc.add_paragraph()
        rule.paragraph_format.space_after = Pt(10)
        _dashed_rule(rule)
        _keep_with_next(rule)

    show_lines = bool(model.get("answerLines", True))
    for block in (model.get("blocks") or []):
        if not isinstance(block, dict):
            continue
        t = block.get("t")
        if t == "instructions":
            if block.get("heading"):
                p = doc.add_paragraph()
                _add_text(p, block["heading"], bold=True)
                _keep_with_next(p)
            for i, item in enumerate(block.get("items") or []):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.first_line_indent = Inches(-0.3)
                p.paragraph_format.space_after = Pt(3)
                _add_text(p, f"{i + 1}.\t{item}")
                p.paragraph_format.tab_stops.add_tab_stop(Inches(0.3))
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        elif t == "section":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            _add_text(p, block.get("text"), bold=True)
            _keep_with_next(p)
        elif t == "note":
            p = doc.add_paragraph()
            _add_text(p, block.get("text"), italic=True)
            p.paragraph_format.space_after = Pt(6)
            _keep_with_next(p)
        elif t == "typehead":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            _add_text(p, block.get("text"), bold=True, italic=True)
            _keep_with_next(p)
        elif t == "wordbank":
            words = [str(w) for w in (block.get("words") or []) if str(w).strip()]
            if words:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                _add_text(p, "(" + "  /  ".join(words) + ")", bold=True)
                _keep_with_next(p)
        elif t == "q":
            _render_question(doc, block, text_width, show_lines)
        elif t == "pagebreak":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif t == "keyhead":
            p = doc.add_paragraph()
            _add_text(p, block.get("text"), bold=True)
            _keep_with_next(p)
        elif t == "keyrow":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
            r = p.add_run(f"{block.get('num', '')}\t")
            r.bold = True
            _add_text(p, block.get("text"))
            if block.get("marks") not in (None, ""):
                m = p.add_run(f"  [{block['marks']}]")
                m.bold = True
    return doc


@require_POST
@csrf_protect
@core_views.login_required_json
def export_exam_docx(request):
    """"Download as Word" for the Exam Builder. Body: the flat document model
    the browser also uses for its live preview (see exam_builder_user.html)."""
    try:
        model = json.loads(request.body.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return JsonResponse({"error": "Invalid JSON payload"}, status=400)
    if not isinstance(model, dict) or not isinstance(model.get("blocks"), list):
        return JsonResponse({"error": "No exam content to export"}, status=400)

    document = _build_exam_docx(model)
    title = (model.get("title") or "School Comprehensive Examination").strip()
    filename = f"{core_views._export_filename(title, 'School Comprehensive Examination')}.docx"
    data = core_views._docx_bytes(document)
    core_views._record_generated_document(
        request,
        {"title": title, "meta": model.get("meta") or []},
        EXAM_DOC_TYPE, filename, core_views.DOCX_CONTENT_TYPE, data,
    )
    return core_views._binary_response(data, core_views.DOCX_CONTENT_TYPE, filename)
