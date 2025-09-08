# processor.py (backend logic)

import tempfile
import requests
from docx import Document
import openpyxl
import pdfplumber
import pandas as pd
# --- at top of file ---
import os
import tempfile
import requests
# (rest of your imports stay the same)

# NEW: pptx support
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ——— Azure OpenAI config ———
# Expect these to be set by the Streamlit frontend via secrets or environment variables
API_KEY = None  # to be set by frontend
API_ENDPOINT = None  # to be set by frontend


def configure(api_key: str, api_endpoint: str):
    global API_KEY, API_ENDPOINT
    API_KEY = api_key
    API_ENDPOINT = api_endpoint


def get_llm_response_azure(prompt: str, context: str) -> str:
    headers = {"Content-Type": "application/json", "api-key": API_KEY}
    system_msg = (
        "You are an expert on Transfer Pricing and financial analysis. "
        "Use the information in the following context to answer the user's question. "
        "Assign the greatest priority to the information that you gather from the financial analysis and the interview transcript. "
        "If asked something not covered in this data, you may search the web."
        "Ensure your analysis is CONCISE, SHARP, in paragraph form, and not long. Never use bullet points. "
        "DO NOT INCLUDE MARKDOWN FORMATTING OR # SIGNS. Keep it to 200-300 words, maintain a professional tone. "
        "Make sure to include direct sources and citations for the data you use for your decisions. Also include your reasoning for conclusions in brackets ()."
        "If something is from the transcript or financial statement, include that citation in brackets with a URL to the specific section. Likewise include a URL to the relevant website if the information you got was from searching the internet. "
        "You **may** consider the OECD guidelines below as helpful targets, but do NOT structure your response around them.\n\n"
    )
    messages = [
        {"role": "system", "content": system_msg},
        {"role": "user", "content":  context + prompt}
    ]
    resp = requests.post(API_ENDPOINT, headers=headers, json={"messages": messages})
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip()


# ---------------------
# Safe loader helpers
# ---------------------
def load_transcript(file) -> str:
    if not file:
        return ""
    try:
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""


def load_pdf(file) -> str:
    if not file:
        return ""
    pages, tables = [], []
    try:
        with pdfplumber.open(file) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                pages.append(f"--- Page {i} ---\n{text}")
                for table in page.extract_tables() or []:
                    # Be robust to ragged rows
                    try:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        tables.append(f"--- Page {i} table ---\n" + df.to_csv(index=False))
                    except Exception:
                        continue
        return "\n\n".join(pages + tables)
    except Exception:
        return ""


def load_guidelines(file) -> str:
    if not file:
        return ""
    try:
        # streamlit's UploadedFile supports .read(); ensure we don't exhaust twice
        content = file.read()
        try:
            return content.decode("utf-8").strip()
        except Exception:
            # fallback: latin-1 to avoid decode crash
            return content.decode("latin-1", errors="ignore").strip()
    except Exception:
        return ""


def load_and_annotate_replacements(excel_file, context: str):
    """
    TEMPLATE:
      - Col A: variable_name (info)
      - Col B: placeholder (token used inside prompts)
      - Col C: (unused here)
      - Col D: prompt (may include placeholders from Col B)
      - Col E: last year value (OLD; text to find in docs)
      - Col F: this year's value (NEW; replacement)
    Behavior:
      - Only use AI if F is blank AND D is non-empty.
      - Before AI call, substitute any known placeholders in D with their resolved NEW values.

    Returns:
      (replacements_dict, filled_excel_path)
        - replacements_dict: { old_value (Col E) -> resolved_new_value (Col F) }
        - filled_excel_path: absolute path to a temp 'variables_filled.xlsx' copy (or None on error)
    """
    if not excel_file:
        return {}, None
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception:
        return {}, None

    filled_excel_path = None
    try:
        ws = wb.active
        replacements = {}
        placeholder_to_new = {}

        for row in ws.iter_rows(min_row=2, max_col=6):
            cells = list(row) + [None] * max(0, 6 - len(row))
            cell_var = cells[0]  # A (unused for logic)
            cell_ph  = cells[1]  # B placeholder
            # cells[2] is C (ignored)
            cell_pr  = cells[3]  # D prompt
            cell_old = cells[4]  # E last year (OLD)
            cell_new = cells[5]  # F this year (NEW)

            placeholder = (str(cell_ph.value).strip() if cell_ph and cell_ph.value is not None else "")
            old_val     = (str(cell_old.value).strip() if cell_old and cell_old.value is not None else "")
            new_val_in  = (str(cell_new.value).strip() if cell_new and cell_new.value is not None else "")
            prompt_text = (str(cell_pr.value).strip() if cell_pr and cell_pr.value is not None else "")

            if not old_val:
                if placeholder and new_val_in:
                    placeholder_to_new[placeholder] = new_val_in
                continue

            if new_val_in:
                resolved_new = new_val_in
            else:
                if prompt_text:
                    prompt_rendered = prompt_text
                    if placeholder_to_new:
                        for ph_tok, ph_new in placeholder_to_new.items():
                            if ph_tok:
                                prompt_rendered = prompt_rendered.replace(ph_tok, ph_new)
                    try:
                        resolved_new = get_llm_response_azure(prompt_rendered, context)
                        resolved_new = str(resolved_new).strip() or old_val
                    except Exception:
                        resolved_new = old_val
                else:
                    resolved_new = old_val

                try:
                    ws.cell(row=cell_old.row, column=6, value=resolved_new)
                except Exception:
                    pass

            replacements[old_val] = resolved_new
            if placeholder:
                placeholder_to_new[placeholder] = resolved_new

        # Try saving back to the uploaded stream (best-effort, may be a no-op for some file-like objects)
        try:
            wb.save(excel_file)
        except Exception:
            pass

        # Always create a downloadable temp copy named variables_filled.xlsx
        try:
            tmpdir = tempfile.mkdtemp()
            filled_excel_path = os.path.join(tmpdir, "variables_filled.xlsx")
            wb.save(filled_excel_path)
        except Exception:
            filled_excel_path = None

        return replacements, filled_excel_path
    except Exception:
        return {}, None


# =========================
# DOCX (existing) helpers
# =========================
def collapse_runs(paragraph):
    from docx.oxml.ns import qn
    text = "".join(r.text for r in paragraph.runs)
    for r in reversed(paragraph.runs):
        r._element.getparent().remove(r._element)
    paragraph.add_run(text)


def replace_in_paragraph(p, replacements):
    collapse_runs(p)
    for run in p.runs:
        for ph, val in replacements.items():
            if ph in run.text:
                run.text = run.text.replace(ph, val)


def replace_placeholders_docx(doc: Document, replacements: dict):
    """Replace placeholders AFTER the first page break (mirrors your original behavior)."""
    from docx.oxml.ns import qn
    seen = False
    br_tag = qn('w:br')
    for p in doc.paragraphs:
        if not seen:
            for r in p.runs:
                for br in r._element.findall(br_tag):
                    if br.get(qn('w:type')) == 'page':
                        seen = True
                        break
                if seen:
                    break
            if not seen:
                continue
        replace_in_paragraph(p, replacements)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)
    for sec in doc.sections:
        if sec.header:
            for p in sec.header.paragraphs:
                replace_in_paragraph(p, replacements)
        if sec.footer:
            for p in sec.footer.paragraphs:
                replace_in_paragraph(p, replacements)


def replace_first_page_placeholders_docx(doc: Document, replacements: dict):
    """Replace placeholders on the first page only (up to first page break)."""
    from docx.oxml.ns import qn
    seen = False
    br_tag = qn("w:br")
    typ = qn("w:type")
    for p in doc.paragraphs:
        replace_in_paragraph(p, replacements)
        for r in p.runs:
            for child in r._element:
                if child.tag == br_tag and child.get(typ) == "page":
                    seen = True
                    break
            if seen:
                break
        if seen:
            break


# =========================
# PPTX (new) helpers
# =========================

def _pptx_replace_text_in_paragraph(paragraph, replacements: dict):
    """Collapse runs then perform in-place string replacements."""
    full = "".join(run.text for run in paragraph.runs) if getattr(paragraph, "runs", None) else getattr(paragraph, "text", "")
    for ph, val in replacements.items():
        if ph in full:
            full = full.replace(ph, val)
    paragraph.text = full


def _pptx_replace_in_text_frame(text_frame, replacements: dict):
    if not text_frame:
        return
    for para in text_frame.paragraphs:
        _pptx_replace_text_in_paragraph(para, replacements)


def _pptx_replace_in_table(table, replacements: dict):
    if not table:
        return
    for row in table.rows:
        for cell in row.cells:
            if getattr(cell, "text_frame", None):
                _pptx_replace_in_text_frame(cell.text_frame, replacements)


def _pptx_replace_in_shape(shape, replacements: dict):
    # Text boxes and placeholders
    if getattr(shape, "has_text_frame", False) and getattr(shape, "text_frame", None):
        _pptx_replace_in_text_frame(shape.text_frame, replacements)

    # Tables
    if getattr(shape, "has_table", False) and getattr(shape, "table", None):
        _pptx_replace_in_table(shape.table, replacements)

    # Charts (replace in chart title if present)
    # IMPORTANT: never touch shape.chart unless shape.has_chart is True,
    # because accessing .chart on non-chart shapes raises:
    #   ValueError: shape does not contain a chart
    if getattr(shape, "has_chart", False):
        try:
            chart = shape.chart
            if getattr(chart, "has_title", False):
                _pptx_replace_in_text_frame(chart.chart_title.text_frame, replacements)
        except Exception:
            # Be defensive; skip any chart we can't access
            pass

    # Grouped shapes — recurse
    if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
        for sub in shape.shapes:
            _pptx_replace_in_shape(sub, replacements)

def _prefill_last_year_from_prompts(excel_file, context: str) -> str | None:
    """
    Pre-step: fill Column E (last year value) using AI, based on prompt in Column D.
    Returns a temp file path to the updated workbook, or None on error.
    """
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception:
        return None

    try:
        ws = wb.active
        for row in ws.iter_rows(min_row=2, max_col=6):
            cells = list(row) + [None] * max(0, 6 - len(row))
            cell_pr  = cells[3]  # D: prompt
            cell_old = cells[4]  # E: last year value to (re)fill

            prompt_text = (str(cell_pr.value).strip() if cell_pr and cell_pr.value is not None else "")
            if not prompt_text:
                # No prompt, skip this row
                continue

            try:
                value = get_llm_response_azure(prompt_text, context)
                value = str(value).strip()
                if value:
                    ws.cell(row=cell_old.row, column=5, value=value)  # write to E
            except Exception:
                # On failure, leave existing E as-is
                pass

        # Save to a temp file and return its path
        tmpdir = tempfile.mkdtemp()
        out_path = os.path.join(tmpdir, "variables_prefilled_last_year.xlsx")
        wb.save(out_path)
        return out_path
    except Exception:
        return None


def replace_first_slide_placeholders_pptx(prs: Presentation, replacements: dict):
    """Replace placeholders on the first slide ONLY."""
    if not getattr(prs, "slides", None):
        return
    slide = prs.slides[0]
    for shp in slide.shapes:
        _pptx_replace_in_shape(shp, replacements)

    # Notes (if present)
    if getattr(slide, "has_notes_slide", False) and slide.has_notes_slide:
        notes = slide.notes_slide
        if hasattr(notes, "notes_text_frame") and notes.notes_text_frame is not None:
            _pptx_replace_in_text_frame(notes.notes_text_frame, replacements)


def replace_placeholders_pptx(prs: Presentation, replacements: dict, start_slide_index: int = 1):
    """Replace placeholders on all slides starting from start_slide_index (default: after first slide)."""
    for idx, slide in enumerate(prs.slides):
        if idx < start_slide_index:
            continue
        for shp in slide.shapes:
            _pptx_replace_in_shape(shp, replacements)

        # Notes (if present)
        if getattr(slide, "has_notes_slide", False) and slide.has_notes_slide:
            notes = slide.notes_slide
            if hasattr(notes, "notes_text_frame") and notes.notes_text_frame is not None:
                _pptx_replace_in_text_frame(notes.notes_text_frame, replacements)


# ---------------------
# Fallback DOCX builder
# ---------------------
def _build_fallback_docx(replacements: dict, context: str) -> str:
    """
    If no template is provided, produce a simple DOCX that lists
    the resolved replacements and includes a context snippet.
    """
    doc = Document()
    doc.add_heading("Transfer Pricing Output (Fallback)", level=1)

    if replacements:
        doc.add_heading("Resolved Placeholders", level=2)
        tbl = doc.add_table(rows=1, cols=2)
        hdr = tbl.rows[0].cells
        hdr[0].text = "Placeholder"
        hdr[1].text = "Value"
        for k, v in replacements.items():
            row = tbl.add_row().cells
            row[0].text = str(k)
            row[1].text = str(v)
    else:
        doc.add_paragraph("No replacements were generated (missing or empty Excel input).")

    if context:
        doc.add_heading("Context (truncated)", level=2)
        doc.add_paragraph(context[:4000])  # keep file small

    out = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(out.name)
    return out.name




def process_and_fill(files: dict, prefill_last_year: bool = False):
    """
    files: {...}
    prefill_last_year: if True, pre-fill Column E using AI prompts in Column D before replacement.
    """
    # Defensive dict access
    guidelines = files.get("guidelines") if files else None
    transcript = files.get("transcript") if files else None
    pdf = files.get("pdf") if files else None
    excel = files.get("excel") if files else None
    template = files.get("template") if files else None

    # Build context
    ctx = ""
    ctx += load_guidelines(guidelines)
    tr_text = load_transcript(transcript)
    if tr_text:
        ctx += ("\n\n" if ctx else "") + tr_text
    pdf_text = load_pdf(pdf)
    if pdf_text:
        ctx += ("\n\n" if ctx else "") + pdf_text

    # NEW: optionally pre-fill Column E based on prompts in Column D
    excel_for_processing = excel
    if prefill_last_year and excel:
        maybe_prefilled_path = _prefill_last_year_from_prompts(excel, ctx)
        if maybe_prefilled_path:
            excel_for_processing = maybe_prefilled_path  # use the updated workbook

    # Build replacements dict (and annotate Excel with generated values if present)
    replacements, filled_excel_path = load_and_annotate_replacements(excel_for_processing, ctx) if excel_for_processing else ({}, None)

    # Decide template type
    template_name = (getattr(template, "name", "") or "").lower()
    is_pptx = template_name.endswith(".pptx") if template else False
    is_docx = template_name.endswith(".docx") if template else False

    # If there's no template at all, produce a fallback DOCX
    if not template:
        doc_path = _build_fallback_docx(replacements, ctx)
        return (doc_path, filled_excel_path)

    if is_pptx:
        try:
            prs = Presentation(template)
        except Exception:
            doc_path = _build_fallback_docx(replacements, ctx)
            return (doc_path, filled_excel_path)
        replace_first_slide_placeholders_pptx(prs, replacements)
        replace_placeholders_pptx(prs, replacements, start_slide_index=1)
        out = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(out.name)
        return (out.name, filled_excel_path)

    elif is_docx:
        try:
            doc = Document(template)
        except Exception:
            doc_path = _build_fallback_docx(replacements, ctx)
            return (doc_path, filled_excel_path)
        replace_first_page_placeholders_docx(doc, replacements)
        replace_placeholders_docx(doc, replacements)
        out = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(out.name)
        return (out.name, filled_excel_path)

    else:
        doc_path = _build_fallback_docx(replacements, ctx)
        return (doc_path, filled_excel_path)
