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
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import openpyxl
import tempfile
import os
import csv
from io import StringIO
import os, io, tempfile, openpyxl
from collections import OrderedDict
from docx import Document
from typing import Tuple, Dict

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


def load_and_annotate_replacements(excel_file, context: str) -> Tuple[Dict[str, str], str | None]:
    """
    New behavior for 4-column sheet:
      Columns (A..D): placeholder_name, old_value, prompt, new_value

      - If 'prompt' is not empty, query LLM via get_llm_response_azure(prompt, context)
        and write the response into 'new_value' (D).
      - Build replacements_dict as: { old_value -> new_value } for rows that have a non-empty new_value.

    Returns:
      (replacements_dict, filled_excel_path)
        - replacements_dict: { old_value -> new_value } (only rows with non-empty new_value)
        - filled_excel_path: absolute path to a temp 'variables_filled.xlsx' copy
                             (a valid path even if workbook was created blank)
    """
    # --- Try to load the workbook; if it fails, create a blank one with headers ---
    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
    except Exception:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "variables"
        ws.append(["placeholder_name", "old_value", "prompt", "new_value"])

    # --- Resolve column indices based on headers (row 1), fallback to A..D ---
    header_map = {}
    try:
        headers = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]
    except Exception:
        headers = []

    def _find_col(name: str, default_idx: int) -> int:
        try:
            idx = headers.index(name)
            return idx + 1  # openpyxl is 1-based
        except ValueError:
            return default_idx

    col_placeholder = _find_col("placeholder_name", 1)  # A
    col_old        = _find_col("old_value", 2)          # B
    col_prompt     = _find_col("prompt", 3)             # C
    col_new        = _find_col("new_value", 4)          # D

    replacements: Dict[str, str] = {}

    # --- Iterate rows, starting at row 2 ---
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        r = row[0].row  # current row index

        # Safely read cells
        def _get(col_idx: int) -> str:
            try:
                val = ws.cell(row=r, column=col_idx).value
                return str(val).strip() if val is not None else ""
            except Exception:
                return ""

        placeholder_name = _get(col_placeholder)
        old_value        = _get(col_old)
        prompt_text      = _get(col_prompt)
        new_value_curr   = _get(col_new)

        # If prompt present, query LLM and write into new_value
        if prompt_text:
            try:
                llm_out = get_llm_response_azure(prompt_text, context)
                llm_out = (llm_out or "").strip()
                # Write to the cell even if empty (clears previous content if any)
                ws.cell(row=r, column=col_new, value=llm_out)
                new_value_curr = llm_out
            except Exception:
                # On error, leave whatever was already in new_value
                pass

        # Build replacements only when new_value is non-empty
        if old_value and new_value_curr:
            replacements[old_value] = new_value_curr

    # --- Always create a downloadable temp copy named variables_filled.xlsx ---
    filled_excel_path = None
    try:
        tmpdir = tempfile.mkdtemp()
        filled_excel_path = os.path.join(tmpdir, "variables_filled.xlsx")
        wb.save(filled_excel_path)
    except Exception:
        filled_excel_path = None

    return replacements, filled_excel_path


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

def ask_variable_list(context: str) -> str:
    headers = {"Content-Type": "application/json", "api-key": API_KEY}
    CHANGE_PLAN_PROMPT = """
    You are a meticulous placeholder auditor for corporate documents (transfer pricing + finance context).

    GOAL
    From the provided company data ("CONTEXT"), output a clean, machine-readable table of everything that must be updated in the document/template.

    TABLE SCHEMA (CSV, no markdown, include header)
    placeholder_name,old_value,prompt,new_value

    DEFINITIONS
    - A placeholder can refer to: a word, number, date, name, paragraph, or a whole section.
    - "old_value" = the value currently implied by CONTEXT (e.g., last year’s value, outdated name/date, prior transaction figure).
    - "new_value" = the updated value you can infer from CONTEXT with high confidence.
    - "prompt" = ONLY used when the placeholder refers to a whole section that needs rewriting. In that case, write a clear instruction another LLM can use to generate the section (reference the concrete facts available in CONTEXT). For non-section placeholders, leave "prompt" blank.

    RULES
    1) Output ONLY CSV with the exact header:
    placeholder_name,old_value,prompt,new_value
    No commentary, no extra columns, no markdown.
    2) If you are NOT ≥80% confident about the correct new value, leave new_value empty.
    3) Keep values terse and precise (no vague phrases like “appears to be”).
    4) Prefer exact spans from CONTEXT for old_value and new_value when possible.
    5) Identify changes including (but not limited to):
    - Company/Group/Entity names; addresses; legal identifiers
    - Fiscal years, dates, periods (e.g., 2023 → 2024)
    - Financials (revenue, EBIT, margins, headcount, transaction values)
    - Functions, assets, risks; intercompany transactions; counterparties
    - Benchmarks/comparables references
    - Governance/personnel changes (roles, titles)
    - Jurisdictions, regulations, citations that changed
    - Any section that must be rewritten due to updated facts
    6) For SECTION-LEVEL placeholders, set placeholder_name to a clear label (e.g., SECTION:Functional Analysis),
    old_value to a short excerpt/summary of the current stance, prompt to a concrete writing instruction,
    and new_value empty (the other LLM will fill it).
    7) Limit each cell to essential text. Avoid line breaks in cells. Escape commas with quotes if needed.

    CONTEXT
    <<<CONTEXT_START
    {context}
    CONTEXT_END>>>

    TASK
    Produce the CSV now.
    """

    system_msg = (
        "You are a precise placeholder auditor that emits ONLY compact CSV suitable for direct ingestion. "
        "Do not use markdown, bullets, or explanations. Keep outputs concise and never vague."
    )

    # If you keep `prompt` external, pass CHANGE_PLAN_PROMPT and format {context} beforehand.
    # Here we always append the contextualized prompt after the system message.
    messages = [
        {"role": "system", "content": system_msg},
        {
            "role": "user",
            "content": CHANGE_PLAN_PROMPT.format(context=context)  # expects {context} in the prompt
        },
    ]
    resp = requests.post(API_ENDPOINT, headers=headers, json={"messages": messages})
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip()




def _prefill_last_year_from_prompts(excel_file, context: str) -> str | None:
    """
    Creates/updates an Excel workbook of variables based on ask_variable_list(context).
    If excel_file cannot be loaded, a new workbook is created with columns:
    placeholder_name, old_value, prompt, new_value.
    
    The output of ask_variable_list(context) (CSV text) is parsed and rows are
    added to this workbook. Returns a temp file path to the updated workbook,
    or None on error.
    """
    try:
        # Try to load the Excel, otherwise create a new workbook with required columns
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
        except Exception:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "variables"
            ws.append(["placeholder_name", "old_value", "prompt", "new_value"])

        # Get CSV from the LLM
        csv_output = ask_variable_list(context)
        if not csv_output:
            return None

        # Parse CSV safely
        reader = csv.DictReader(StringIO(csv_output))
        required_cols = ["placeholder_name", "old_value", "prompt", "new_value"]

        for row in reader:
            # Ensure all required columns exist in parsed row
            data = [row.get(col, "").strip() for col in required_cols]
            ws.append(data)

        # Save to a temp file and return its path
        tmpdir = tempfile.mkdtemp()
        out_path = os.path.join(tmpdir, "variables_prefilled.xlsx")
        wb.save(out_path)
        return out_path

    except Exception as e:
        print(f"Error in _prefill_last_year_from_prompts: {e}")
        return None

def generate_doc_from_excel_map(file_map, context: str = ""):
    """
    Inputs:
      file_map = {"excel": <path|file-like>, "template": <path|file-like>, ...}
    Behavior:
      - Read Excel with columns: placeholder_name, old_value, prompt, new_value
      - Build replacements = {old_value -> new_value} (non-empty on both)
      - Apply replacements to DOCX template; else build fallback
      - Return (doc_path, filled_excel_path)
    """

    # --- load workbook or create blank with headers ---
    def _load_wb(x):
        try:
            if isinstance(x, str) and os.path.exists(x):  # path
                return openpyxl.load_workbook(x)
            if hasattr(x, "read"):                        # file-like
                b = x.read()
                try: x.seek(0)
                except Exception: pass
                return openpyxl.load_workbook(io.BytesIO(b))
        except Exception:
            pass
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["placeholder_name", "old_value", "prompt", "new_value"])
        return wb

    wb = _load_wb(file_map.get("excel"))
    ws = wb.active

    # --- header map (fallback to A..D) ---
    try:
        hdr = [str(c.value).strip().lower() if c.value else "" for c in ws[1]]
    except Exception:
        hdr = []
    def col(name, default_idx): 
        return (hdr.index(name)+1) if name in hdr else default_idx
    c_old, c_new = col("old_value", 2), col("new_value", 4)

    # --- build replacements (keep order of appearance) ---
    repl = OrderedDict()
    for r in range(2, ws.max_row+1):
        old_s = (str(ws.cell(r, c_old).value).strip() if ws.cell(r, c_old).value else "")
        new_s = (str(ws.cell(r, c_new).value).strip() if ws.cell(r, c_new).value else "")
        if old_s and new_s and old_s != new_s:
            repl[old_s] = new_s

    # --- save a temp copy of the excel ---
    tmpdir = tempfile.mkdtemp()
    filled_excel_path = os.path.join(tmpdir, "variables_filled.xlsx")
    try: wb.save(filled_excel_path)
    except Exception: filled_excel_path = None

    # --- try to render DOCX with replacements; fallback if needed ---
    template = file_map.get("template")
    try:
        doc = Document(template) if template else None
    except Exception:
        doc = None

    if doc:
        if 'replace_first_page_placeholders_docx' in globals(): replace_first_page_placeholders_docx(doc, repl)
        if 'replace_placeholders_docx' in globals():            replace_placeholders_docx(doc, repl)
        out = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(out.name)
        return out.name, filled_excel_path

    if '_build_fallback_docx' in globals():
        return _build_fallback_docx(repl, context), filled_excel_path

    # ultra-simple fallback
    doc = Document(); doc.add_heading("Generated Document (Fallback)", 1)
    doc.add_paragraph("Template unavailable. Applied replacements:")
    for k, v in repl.items(): doc.add_paragraph(f"{k} → {v}")
    out = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(out.name)
    return out.name, filled_excel_path

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

def load_template(template_file) -> str:
    """
    Extract plain text from a DOCX template file.
    Returns a single string with paragraphs separated by newlines.
    """
    if not template_file:
        return ""

    try:
        doc = Document(template_file)
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(text)
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"Error loading template: {e}")
        return ""


def find_relevant_variables(files: dict):
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
    
    template_text = load_template(template) if template else None
    if template_text:
        ctx += ("\n\n" if ctx else "") + template_text
    
    # NEW: optionally pre-fill Column E based on prompts in Column D
    excel_for_processing = excel
    maybe_prefilled_path = _prefill_last_year_from_prompts(excel, ctx)
    if maybe_prefilled_path:
        excel_for_processing = maybe_prefilled_path  # use the updated workbook
    # Build replacements dict (and annotate Excel with generated values if present)
    replacements, filled_excel_path = load_and_annotate_replacements(excel_for_processing, ctx) if excel_for_processing else ({}, None)
    doc_path = _build_fallback_docx(replacements, ctx)
    return (doc_path, filled_excel_path)

def fill_section_values(files):
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
    
    template_text = load_template(template) if template else None
    if template_text:
        ctx += ("\n\n" if ctx else "") + template_text
    
    # NEW: optionally pre-fill Column E based on prompts in Column D
    excel_for_processing = excel
    print(excel_for_processing)
    (rep, filled_path) = load_and_annotate_replacements(excel_for_processing, ctx)
    return filled_path



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
    
    template_text = load_template(template) if template else None
    if template_text:
        ctx += ("\n\n" if ctx else "") + template_text
    print("m: ", template)
    
    # NEW: optionally pre-fill Column E based on prompts in Column D
    excel_for_processing = excel
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
    # if not template:
    #     doc_path = _build_fallback_docx(replacements, ctx)
    #     return (doc_path, filled_excel_path)
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
