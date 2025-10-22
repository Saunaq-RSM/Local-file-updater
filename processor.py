# processor.py (backend logic) — cleaned and reordered
#
# Purpose:
# - Keep functions that streamlit_app.py uses at the top, in this order:
#     1) configure
#     2) find_relevant_variables
#     3) fill_section_values
#     4) generate_doc_from_excel_map
# - Keep all helpers required by the above functions, after the top-level functions.
# - Remove unused imports and unused helpers/constants.
#
# Note: Behavior is unchanged except for removal of unused definitions and reordering.

import requests
from docx import Document
import pdfplumber
import pandas as pd

# --- at top of file ---
import os
import io
import time
import tempfile
import openpyxl
from collections import OrderedDict
from typing import Tuple, Dict, List
from pydantic import BaseModel, TypeAdapter
from docx.oxml.ns import qn
import re

# ——— Azure OpenAI config ———
# Expect these to be set by the Streamlit frontend via secrets or environment variables
API_KEY = None  # to be set by frontend
API_ENDPOINT = None  # to be set by frontend

class Row(BaseModel):
    placeholder_name: str
    old_value: str
    prompt: str
    new_value: str

    # Pydantic v2 config
    model_config = {
        "extra": "forbid",   # reject unknown fields
        "str_strip_whitespace": True
    }

# v2 parsing helpers
_row_list_adapter = TypeAdapter(List[Row])

def parse_rows_json(s: str) -> List[Row]:
    # Accepts a JSON string; raises ValidationError on mismatch
    return _row_list_adapter.validate_json(s)


# -------------------------
# Top-level API (front-end)
# Order matches usage in streamlit_app.py
# -------------------------

def configure(api_key: str, api_endpoint: str):
    """
    Called once by the Streamlit frontend to set API credentials/endpoints.
    """
    global API_KEY, API_ENDPOINT
    API_KEY = api_key
    API_ENDPOINT = api_endpoint


def find_relevant_variables(files: dict):
    """
    Used by Step 1 (Fill & preview variables).
    - Builds context from optional files
    - Optionally pre-fills prompts into an excel
    - Annotates the excel (fills new_value for prompt rows)
    - Returns (doc_path, filled_excel_path)
      doc_path is a fallback doc summarizing replacements (used by frontend)
    """
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

    # Optionally prefill prompts with ask_variable_list
    excel_for_processing = excel
    maybe_prefilled_path = _prefill_last_year_from_prompts(excel, ctx)
    if maybe_prefilled_path:
        excel_for_processing = maybe_prefilled_path  # use the updated workbook

    # Build replacements dict (and annotate Excel with generated values if present)
    replacements, filled_excel_path = load_and_annotate_replacements(excel_for_processing, ctx) if excel_for_processing else ({}, None)

    # Provide a fallback doc that lists replacements + context; frontend expects a doc path
    doc_path = _build_fallback_docx(replacements, ctx)
    return (doc_path, filled_excel_path)


def fill_section_values(files):
    """
    Used by Step 2 (Fill section values).
    - Rebuilds context
    - Calls load_and_annotate_replacements to fill prompt-driven new_value cells
    - Returns path to annotated excel (variables_filled.xlsx) for frontend to load
    """
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

    # Fill prompt-driven cells and return path
    excel_for_processing = excel
    (rep, filled_path) = load_and_annotate_replacements(excel_for_processing, ctx)
    return filled_path


def generate_doc_from_excel_map(file_map, context: str = ""):
    """
    Used by Step 3 (Generate final document).
    - Reads annotated excel and builds replacements {old_value -> new_value}
    - Saves a temp copy variables_filled.xlsx and returns its path as well
    - Attempts to load the provided template and apply replacements; if not possible, builds fallback doc
    - Returns (doc_path, filled_excel_path)
    """
    # --- load workbook or create blank with headers ---
    def _load_wb(x):
        try:
            if isinstance(x, str) and os.path.exists(x):  # path
                return openpyxl.load_workbook(x)
            if hasattr(x, "read"):                        # file-like
                b = x.read()
                try:
                    x.seek(0)
                except Exception:
                    pass
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
        if 'replace_first_page_placeholders_docx' in globals():
            replace_first_page_placeholders_docx(doc, repl)
        if 'replace_placeholders_docx' in globals():
            replace_placeholders_docx(doc, repl)
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


# -------------------------
# Helper functions (used by the top-level functions above)
# -------------------------

def ask_variable_list(context: str, wait_seconds: float = 2.0, max_retries: int = 6):
    """
    Keep calling the LLM until it returns valid JSON (parsed as List[Row]).
    Stops after max_retries and raises an exception if still invalid.
    Returns List[Row].
    """
    headers = {"Content-Type": "application/json", "api-key": API_KEY}

    system_msg = (
        "You output ONLY JSON: an array of objects with exactly these keys: "
        "placeholder_name, old_value, prompt, new_value. No prose/markdown. "
        "Never output any object where a value equals its key name "
        '(e.g., "old_value":"old_value"). If you cannot find valid rows, return [].'
    )

    user_prompt = f"""
You are a meticulous placeholder auditor for corporate documents (transfer pricing + finance context).

OUTPUT
Return ONLY a JSON array. Each element must have exactly:
- "placeholder_name": string (for section-level use "SECTION:<name>")
- "old_value": string
- "prompt": string (non-empty ONLY if the placeholder is a whole section needing rewrite; otherwise "")
- "new_value": string (empty if <80% confident)

RULES
- No markdown, no comments, no code fences—JSON array only.
- Keep strings terse; prefer exact spans from CONTEXT.
- Identify changes in names/addresses/IDs; dates/fiscal years; financials; FAR; intercompany; benchmarks; governance; jurisdictions/regulations; and any sections to be rewritten.
- **old_value must be an exact substring of CONTEXT that is unique (occurs once).**
- **If a bare value would occur multiple times, expand old_value by including adjacent label/unit/words from CONTEXT until the span is unique (e.g., prefer 'employees: 14 FTEs' over '14').**
- **Do not invent header/example rows; do not use key names as values.**
- **For non-section placeholders, prompt must be empty (''); for SECTION:<name> rows, prompt must be non-empty.**
- **Deduplicate: do not output multiple rows with the same (placeholder_name, old_value).**

CONTEXT
<<<CONTEXT_START
{context}
CONTEXT_END>>>
""".strip()

    attempt = 0
    last_exception = None

    while attempt < max_retries:
        attempt += 1
        try:
            payload = {
                "messages": [
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0,
                "top_p": 1,
                "seed": 7,  # ignored if unsupported
            }
            resp = requests.post(API_ENDPOINT, headers=headers, json=payload, timeout=60)
            resp.raise_for_status()
            data = resp.json()

            text = (
                data.get("choices", [{}])[0]
                    .get("message", {})
                    .get("content", "")
                    .strip()
            )

            rows = parse_rows_json(text)

            # Normalize: if a non-section has a non-empty prompt, blank it.
            for r in rows:
                if r.prompt is None:
                    r.prompt = ""
                if r.new_value is None:
                    r.new_value = ""

            return rows  # success

        except Exception as e:
            last_exception = e
            print(f"[ask_variable_list] attempt {attempt}/{max_retries} failed: {e}")
            time.sleep(wait_seconds)

    # If we exhaust retries, raise to let caller decide how to proceed
    raise RuntimeError(f"ask_variable_list failed after {max_retries} attempts: {last_exception}")


def fill_excel_prompts(prompt: str, context: str, old_value: str, variable_name: str) -> str:
    """
    Call the configured LLM endpoint to fulfill a given prompt for a single excel row.
    Returns the assistant text (stripped) or raises on HTTP error.
    """
    headers = {"Content-Type": "application/json", "api-key": API_KEY}
    system_msg = (
        "You are an expert on Transfer Pricing and financial analysis. "
        "You are provided the old value of a variable, and the name of the variable."
        "Update this value with the latest data and STAY AS CLOSE TO OLD VALUE AS POSSIBLE"
        "Use either the internet or the context to infer the new value, and LEAVE IT THE SAME if it is not determinate. If you use the internet then cite sources. "
    )
    user_content = (
        f"old_value: {old_value}\n"
        f"variable_name: {variable_name}\n\n"
        f"Prompt: {prompt}\n\n"
        f"Context:\n{context}"
    )
    messages = [
        {"role": "system", "content": system_msg},
        {"role": "user", "content": user_content},
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


def _prefill_last_year_from_prompts(excel_file, context: str) -> str | None:
    """
    Creates/updates an Excel workbook of variables based on ask_variable_list(context).
    If excel_file cannot be loaded, a new workbook is created with columns:
    placeholder_name, old_value, prompt, new_value.

    Avoids appending duplicate (placeholder_name, old_value) rows by updating existing rows.
    """
    try:
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
        except Exception:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "variables"

        required_cols = ["placeholder_name", "old_value", "prompt", "new_value"]

        # Force header to row 1
        current_header = [ws.cell(row=1, column=i).value for i in range(1, len(required_cols)+1)]
        if current_header != required_cols:
            for i, col in enumerate(required_cols, start=1):
                ws.cell(row=1, column=i, value=col)

        existing_old_to_row = {}
        for r in range(2, ws.max_row + 1):
            try:
                ov = ws.cell(row=r, column=2).value
                if ov is not None:
                    existing_old_to_row[str(ov).strip()] = r
            except Exception:
                continue

        rows = ask_variable_list(context, 4)
        if rows is None:
            return None

        for r in rows:
            placeholder_name = getattr(r, "placeholder_name", "") or ""
            old_value = (getattr(r, "old_value", "") or "").strip()
            prompt = getattr(r, "prompt", "") or ""
            new_value = getattr(r, "new_value", "") or ""

            if not old_value and not placeholder_name:
                continue

            if old_value in existing_old_to_row:
                row_idx = existing_old_to_row[old_value]
                ws.cell(row=row_idx, column=1, value=placeholder_name)
                ws.cell(row=row_idx, column=2, value=old_value)
                if prompt:
                    ws.cell(row=row_idx, column=3, value=prompt)
                if new_value:
                    ws.cell(row=row_idx, column=4, value=new_value)
            else:
                ws.append([placeholder_name, old_value, prompt, new_value])
                existing_old_to_row[old_value] = ws.max_row

        tmpdir = tempfile.mkdtemp()
        out_path = os.path.join(tmpdir, "variables_prefilled.xlsx")
        wb.save(out_path)
        return out_path

    except Exception as e:
        print(f"Error in _prefill_last_year_from_prompts: {e}")
        return None


def load_and_annotate_replacements(excel_file, context: str) -> Tuple[Dict[str, str], str | None]:
    """
    Loads or creates workbook and for each row with a prompt calls the LLM
    to fill new_value. Avoids returning duplicate replacement keys.
    Returns (replacements_dict, filled_excel_path).
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

        variable_name = _get(col_placeholder)
        old_value        = _get(col_old)
        prompt_text      = _get(col_prompt)
        new_value_curr   = _get(col_new)

        # If prompt present, query LLM and write into new_value (only overwrite if result non-empty)
        if prompt_text:
            try:
                llm_out = fill_excel_prompts(prompt_text, context, old_value, variable_name)
                llm_out = (llm_out or "").strip()
                # Only write back if LLM gave a non-empty answer (avoid erasing manual values)
                if llm_out:
                    ws.cell(row=r, column=col_new, value=llm_out)
                    new_value_curr = llm_out
            except Exception:
                # On error, leave whatever was already in new_value
                pass

        # Build replacements only when new_value is non-empty and different from old_value
        if old_value and new_value_curr and old_value != new_value_curr:
            # Avoid collisions: prefer the first mapping (do not overwrite existing keys)
            if old_value not in replacements:
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
# DOCX helpers
# =========================

def replace_in_paragraph(p, replacements: dict):
    """
    Hybrid replacer:
      - Concatenate run texts, replace unique keys longest-first, write back preserving run boundaries.
    """
    if not replacements:
        return

    # ensure keys/values are strings
    repl = {str(k): ("" if v is None else str(v)) for k, v in replacements.items()}

    p_elm = p._p
    ns = p_elm.nsmap or {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xml": "http://www.w3.org/XML/1998/namespace",
    }

    t_nodes = p_elm.findall(".//w:t", namespaces=ns)
    if not t_nodes:
        return

    originals = [(t, t.text or "") for t in t_nodes]
    full = "".join(txt for _, txt in originals)
    if not full:
        return

    keys = [k for k in repl.keys() if k and (k in full)]
    if not keys:
        return

    keys.sort(key=len, reverse=True)
    unique_keys = [k for k in keys if full.count(k) == 1]
    if not unique_keys:
        return  # nothing unambiguous in this paragraph

    pattern = re.compile("|".join(re.escape(k) for k in unique_keys))
    new_full, nsubs = pattern.subn(lambda m: repl[m.group(0)], full)

    if nsubs == 0 or new_full == full:
        return  # no change

    # Redistribute back using original lengths to keep run boundaries intact
    lengths = [len(txt) for _, txt in originals]
    pos = 0
    n = len(originals)
    for i in range(n):
        t, _oldtxt = originals[i]
        take = lengths[i] if i < n - 1 else max(0, len(new_full) - pos)
        segment = new_full[pos:pos + take] if take >= 0 else ""
        t.text = segment
        pos += lengths[i] if i < n - 1 else len(new_full) - pos

        # Preserve leading/trailing spaces for this text node
        if segment and (segment[0].isspace() or segment[-1].isspace()):
            t.set(qn("xml:space"), "preserve")


def replace_placeholders_docx(doc: Document, replacements: dict):
    """Replace placeholders AFTER the first page break, preserving images and footnotes."""
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

    # Tables
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)

    # Headers/footers
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
    Extract plain text from a DOCX template file, including:
      - headers (all sections)
      - body paragraphs (incl. TOC paragraphs)
      - tables (body + headers/footers)
      - footers (all sections)
    Returns a single string with items separated by newlines.
    """
    if not template_file:
        return ""

    try:
        doc = Document(template_file)
        chunks = []
        seen = set()  # avoid accidental duplicates if the same text is reachable twice

        def add(text: str):
            t = (text or "").strip()
            if t and t not in seen:
                chunks.append(t)
                seen.add(t)

        def add_paragraphs(paragraphs):
            for p in paragraphs:
                # include all body paragraphs; this already captures TOC text if present
                add(p.text)

        def add_tables(tables):
            # Flatten table content row-by-row
            for tbl in tables:
                for row in tbl.rows:
                    cells_txt = []
                    for cell in row.cells:
                        # cell.text returns the concatenated text of all paragraphs in the cell
                        ct = (cell.text or "").strip()
                        if ct:
                            cells_txt.append(ct)
                    if cells_txt:
                        # Use a lightweight delimiter to keep context readable
                        add(" | ".join(cells_txt))

        # 1) Headers (per section)
        for sec in doc.sections:
            hdr = sec.header
            if hdr:
                add_paragraphs(hdr.paragraphs)
                add_tables(hdr.tables)

        # 2) Body paragraphs (this includes TOC content if the document has a generated TOC)
        add_paragraphs(doc.paragraphs)

        # 3) Body tables
        add_tables(doc.tables)

        # 4) Footers (per section)
        for sec in doc.sections:
            ftr = sec.footer
            if ftr:
                add_paragraphs(ftr.paragraphs)
                add_tables(ftr.tables)

        return "\n".join(chunks)

    except Exception as e:
        print(f"Error loading template: {e}")
        return ""
