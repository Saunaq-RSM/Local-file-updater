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
import json
from typing import List, Dict
import time
from zipfile import ZipFile
from lxml import etree
from io import BytesIO
from docx.oxml.ns import qn
from docx.shared import RGBColor
from zipfile import ZipFile
from lxml import etree
import re

# ——— Azure OpenAI config ———
# Expect these to be set by the Streamlit frontend via secrets or environment variables
API_KEY = None  # to be set by frontend
API_ENDPOINT = None  # to be set by frontend

from pydantic import BaseModel, TypeAdapter
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

def row_to_dict(r: Row) -> dict:
    return r.model_dump()



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

def fill_excel_prompts(prompt: str, context: str, old_value: str, variable_name: str) -> str:
    headers = {"Content-Type": "application/json", "api-key": API_KEY}
    system_msg = (
        "You are an expert on Transfer Pricing and financial analysis. "
        "You are provided the old value of a variable, and the name of the variable."
        "Update this value with the latest data and STAY AS CLOSE TO OLD VALUE AS POSSIBLE"
        "Use either the internet or the context to infer the new value, and LEAVE IT THE SAME if it is not determinate. If you use the internet then cite sources. "
    )
    prompt = ("old_value: {old_value}"
              "variable_name: {variable_name}\n\n"
              "Prompt: " + prompt + 
              "Context: \n") + context
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
    Loads or creates workbook and for each row with a prompt calls the LLM
    to fill new_value. Avoids returning duplicate replacement keys.
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
# DOCX (existing) helpers
# =========================
def collapse_runs(paragraph):
    from docx.oxml.ns import qn
    text = "".join(r.text for r in paragraph.runs)
    for r in reversed(paragraph.runs):
        r._element.getparent().remove(r._element)
    paragraph.add_run(text)


def replace_in_paragraph(p, replacements: dict):
    """
    Hybrid replacer (upgraded):
      1) Concatenate all <w:t> texts.
      2) Replace using a single regex alternation of unique keys (longest-first).
         - Skips keys that appear >1 time in this paragraph (ambiguous here).
      3) Write back across the SAME number of <w:t> nodes using original lengths.
      4) Preserve spaces via xml:space="preserve".
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

    # ---- Build a longest-first set of present, unique keys in THIS paragraph ----
    keys = [k for k in repl.keys() if k and (k in full)]
    if not keys:
        return

    # Prefer longest keys first to avoid partial overlaps
    keys.sort(key=len, reverse=True)

    # Enforce paragraph-level uniqueness (skip ambiguous keys here)
    unique_keys = [k for k in keys if full.count(k) == 1]
    if not unique_keys:
        return  # nothing unambiguous in this paragraph

    # Single-pass regex alternation
    pattern = re.compile("|".join(re.escape(k) for k in unique_keys))
    new_full, nsubs = pattern.subn(lambda m: repl[m.group(0)], full)

    if nsubs == 0 or new_full == full:
        return  # no change

    # ---- Redistribute back using original lengths to keep run boundaries intact ----
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

    # Optional: cleanup formatting on runs that no longer contain placeholders
    # _clear_red_on_non_placeholder_runs(p, repl)
    # _clear_paragraph_bullet_color(p)


def _clear_paragraph_bullet_color(p):
    try:
        if p.style and p.style.font and getattr(p.style.font, "color", None):
            p.style.font.color.rgb = None
            p.style.font.color.theme_color = None
    except Exception:
        pass

def _run_is_explicit_red(run) -> bool:
    c = getattr(run.font, "color", None)
    if not c or getattr(c, "rgb", None) is None:
        return False
    try:
        r, g, b = c.rgb[0], c.rgb[1], c.rgb[2]
        return (200 <= r <= 255 and 0 <= g <= 80 and 0 <= b <= 80) or (r >= 100 and b <=20 and g <=20)
    except Exception:
        return False

def _clear_run_color(run):
    if getattr(run.font, "color", None):
        try:
            run.font.color.rgb = None
        except Exception:
            pass
        try:
            run.font.color.theme_color = None
        except Exception:
            pass

def _clear_red_on_non_placeholder_runs(p, replacements: dict):
    keys = list(replacements.keys())
    for run in p.runs:
        text = run.text or ""
        if not text:
            continue
        # If this run used to be a placeholder (red), but now has no placeholders, clear color
        if _run_is_explicit_red(run):
            if (not any(k in text for k in keys)) and ("{{" not in text and "}}" not in text):
                _clear_run_color(run)




def _rewrite_footnotes_xml_bytes(docx_bytes: bytes, replacements: dict) -> bytes:
    """
    Open a .docx (zip) from bytes, replace placeholders inside word/footnotes.xml,
    and return new .docx bytes. If footnotes.xml is missing, return the original bytes.
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(BytesIO(docx_bytes)) as zin:
        names = {i.filename for i in zin.infolist()}
        if "word/footnotes.xml" not in names:
            return docx_bytes  # no footnotes part

        # Read and parse original footnotes.xml
        foot_xml = zin.read("word/footnotes.xml")
        root = etree.fromstring(foot_xml)

        # Replace inside every w:t under w:footnote
        # (Note: this is robust for tokens contained in a single text node.
        # If your placeholders can split across runs, prefer the python-docx path.)
        for t in root.findall(".//w:footnote//w:t", namespaces=ns):
            if t.text:
                new_text = t.text
                for ph, val in replacements.items():
                    if ph in new_text:
                        new_text = new_text.replace(ph, val)
                if new_text != t.text:
                    t.text = new_text

        # Build a new .docx with modified footnotes.xml
        out_buf = BytesIO()
        with ZipFile(out_buf, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/footnotes.xml":
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
                zout.writestr(item, data)
        return out_buf.getvalue()


def _apply_footnotes_xml_fallback_in_place(docx_path: str, replacements: dict) -> None:
    """
    Read a saved .docx from disk, run the XML fallback, and overwrite it in place.
    Safe no-op if the document has no footnotes.xml.
    """
    try:
        with open(docx_path, "rb") as f:
            original = f.read()
        updated = _rewrite_footnotes_xml_bytes(original, replacements)
        if updated != original:
            with open(docx_path, "wb") as f:
                f.write(updated)
    except Exception:
        # Be defensive: never fail the whole pipeline if footnote rewrite trips.
        pass


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

    # Footnotes (if present)


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

_BOUND = set("\n\r\t,.;:()[]{}<>—–-•|")

def _first_occurrence_span(haystack: str, needle: str) -> Tuple[int, int]:
    i = haystack.find(needle)
    return (i, i + len(needle)) if i != -1 else (-1, -1)

def make_unique_span(context: str, core: str, max_expand: int = 120) -> Tuple[str, Tuple[int,int]]:
    """
    Returns (unique_span, (L, R)). If core not found, returns ("", (0,0)).
    Expands to nearest punctuation/newline boundaries until unique or limit hit.
    """
    core = (core or "").strip()
    if not core:
        return "", (0, 0)

    if context.count(core) == 1:
        s, e = _first_occurrence_span(context, core)
        return context[s:e], (s, e)

    s, e = _first_occurrence_span(context, core)
    if s == -1:
        return "", (0, 0)

    L, R = s, e
    expanded = 0
    while expanded < max_expand and context.count(context[L:R]) != 1:
        # expand left to previous boundary
        l = L
        while l > 0 and context[l-1] not in _BOUND:
            l -= 1
        if l == L and L > 0:  # no boundary nearby, nudge
            l = max(0, L - 8)
        L = l

        if context.count(context[L:R]) == 1:
            break

        # expand right to next boundary
        r = R
        while r < len(context) and context[r] not in _BOUND:
            r += 1
        if r == R and R < len(context):  # no boundary nearby, nudge
            r = min(len(context), R + 8)
        R = r

        expanded = (s - L) + (R - e)

    span = context[L:R]
    return (span if span and context.count(span) == 1 else ""), (L, R)



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
                # keep prior behavior but do not silently override new_value if empty here;
                # let the caller handle defaults/overwrites.
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


def _prefill_last_year_from_prompts(excel_file, context: str) -> str | None:
    """
    Creates/updates an Excel workbook of variables based on ask_variable_list(context).
    If excel_file cannot be loaded, a new workbook is created with columns:
    placeholder_name, old_value, prompt, new_value.

    This version avoids appending duplicate (placeholder_name, old_value) rows:
    - If the same old_value already exists, update the existing row's prompt/new_value
      instead of appending a duplicate.
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

        # --- Force header to row 1 ---
        current_header = [ws.cell(row=1, column=i).value for i in range(1, len(required_cols)+1)]
        if current_header != required_cols:
            # overwrite row 1 explicitly
            for i, col in enumerate(required_cols, start=1):
                ws.cell(row=1, column=i, value=col)

        # Build a lookup of existing old_values -> row index (to avoid duplicates)
        existing_old_to_row = {}
        for r in range(2, ws.max_row + 1):
            try:
                ov = ws.cell(row=r, column=2).value  # default old_value is column 2 (B)
                if ov is not None:
                    existing_old_to_row[str(ov).strip()] = r
            except Exception:
                continue

        # --- Get rows from LLM (list of row objects) ---
        rows = ask_variable_list(context, 4)
        if rows is None:
            return None

        # --- Insert or update rows ---
        for r in rows:
            placeholder_name = getattr(r, "placeholder_name", "") or ""
            old_value = (getattr(r, "old_value", "") or "").strip()
            prompt = getattr(r, "prompt", "") or ""
            new_value = getattr(r, "new_value", "") or ""

            if not old_value and not placeholder_name:
                # ignore empty rows
                continue

            if old_value in existing_old_to_row:
                # update existing row in place (columns A..D)
                row_idx = existing_old_to_row[old_value]
                ws.cell(row=row_idx, column=1, value=placeholder_name)
                ws.cell(row=row_idx, column=2, value=old_value)
                # Only overwrite prompt/new_value if not empty (so we don't unintentionally erase existing data)
                if prompt:
                    ws.cell(row=row_idx, column=3, value=prompt)
                if new_value:
                    ws.cell(row=row_idx, column=4, value=new_value)
            else:
                # append new row
                ws.append([placeholder_name, old_value, prompt, new_value])
                existing_old_to_row[old_value] = ws.max_row  # update map

        # --- Save to a temp file ---
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
    (rep, filled_path) = load_and_annotate_replacements(excel_for_processing, ctx)
    return filled_path
