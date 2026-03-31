#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
US Economic Indicators Tracker  v6
────────────────────────────────────
Grabs US macro data from FRED, rips PMI numbers straight out of PDFs,
and spits out a polished Excel dashboard. That's basically it.

Drop the PDFs in, run the script, open the xlsx. Done.

Run: py -3 economic_tracker.py
"""

import os, re, requests, pdfplumber, openpyxl
from openpyxl.styles      import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils       import get_column_letter
from openpyxl.chart       import LineChart, BarChart, Reference
from openpyxl.chart.axis  import ChartLines
from datetime             import datetime


# ── CONFIG ────────────────────────────────────────────────────────────────────

# FRED API key — free from fred.stlouisfed.org. See README if it stops working.
FRED_API_KEY  = "YOUR_FRED_API_KEY_HERE"

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))   # keeps the project portable
ISM_FOLDER    = os.path.join(BASE_DIR, "ISM and MNI Reports")
MANUAL_FILE   = os.path.join(ISM_FOLDER, "ISM_Manual_Input.xlsx")
OUTPUT_FILE   = os.path.join(BASE_DIR,   "US_Economic_Tracker.xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR,   "US_Economic_Tracker_TEMPLATE.xlsx")  # optional — see README
FONT_NAME     = "Calibri"

# Gets flipped to True in main() when a template is detected.
# In template mode the build functions update values only — your custom
# formatting, colours, column widths etc. are left completely untouched.
_TEMPLATE_MODE = False


# ── COLOURS — edit colorz!!! ──────────────────────────────────────────────────
# (NOT the traffic-light signal ones at the top — those are sacred and reserved
#  exclusively for the expansionary / neutral / contractionary indicators)
C = {
    # Signal colours — green / amber / red. Don't touch these.
    "exp":       "1E8449",
    "neu":       "D4870A",
    "con":       "B03A2E",

    # Section header colours — safe to tweak here or just edit the template, chose random colours
    "jm":        "1A3A5C",
    "inf":       "4A235A",
    "ea":        "7E5109",

    # UI chrome
    "title":     "0D1B2A",
    "sub":       "1A2E42",
    "col_hdr":   "1F3A5C",
    "chart_hdr": "1C2833",

    # Row backgrounds
    "white":     "FFFFFF",
    "row0":      "FFFFFF",
    "row1":      "F2F3F4",
    "thresh_bg": "EEF2F7",
    "thresh_fc": "1A252F",

    # Pending cells — shows when a PDF hasn't been dropped in yet
    "input":     "FEF9E7",
    "input_fc":  "6E4800",

    # Text
    "dgray":     "4A4A4A",
    "lgray":     "E8EAED",
    "note_bg":   "F8F9FA",

    # Chart line colours — one per series, edit freely
    "ch_ur":     "2980B9",
    "ch_ic":     "E67E22",
    "ch_nfp":    "27AE60",
    "ch_cpi":    "C0392B",
    "ch_ccpi":   "8E44AD",
    "ch_ppi":    "7F8C8D",
    "ch_imfg":   "1ABC9C",
    "ch_isvc":   "F39C12",
    "ch_chi":    "2C3E50",
    "ch_cc":     "3498DB",
    "ch_ff":     "E74C3C",
}

CAT_COLOR = {
    "Job Market":          C["jm"],
    "Inflation":           C["inf"],
    "Economic Activities": C["ea"],
}
CAT_TAB = {
    "Job Market":          "1A3A5C",
    "Inflation":           "4A235A",
    "Economic Activities": "7E5109",
}

MONTH_MAP = dict(
    january="01",  february="02", march="03",    april="04",
    may="05",      june="06",     july="07",     august="08",
    september="09",october="10",  november="11", december="12",
)

# Keys whose values come from PDFs rather than FRED
PDF_KEYS = {"ism_mfg", "ism_svc", "chicago_pmi"}

# Y-axis number format per metric — used in charts
CHART_Y_FMT = {
    "unemployment_rate": "0.0",
    "initial_claims":    "#,##0",
    "nonfarm_payrolls":  '+#,##0;-#,##0',
    "cpi_yoy":           "0.00",
    "core_cpi_yoy":      "0.00",
    "ppi_mom":           "0.00",
    "ism_mfg":           "0.0",
    "ism_svc":           "0.0",
    "chicago_pmi":       "0.0",
    "consumer_conf":     "0.0",
    "fedfunds":          "0.00",
}


# ── STYLE HELPERS ─────────────────────────────────────────────────────────────
# Shortcuts so I don't have to type PatternFill(...) every five seconds

_fill  = lambda h: PatternFill(start_color=h, end_color=h, fill_type="solid")
_font  = lambda bold=False, color="000000", size=10, italic=False: \
             Font(bold=bold, color=color, size=size, italic=italic, name=FONT_NAME)
_align = lambda h="center", v="center", wrap=False: \
             Alignment(horizontal=h, vertical=v, wrap_text=wrap)
_S     = lambda s: Side(style=s)

def _border_all(s="thin"):
    b = _S(s)
    return Border(left=b, right=b, top=b, bottom=b)

def box_border(ws, r1, c1, r2, c2, weight="medium"):
    if _TEMPLATE_MODE: return
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = Border(
                left   = _S(weight) if c == c1 else _S("thin"),
                right  = _S(weight) if c == c2 else _S("thin"),
                top    = _S(weight) if r == r1 else _S("thin"),
                bottom = _S(weight) if r == r2 else _S("thin"),
            )

def W(ws, row, col, value=None, bg=None, fc="000000", bold=False, size=10,
      italic=False, h="center", v="center", wrap=False, fmt=None, bs="thin"):
    c = ws.cell(row=row, column=col, value=value)
    if not _TEMPLATE_MODE:
        c.font      = _font(bold=bold, color=fc, size=size, italic=italic)
        c.alignment = _align(h=h, v=v, wrap=wrap)
        if bg:  c.fill = _fill(bg)
        c.border = _border_all(bs)
    if fmt: c.number_format = fmt
    return c

def M(ws, r1, c1, r2, c2, value=None, bg=None, fc="000000",
      bold=False, size=10, h="center", v="center", wrap=False, italic=False):
    if not _TEMPLATE_MODE:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    if not _TEMPLATE_MODE:
        c.font      = _font(bold=bold, color=fc, size=size, italic=italic)
        c.alignment = _align(h=h, v=v, wrap=wrap)
        if bg: c.fill = _fill(bg)
        box_border(ws, r1, c1, r2, c2)
    return c

def signal_M(ws, r1, c1, r2, c2, value=None, bg=None, fc="000000",
             bold=False, size=10, h="center", v="center", italic=False):
    """Like M() but always applies styling — signal colour IS the data, can't skip it."""
    if not _TEMPLATE_MODE:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    c.font      = _font(bold=bold, color=fc, size=size, italic=italic)
    c.alignment = _align(h=h, v=v)
    if bg: c.fill = _fill(bg)
    if not _TEMPLATE_MODE:
        box_border(ws, r1, c1, r2, c2)
    return c

def set_widths(ws, widths):
    if _TEMPLATE_MODE: return   # preserve the user's column widths
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def rh(ws, row, height):
    if _TEMPLATE_MODE: return   # preserve the user's row heights
    ws.row_dimensions[row].height = height

def p_bg(p):  return {"expansionary": C["exp"], "neutral": C["neu"], "contractionary": C["con"]}.get(p, C["neu"])
def p_lbl(p): return {"expansionary": "EXPANSIONARY", "neutral": "NEUTRAL", "contractionary": "CONTRACTIONARY"}.get(p, "NEUTRAL")
def p_arr(p): return {"expansionary": "▲  GROWTH", "neutral": "◆  STABLE", "contractionary": "▼  CAUTION"}.get(p, "◆  STABLE")

def pressure_cell(ws, row, col, p, pending=False, size=9):
    if pending:
        c = ws.cell(row=row, column=col, value="PENDING")
        c.fill = _fill(C["input"]); c.font = _font(bold=True, color=C["input_fc"], size=size)
        c.alignment = _align(); c.border = _border_all()
        return c
    c = ws.cell(row=row, column=col, value=p_lbl(p))
    c.fill = _fill(p_bg(p)); c.font = _font(bold=True, color="FFFFFF", size=size)
    c.alignment = _align(); c.border = _border_all()
    return c

def signal_cell(ws, row, col, p, pending=False, size=9):
    if pending:
        c = ws.cell(row=row, column=col, value="? PENDING")
        c.fill = _fill(C["input"]); c.font = _font(bold=True, color=C["input_fc"], size=size)
        c.alignment = _align(); c.border = _border_all()
        return c
    c = ws.cell(row=row, column=col, value=p_arr(p))
    c.fill = _fill(p_bg(p)); c.font = _font(bold=True, color="FFFFFF", size=size)
    c.alignment = _align(); c.border = _border_all()
    return c


# ── FRED HELPERS ──────────────────────────────────────────────────────────────

def _fred(series_id, limit=26):
    url = "https://api.stlouisfed.org/fred/series/observations"
    try:
        r = requests.get(url, params=dict(
            series_id=series_id, api_key=FRED_API_KEY,
            file_type="json", sort_order="desc", limit=limit), timeout=20)
        r.raise_for_status()
        return [(o["date"], float(o["value"]))
                for o in r.json().get("observations", []) if o["value"] != "."]
    except Exception as e:
        print(f"  [WARN] FRED {series_id}: {e}"); return []

def fetch_latest(sid):
    raw = _fred(sid, 14)
    return (None, []) if not raw else (raw[0][1], raw[:12])

def fetch_yoy(sid):
    raw = _fred(sid, 26)
    if len(raw) < 13: return None, []
    hist = [(raw[i][0], round((raw[i][1] - raw[i+12][1]) / raw[i+12][1] * 100, 2))
            for i in range(min(12, len(raw) - 12))]
    return round((raw[0][1] - raw[12][1]) / raw[12][1] * 100, 2), hist

def fetch_mom_pct(sid):
    raw = _fred(sid, 14)
    if len(raw) < 2: return None, []
    hist = [(raw[i][0], round((raw[i][1] - raw[i+1][1]) / raw[i+1][1] * 100, 2))
            for i in range(min(12, len(raw) - 1))]
    return round((raw[0][1] - raw[1][1]) / raw[1][1] * 100, 2), hist

def fetch_payrolls():
    raw = _fred("PAYEMS", 14)
    if len(raw) < 2: return None, []
    hist = [(raw[i][0], round(raw[i][1] - raw[i+1][1], 1))
            for i in range(min(12, len(raw) - 1))]
    return round(raw[0][1] - raw[1][1], 1), hist

def fetch_fedfunds(months=48):
    raw = _fred("FEDFUNDS", months + 2)
    return raw[:months] if raw else []


# ── PDF PARSING ───────────────────────────────────────────────────────────────

def _pdf_text(path):
    try:
        with pdfplumber.open(path) as pdf:
            return "".join(p.extract_text() or "" for p in pdf.pages[:2])
    except Exception as e:
        print(f"  [WARN] PDF read failed — {os.path.basename(path)}: {e}")
        return ""

def _extract_report_month(txt, fpath=""):
    m = re.search(
        r'(january|february|march|april|may|june|july|august|'
        r'september|october|november|december)\s+(\d{4})',
        txt, re.IGNORECASE)
    if m:
        return f"{m.group(2)}-{MONTH_MAP[m.group(1).lower()]}-01"
    m_fn = re.match(r'(\d{4})(\d{2})\d{2}', os.path.basename(fpath))
    if m_fn:
        return f"{m_fn.group(1)}-{m_fn.group(2)}-01"
    return None

def parse_ism_pdf(path):
    txt = _pdf_text(path)
    if not txt: return None, None
    m = re.search(r'PMI[®\u00ae\u2122]?\s+at\s+(\d+\.?\d*)\s*%', txt, re.IGNORECASE)
    if not m:
        m = re.search(r'PMI[®\u00ae\u2122]?[^\n]{0,40}?at\s+(\d+\.?\d*)', txt, re.IGNORECASE)
    if not m:
        m = re.search(r'registered\s+(\d+\.?\d*)\s+percent', txt, re.IGNORECASE)
    if not m:
        m = re.search(r'PMI[^\n]{0,20}at\s+(\d+\.?\d*)', txt, re.IGNORECASE)
    pmi = float(m.group(1)) if m else None
    return _extract_report_month(txt, path), pmi

def parse_mni_pdf(path):
    txt = _pdf_text(path)
    if not txt: return None, None
    # Four fallback patterns because MNI can't decide how to write the same sentence
    m = re.search(
        r'\bto\s+(\d+\.?\d*)\s+in\s+'
        r'(?:january|february|march|april|may|june|july|august|'
        r'september|october|november|december)\b',
        txt, re.IGNORECASE)
    if not m:
        m = re.search(
            r'(?:climbed|fell|rose|dropped|slipped|edged|moved|eased)\s+'
            r'\d+\.?\d*\s+points?\s+to\s+(\d+\.?\d*)',
            txt, re.IGNORECASE)
    if not m:
        m = re.search(r'Barometer[^.]{0,80}?(\d{2}\.\d)\s*(?:percent\b|in\b)', txt, re.IGNORECASE)
    if not m:
        m = re.search(r'\bat\s+(\d{2}\.\d)\b', txt)
    val = float(m.group(1)) if m else None
    return _extract_report_month(txt, path), val

def parse_reports_folder():
    mfg_date = mfg_val = svc_date = svc_val = chi_date = chi_val = None
    if not os.path.isdir(ISM_FOLDER):
        print(f"  [WARN] Reports folder not found: {ISM_FOLDER}")
        return mfg_date, mfg_val, svc_date, svc_val, chi_date, chi_val

    for fname in sorted(os.listdir(ISM_FOLDER)):
        if not fname.lower().endswith(".pdf"): continue
        fpath = os.path.join(ISM_FOLDER, fname)
        fl    = fname.lower()
        is_mni = "mni" in fl or "barometer" in fl
        is_mfg = (not is_mni) and ("manuf" in fl)
        is_svc = (not is_mni) and ("serv" in fl or "non-manuf" in fl
                                   or "nonmanuf" in fl or "non_manuf" in fl)
        if is_mni:
            d, v = parse_mni_pdf(fpath)
            if v is not None:
                chi_date, chi_val = d, v
                print(f"  PDF -> Chicago PMI        {v:5.1f}   ({d})")
            else:
                print(f"  [WARN] Chicago PMI extraction failed: {fname}")
        elif is_mfg:
            d, v = parse_ism_pdf(fpath)
            if v is not None:
                mfg_date, mfg_val = d, v
                print(f"  PDF -> ISM Mfg PMI        {v:5.1f}   ({d})")
            else:
                print(f"  [WARN] ISM Mfg PMI extraction failed: {fname}")
        elif is_svc:
            d, v = parse_ism_pdf(fpath)
            if v is not None:
                svc_date, svc_val = d, v
                print(f"  PDF -> ISM Services PMI   {v:5.1f}   ({d})")
            else:
                print(f"  [WARN] ISM Svc PMI extraction failed: {fname}")
        else:
            print(f"  [SKIP] Unrecognised PDF (rename to include manuf/serv/mni): {fname}")

    return mfg_date, mfg_val, svc_date, svc_val, chi_date, chi_val


# ── ISM MANUAL INPUT WORKBOOK ─────────────────────────────────────────────────

MI_HDR_ROW  = 4
MI_DATA_ROW = 5

def _mi_col_headers():
    return ["Date (YYYY-MM-DD)", "ISM Manufacturing PMI",
            "ISM Non-Manufacturing PMI", "Chicago PMI", "Source / Notes"]

def create_manual_input_if_needed():
    if os.path.exists(MANUAL_FILE): return
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Monthly PMI Data"
    ws.sheet_view.showGridLines = False
    set_widths(ws, [2, 18, 22, 22, 18, 44])
    rh(ws, 1, 44); rh(ws, 2, 20); rh(ws, 3, 20); rh(ws, 4, 28)
    M(ws, 1, 2, 1, 6,
      "ISM & CHICAGO PMI  —  Monthly Data Log (Auto-populated from PDFs)",
      bg=C["chart_hdr"], fc="FFFFFF", bold=True, size=14, h="left")
    M(ws, 2, 2, 2, 6,
      "This file is read automatically by economic_tracker.py on every run.  "
      "Drop the latest ISM and MNI PDFs into 'ISM and MNI Reports' and re-run.",
      bg=C["lgray"], fc=C["dgray"], size=9, italic=True, h="left")
    M(ws, 3, 2, 3, 6,
      "ISM Manufacturing: 1st business day  |  "
      "ISM Non-Mfg: 3rd business day  |  "
      "Chicago PMI (MNI): last business day of month",
      bg="FFF8DC", fc=C["input_fc"], size=9, italic=True, h="left")
    for ci, hdr in enumerate(_mi_col_headers(), 2):
        W(ws, MI_HDR_ROW, ci, hdr, bg=C["col_hdr"], fc="FFFFFF",
          bold=True, size=9, bs="medium")
    wb.save(MANUAL_FILE)
    print(f"  Created: {MANUAL_FILE}")

def _mi_find_row_for_month(ws, month_str):
    for r in range(MI_DATA_ROW, ws.max_row + 1):
        cell_val = ws.cell(r, 2).value
        if cell_val and str(cell_val)[:7] == month_str[:7]:
            return r
    return None

def update_manual_input(mfg_date, mfg_val, svc_date, svc_val,
                        chi_date=None, chi_val=None):
    create_manual_input_if_needed()
    wb = openpyxl.load_workbook(MANUAL_FILE)
    ws = wb["Monthly PMI Data"]
    report_date = mfg_date or svc_date or chi_date
    if not report_date:
        print("  No report date — skipping data log update"); wb.close(); return

    month_str    = report_date[:7]
    existing_row = _mi_find_row_for_month(ws, month_str)

    if existing_row:
        def _patch(col, val):
            cell = ws.cell(existing_row, col)
            if val is not None and cell.value in (None, ""):
                cell.value = val
        _patch(3, mfg_val); _patch(4, svc_val); _patch(5, chi_val)
        src = ws.cell(existing_row, 6)
        if not src.value: src.value = "PDF auto-extracted"
        wb.save(MANUAL_FILE)
        print(f"  Data log patched  {month_str}  MFG={mfg_val}  SVC={svc_val}  CHI={chi_val}")
        return

    next_row = MI_DATA_ROW
    for row in ws.iter_rows(min_row=MI_DATA_ROW, min_col=2, max_col=2, values_only=True):
        if row[0] is None: break
        next_row += 1

    row_bg = C["row0"] if (next_row - MI_DATA_ROW) % 2 == 0 else C["row1"]
    vals   = [report_date,
              mfg_val if mfg_val is not None else "",
              svc_val if svc_val is not None else "",
              chi_val if chi_val is not None else "",
              "PDF auto-extracted"]
    for ci, v in enumerate(vals, 2):
        c = ws.cell(row=next_row, column=ci, value=v)
        c.fill = _fill(row_bg); c.font = _font(size=9, bold=(ci == 2))
        c.alignment = _align(h="center" if ci != 6 else "left")
        c.border = _border_all()
    ws.row_dimensions[next_row].height = 20
    for r in range(MI_DATA_ROW, next_row + 1):
        stripe = C["row0"] if (r - MI_DATA_ROW) % 2 == 0 else C["row1"]
        for ci in range(2, 7):
            ws.cell(r, ci).fill = _fill(stripe)
    wb.save(MANUAL_FILE)
    print(f"  Data log updated  {month_str}  MFG={mfg_val}  SVC={svc_val}  CHI={chi_val}")

def read_manual_input():
    empty = {"ism_mfg": (None, []), "ism_svc": (None, []), "chicago_pmi": (None, [])}
    if not os.path.exists(MANUAL_FILE): return empty
    try:
        wb = openpyxl.load_workbook(MANUAL_FILE, data_only=True)
        ws = wb["Monthly PMI Data"]
    except Exception as e:
        print(f"  [WARN] Manual input read failed: {e}"); return empty
    rows = []
    for row in ws.iter_rows(min_row=MI_DATA_ROW, min_col=2, max_col=6, values_only=True):
        date, mfg, svc, chi, *_ = row
        if not date: continue
        rows.append((str(date)[:10], mfg, svc, chi))
    rows.sort(key=lambda x: x[0], reverse=True)
    def _series(idx):
        vals = [(r[0], float(r[idx])) for r in rows if r[idx] not in (None, "", "N/A")]
        return (vals[0][1] if vals else None), vals[:12]
    return {"ism_mfg": _series(1), "ism_svc": _series(2), "chicago_pmi": _series(3)}


# ── CLASSIFICATION ────────────────────────────────────────────────────────────
# Thresholds are conventional — widely used by economists, not plucked from thin air.
# Change them here if you want different trigger points.

RULES = {
    "unemployment_rate": [
        (lambda v: v > 6.5, "contractionary", "Weak economy — elevated unemployment, above 6.5%",          "Expansionary  —  Cut rates / QE"),
        (lambda v: v < 4.0, "expansionary",   "Strong economy — tight labour market, below 4%",             "Contractionary  —  Hike rates / QT"),
        (lambda v: True,    "neutral",         "Moderate unemployment within the 4–6.5% target range",      "Neutral  —  Hold rates"),
    ],
    "initial_claims": [
        (lambda v: v > 350_000, "contractionary", "Elevated layoffs signal a weakening labour market",          "Expansionary  —  Cut rates / QE"),
        (lambda v: v < 250_000, "expansionary",   "Low jobless claims reflect a healthy labour market",          "Contractionary  —  Hike rates / QT"),
        (lambda v: True,        "neutral",         "Claims within the normal 250–350k range",                    "Neutral  —  Hold rates"),
    ],
    "nonfarm_payrolls": [
        (lambda v: v > 250, "expansionary",   "Robust hiring — payrolls growing well above trend",            "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 50,  "contractionary", "Near-stall job creation — economy losing momentum",            "Expansionary  —  Cut rates / QE"),
        (lambda v: True,    "neutral",         "Moderate job creation within the 50–250k range",              "Neutral  —  Hold rates"),
    ],
    "cpi_yoy": [
        (lambda v: v > 2.0, "contractionary", "Inflation above the 2% target — purchasing power eroding",    "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 1.5, "expansionary",   "Below-target inflation — room to stimulate growth",           "Expansionary  —  Cut rates / QE"),
        (lambda v: True,    "neutral",         "Inflation close to the 2% target",                           "Neutral  —  Hold rates"),
    ],
    "core_cpi_yoy": [
        (lambda v: v > 2.0, "contractionary", "Sticky core inflation above the 2% target",                   "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 1.5, "expansionary",   "Core inflation below target — deflationary risk",             "Expansionary  —  Cut rates / QE"),
        (lambda v: True,    "neutral",         "Core inflation on target at approximately 2%",                "Neutral  —  Hold rates"),
    ],
    "ppi_mom": [
        (lambda v: v > 0.2, "contractionary", "Elevated producer prices likely to feed through to CPI",      "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 0.0, "expansionary",   "Falling producer prices — potential deflationary signal",     "Expansionary  —  Cut rates / QE"),
        (lambda v: True,    "neutral",         "Producer price growth contained within 0–0.2% MoM",          "Neutral  —  Hold rates"),
    ],
    "ism_mfg": [
        (lambda v: v > 56, "expansionary",   "Strong manufacturing expansion — reading above 56",            "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 50, "contractionary", "Manufacturing sector in contraction — reading below 50",       "Expansionary  —  Cut rates / QE"),
        (lambda v: True,   "neutral",         "Moderate manufacturing expansion — reading between 50 and 56","Neutral  —  Hold rates"),
    ],
    "ism_svc": [
        (lambda v: v > 56, "expansionary",   "Strong services expansion — reading above 56",                 "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 50, "contractionary", "Services sector in contraction — reading below 50",            "Expansionary  —  Cut rates / QE"),
        (lambda v: True,   "neutral",         "Moderate services expansion — reading between 50 and 56",     "Neutral  —  Hold rates"),
    ],
    "chicago_pmi": [
        (lambda v: v > 56, "expansionary",   "Strong Chicago-area business activity — reading above 56",     "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 50, "contractionary", "Chicago-area business activity contracting — below 50",        "Expansionary  —  Cut rates / QE"),
        (lambda v: True,   "neutral",         "Moderate Chicago-area activity — reading between 50 and 56",  "Neutral  —  Hold rates"),
    ],
    "consumer_conf": [
        (lambda v: v > 120, "expansionary",   "High consumer confidence — strong spending outlook",          "Contractionary  —  Hike rates / QT"),
        (lambda v: v < 100, "contractionary", "Weak consumer confidence — spending headwinds ahead",         "Expansionary  —  Cut rates / QE"),
        (lambda v: True,    "neutral",         "Consumer confidence within the normal 100–120 range",        "Neutral  —  Hold rates"),
    ],
}

def classify(key, val):
    if val is None:
        return "neutral", "No data — add PDF to 'ISM and MNI Reports' folder", "Neutral  —  Hold rates"
    for cond, pressure, impact, fed in RULES.get(key, []):
        if cond(val): return pressure, impact, fed
    return "neutral", "N/A", "N/A"

def overall_signal(pressures):
    known = [p for p in pressures if p in ("expansionary", "neutral", "contractionary")]
    if not known: return "neutral"
    c = {p: known.count(p) for p in ("expansionary", "neutral", "contractionary")}
    if c["contractionary"] / len(known) >= 0.5: return "contractionary"
    if c["expansionary"]   / len(known) >= 0.5: return "expansionary"
    if c["contractionary"] > c["expansionary"]:  return "contractionary"
    if c["expansionary"]   > c["contractionary"]: return "expansionary"
    return "neutral"


# ── METRIC CATALOGUE ──────────────────────────────────────────────────────────

METRICS = [
    ("unemployment_rate", "Unemployment Rate",         "Job Market",          '#,##0.0"%"',        ">6.5% Weak  |  4–6.5% Moderate  |  <4% Strong",    C["ch_ur"],   "Rate (%)"),
    ("initial_claims",    "Initial Jobless Claims",    "Job Market",          "#,##0",              ">350k Weak  |  250–350k Moderate  |  <250k Strong",  C["ch_ic"],   "Claims"),
    ("nonfarm_payrolls",  "Nonfarm Payrolls (MoM k)",  "Job Market",          '+#,##0.0;-#,##0.0', "<50k Weak  |  50–250k Moderate  |  >250k Strong",   C["ch_nfp"],  "Thousands"),
    ("cpi_yoy",           "CPI  (YoY %)",              "Inflation",           '#,##0.00"%"',        ">2% Elevated  |  ~2% On Target  |  <1.5% Low",      C["ch_cpi"],  "YoY (%)"),
    ("core_cpi_yoy",      "Core CPI  (YoY %)",         "Inflation",           '#,##0.00"%"',        ">2% Elevated  |  ~2% On Target  |  <1.5% Low",      C["ch_ccpi"], "YoY (%)"),
    ("ppi_mom",           "PPI  (MoM %)",              "Inflation",           '#,##0.00"%"',        ">0.2% Elevated  |  0–0.2% Moderate  |  <0% Falling",C["ch_ppi"],  "MoM (%)"),
    ("ism_mfg",           "ISM Manufacturing PMI",     "Economic Activities", "#,##0.0",            ">56 Strong  |  50–56 Moderate  |  <50 Contraction", C["ch_imfg"], "Index"),
    ("ism_svc",           "ISM Non-Manufacturing PMI", "Economic Activities", "#,##0.0",            ">56 Strong  |  50–56 Moderate  |  <50 Contraction", C["ch_isvc"], "Index"),
    ("chicago_pmi",       "Chicago PMI",               "Economic Activities", "#,##0.0",            ">56 Strong  |  50–56 Moderate  |  <50 Contraction", C["ch_chi"],  "Index"),
    ("consumer_conf",     "Consumer Confidence (UoM)", "Economic Activities", "#,##0.0",            ">120 High  |  100–120 Moderate  |  <100 Low",       C["ch_cc"],   "Index"),
]
METRIC_INFO = {m[0]: m for m in METRICS}
UNIT_MAP = {
    "unemployment_rate": "%",       "initial_claims":  "claims",  "nonfarm_payrolls": "k MoM",
    "cpi_yoy":           "% YoY",   "core_cpi_yoy":    "% YoY",  "ppi_mom":           "% MoM",
    "ism_mfg":           "index",   "ism_svc":         "index",   "chicago_pmi":       "index",
    "consumer_conf":     "index",
}


# ── DATA FETCHING ─────────────────────────────────────────────────────────────

def fetch_all():
    print("  Fetching FRED series...")
    ism = read_manual_input()
    return {
        "unemployment_rate": fetch_latest("UNRATE"),
        "initial_claims":    fetch_latest("IC4WSA"),
        "nonfarm_payrolls":  fetch_payrolls(),
        "cpi_yoy":           fetch_yoy("CPIAUCSL"),
        "core_cpi_yoy":      fetch_yoy("CPILFESL"),
        "ppi_mom":           fetch_mom_pct("PPIACO"),
        "ism_mfg":           ism["ism_mfg"],
        "ism_svc":           ism["ism_svc"],
        "chicago_pmi":       ism["chicago_pmi"],
        "consumer_conf":     fetch_latest("UMCSENT"),
    }


# ── CHART DATA SHEET  (hidden — all charts pull their data from here) ─────────

CD = {k: 1 + i * 3 for i, k in enumerate([
    "unemployment_rate", "initial_claims", "nonfarm_payrolls",
    "cpi_yoy", "core_cpi_yoy", "ppi_mom",
    "ism_mfg", "ism_svc", "chicago_pmi", "consumer_conf"])}
CD["fedfunds"] = max(CD.values()) + 3

def build_chart_data(wb, data, ff_hist):
    ws = wb.create_sheet("_Chart Data")
    ws.sheet_state = "veryHidden"
    for key, (val, hist) in data.items():
        col = CD[key]
        _, name, _, _, _, _, y_lbl = METRIC_INFO[key]
        ws.cell(1, col,     f"{name} — Date")
        ws.cell(1, col + 1, f"{name} ({y_lbl})")
        for i, (date, v) in enumerate(reversed(hist[:12]), 2):
            ws.cell(i, col,     date[:7])
            ws.cell(i, col + 1, v)
    col = CD["fedfunds"]
    ws.cell(1, col,     "Fed Funds Rate — Date")
    ws.cell(1, col + 1, "Fed Funds Rate (%)")
    for i, (date, v) in enumerate(reversed(ff_hist[:48]), 2):
        ws.cell(i, col,     date[:7])
        ws.cell(i, col + 1, v)
    return ws

def _make_chart(title, y_label, ch_ws, col, n_rows, line_color,
                width=24.0, height=13.0, bar=False, y_fmt="#,##0.0"):
    # builds a chart — somehow this works, don't question it
    chart = BarChart() if bar else LineChart()
    chart.title  = title
    chart.style  = 2
    chart.width  = width
    chart.height = height

    data_ref = Reference(ch_ws, min_col=col + 1, min_row=1, max_row=n_rows + 1)
    chart.add_data(data_ref, titles_from_data=True)
    cat_ref = Reference(ch_ws, min_col=col, min_row=2, max_row=n_rows + 1)
    chart.set_categories(cat_ref)

    chart.x_axis.delete      = False
    chart.y_axis.delete      = False
    chart.x_axis.tickLblPos  = "low"

    chart.y_axis.majorGridlines = ChartLines()
    chart.x_axis.majorGridlines = ChartLines()

    chart.y_axis.numFmt  = y_fmt
    chart.y_axis.title   = y_label

    if chart.series:
        s = chart.series[0]
        if bar:
            s.graphicalProperties.solidFill = line_color
        else:
            s.graphicalProperties.line.solidFill = line_color
            s.graphicalProperties.line.width      = 22000
            s.smooth = True

    chart.legend = None
    return chart


# ── EXECUTIVE SUMMARY DASHBOARD ───────────────────────────────────────────────

def build_dashboard(wb, data, ff_hist, ch_ws):
    if _TEMPLATE_MODE and "Executive Summary" in wb.sheetnames:
        ws = wb["Executive Summary"]
    else:
        ws = wb.create_sheet("Executive Summary", 0)
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = "0D1B2A"
        set_widths(ws, [2, 22, 30, 16, 11, 42, 34, 22, 20, 44, 14])

    rh(ws, 1, 48)
    M(ws, 1, 1, 1, 11, "US MACROECONOMIC INDICATORS DASHBOARD",
      bg=C["title"], fc="FFFFFF", bold=True, size=20)
    rh(ws, 2, 18)
    M(ws, 2, 1, 2, 11,
      f"Last Refreshed: {datetime.now().strftime('%d %B %Y  |  %H:%M')}     "
      "  |     Data: FRED (St. Louis Fed)     |     "
      "ISM & Chicago PMI: Auto-extracted from PDF Reports",
      bg=C["sub"], fc="BDD7EE", size=9)
    rh(ws, 3, 8)

    rh(ws, 4, 24)
    M(ws, 4, 1, 4, 2, "SIGNAL LEGEND", bg=C["chart_hdr"], fc="FFFFFF", bold=True, size=9)
    for i, (lbl, p) in enumerate([
        ("EXPANSIONARY", "exp"), ("NEUTRAL", "neu"), ("CONTRACTIONARY", "con")
    ]):
        signal_M(ws, 4, 3 + i*3, 4, 5 + i*3, lbl, bg=C[p], fc="FFFFFF", bold=True, size=9)

    rh(ws, 5, 6); rh(ws, 6, 32)
    hdrs = ["", "Category", "Metric", "Latest Value", "Unit",
            "Economic Impact", "Fed Reserve Action",
            "Pressure on Economy", "Macro Signal", "Thresholds", "Data Date"]
    for j, h in enumerate(hdrs, 1):
        W(ws, 6, j, h, bg=C["col_hdr"], fc="FFFFFF", bold=True, size=9,
          h="left" if j in (3, 6, 7, 10) else "center", wrap=True, bs="medium")
    if not _TEMPLATE_MODE:
        ws.freeze_panes = "B7"

    pressures = []
    for i, (key, name, cat, val_fmt, thresh, _, _) in enumerate(METRICS):
        r   = 7 + i
        val, hist = data.get(key, (None, []))
        pressure, impact, fed = classify(key, val)
        pressures.append(pressure)
        bg      = C["row0"] if i % 2 == 0 else C["row1"]
        pending = (key in PDF_KEYS and val is None)
        rh(ws, r, 34)

        W(ws, r, 1, i + 1, bg=bg, fc=C["dgray"], size=9)

        cb = ws.cell(row=r, column=2, value=cat)
        cb.fill      = _fill(CAT_COLOR[cat])
        cb.font      = _font(bold=True, color="FFFFFF", size=8)
        cb.alignment = _align(h="center", wrap=True)
        cb.border    = _border_all()

        W(ws, r, 3, name, bg=bg, bold=True, size=10, h="left")

        if pending:
            cd = ws.cell(row=r, column=4, value="ADD PDF")
            cd.fill = _fill(C["input"]); cd.font = _font(bold=True, color=C["input_fc"], size=9)
            cd.alignment = _align(); cd.border = _border_all()
        else:
            cd = ws.cell(row=r, column=4, value=val)
            cd.fill = _fill(p_bg(pressure)); cd.font = _font(bold=True, color="FFFFFF", size=14)
            cd.alignment = _align(); cd.number_format = val_fmt; cd.border = _border_all()

        W(ws, r, 5, UNIT_MAP.get(key, ""), bg=bg, fc=C["dgray"], size=8, italic=True)
        W(ws, r, 6, impact, bg=bg, size=9, h="left", wrap=True)
        W(ws, r, 7, fed,    bg=bg, size=9, h="left", wrap=True)
        pressure_cell(ws, r, 8, pressure, pending=pending, size=9)
        signal_cell(  ws, r, 9, pressure, pending=pending, size=9)
        W(ws, r, 10, thresh, bg=bg, size=8, h="left", wrap=True, italic=True,
          fc=C["dgray"])
        date_s = hist[0][0] if hist else ("PDF Not Found" if key in PDF_KEYS else "N/A")
        W(ws, r, 11, date_s, bg=bg, size=8, fc=C["dgray"])

    rh(ws, 17, 10)
    overall = overall_signal(pressures)
    counts  = {p: pressures.count(p) for p in ("expansionary", "neutral", "contractionary")}
    rh(ws, 18, 44)
    signal_M(ws, 18, 1, 18, 11,
      f"OVERALL MACRO ASSESSMENT:  {p_lbl(overall)}     |     "
      f"Expansionary: {counts['expansionary']}   "
      f"Neutral: {counts['neutral']}   "
      f"Contractionary: {counts['contractionary']}",
      bg=p_bg(overall), fc="FFFFFF", bold=True, size=15)

    rh(ws, 19, 10); rh(ws, 20, 28); rh(ws, 21, 40)
    for ci, (cat, keys) in enumerate([
        ("Job Market",          [m[0] for m in METRICS if m[2] == "Job Market"]),
        ("Inflation",           [m[0] for m in METRICS if m[2] == "Inflation"]),
        ("Economic Activities", [m[0] for m in METRICS if m[2] == "Economic Activities"]),
    ]):
        cat_p = [classify(k, data.get(k, (None, []))[0])[0] for k in keys]
        co = overall_signal(cat_p)
        cs = 1 + ci * 4
        signal_M(ws, 20, cs, 20, cs+3, cat,       bg=CAT_COLOR[cat], fc="FFFFFF", bold=True, size=11)
        signal_M(ws, 21, cs, 21, cs+3, p_lbl(co), bg=p_bg(co),       fc="FFFFFF", bold=True, size=15)
    box_border(ws, 20, 1, 21, 12)

    rh(ws, 22, 12); rh(ws, 23, 32)
    M(ws, 23, 1, 23, 11, "  FEDERAL FUNDS RATE  —  Historical Context (48 Months)",
      bg=C["chart_hdr"], fc="FFFFFF", bold=True, size=11, h="left")
    n_ff = min(len(ff_hist), 48)
    if not _TEMPLATE_MODE and n_ff >= 2:
        ws.add_chart(
            _make_chart("Federal Funds Rate — Effective Rate", "Rate (%)",
                        ch_ws, CD["fedfunds"], n_ff, C["ch_ff"],
                        width=30.0, height=14.0,
                        y_fmt=CHART_Y_FMT["fedfunds"]),
            "A24")
    print("  Built: Executive Summary")


# ── CATEGORY SHEETS  (Job Market / Inflation / Economic Activities) ────────────

def build_category_sheet(wb, sheet_name, category, data, cat_col, ch_ws):
    if _TEMPLATE_MODE and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = CAT_TAB[category]
        set_widths(ws, [2, 26, 16, 24, 22, 58])

    LAST = 6

    rh(ws, 1, 44)
    M(ws, 1, 1, 1, LAST,
      f"  {category.upper()}   —   US ECONOMIC INDICATORS",
      bg=cat_col, fc="FFFFFF", bold=True, size=16)
    rh(ws, 2, 18)
    M(ws, 2, 1, 2, LAST,
      f"  Refreshed {datetime.now().strftime('%d %B %Y, %H:%M')}     |     "
      "Source: FRED API  |  ISM & Chicago PMI: Auto-extracted from PDF Reports",
      bg=C["sub"], fc="BDD7EE", size=9)
    rh(ws, 3, 10)
    if not _TEMPLATE_MODE:
        ws.freeze_panes = "A4"
    cr = 4

    cat_metrics = [(k, n, vfmt, t, ch_c, y_lbl)
                   for k, n, c, vfmt, t, ch_c, y_lbl in METRICS if c == category]

    for key, name, val_fmt, thresh, ch_c, y_lbl in cat_metrics:
        val, hist    = data.get(key, (None, []))
        pressure, impact, fed = classify(key, val)
        pending = (key in PDF_KEYS and val is None)

        rh(ws, cr, 32)
        M(ws, cr, 1, cr, LAST, f"  {name.upper()}",
          bg=cat_col, fc="FFFFFF", bold=True, size=13, h="left")
        cr += 1

        rh(ws, cr, 30); rh(ws, cr+1, 30)
        if pending:
            signal_M(ws, cr, 2, cr+1, 3, "ADD PDF TO REPORTS FOLDER",
                     bg=C["input"], fc=C["input_fc"], bold=True, size=11)
        else:
            if not _TEMPLATE_MODE:
                ws.merge_cells(start_row=cr, start_column=2, end_row=cr+1, end_column=3)
            vb = ws.cell(row=cr, column=2, value=val)
            vb.fill = _fill(p_bg(pressure)); vb.font = _font(bold=True, color="FFFFFF", size=22)
            vb.alignment = _align(); vb.number_format = val_fmt
            if not _TEMPLATE_MODE:
                box_border(ws, cr, 2, cr+1, 3)

        signal_M(ws, cr, 4, cr+1, 4,
                 p_lbl(pressure) if not pending else "PENDING",
                 bg=p_bg(pressure) if not pending else C["input"],
                 fc="FFFFFF"      if not pending else C["input_fc"],
                 bold=True, size=13)

        W(ws, cr,   5, "Economic Impact",    bg=cat_col, fc="FFFFFF", bold=True, size=9, h="left")
        W(ws, cr,   6, impact, bg=C["row1"], size=9, h="left", wrap=True)
        W(ws, cr+1, 5, "Fed Reserve Action", bg=cat_col, fc="FFFFFF", bold=True, size=9, h="left")
        W(ws, cr+1, 6, fed,    bg=C["row0"], size=9, h="left", wrap=True)
        cr += 2

        rh(ws, cr, 22)
        M(ws, cr, 2, cr, LAST,
          f"  Thresholds  |  {thresh}",
          bg=C["thresh_bg"], fc=C["thresh_fc"], size=9, h="left", bold=False)
        cr += 1

        rh(ws, cr, 22)
        for col_idx, lbl in [(2, "Date"), (3, "Value"),
                             (4, "Signal"), (5, "Pressure"), (6, "Assessment")]:
            W(ws, cr, col_idx, lbl,
              bg=cat_col, fc="FFFFFF", bold=True, size=9, bs="medium",
              h="left" if col_idx == 6 else "center")
        cr += 1

        hist_d = list(hist[:12])
        if hist_d:
            for hi, (hdate, hval) in enumerate(hist_d):
                hp, ctx, _ = classify(key, hval)
                rbg = C["row0"] if hi % 2 == 0 else C["row1"]
                rh(ws, cr, 20)

                W(ws, cr, 2, hdate, bg=rbg, size=9, fc=C["dgray"])
                hc = ws.cell(row=cr, column=3, value=hval)
                hc.fill = _fill(rbg); hc.font = _font(size=9, bold=True)
                hc.alignment = _align(); hc.number_format = val_fmt
                hc.border = _border_all()

                signal_cell(  ws, cr, 4, hp, pending=False, size=8)
                pressure_cell(ws, cr, 5, hp, pending=False, size=8)

                W(ws, cr, 6, ctx, bg=rbg, size=8, h="left", wrap=True,
                  fc=C["dgray"], italic=True)
                cr += 1
            box_border(ws, cr - len(hist_d), 2, cr - 1, 6)
        else:
            rh(ws, cr, 22)
            M(ws, cr, 2, cr, LAST,
              "No data — drop the PDF into 'ISM and MNI Reports' and re-run" if pending
              else "No data available — check FRED API key",
              bg="FFF8DC", fc="7D5A00", size=9, italic=True)
            cr += 1

        rh(ws, cr, 16); cr += 2

    rh(ws, cr, 12); cr += 1
    rh(ws, cr, 34)
    M(ws, cr, 1, cr, LAST, "  HISTORICAL TREND CHARTS  —  Last 12 Months",
      bg=C["chart_hdr"], fc="FFFFFF", bold=True, size=12, h="left")
    chart_row = cr + 1; cr += 1

    for key, name, val_fmt, thresh, ch_c, y_lbl in cat_metrics:
        _, hist = data.get(key, (None, []))
        n       = len(hist[:12])
        use_bar = (key == "nonfarm_payrolls")
        if not _TEMPLATE_MODE:
            ws.add_chart(
                _make_chart(
                    name,
                    y_lbl, ch_ws, CD[key], n if n >= 2 else 2,
                    ch_c, width=24.0, height=13.0, bar=use_bar,
                    y_fmt=CHART_Y_FMT.get(key, "#,##0.0")
                ),
                f"B{chart_row}")
        chart_row += 26

    print(f"  Built: {sheet_name}")


# ── GUIDE & SOURCES SHEET ─────────────────────────────────────────────────────

def build_notes(wb):
    ws = wb.create_sheet("Guide & Sources")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = C["dgray"]
    set_widths(ws, [2, 28, 66, 2])
    rh(ws, 1, 40)
    M(ws, 1, 1, 1, 4, "  GUIDE, DATA SOURCES & INTERPRETATION",
      bg=C["title"], fc="FFFFFF", bold=True, size=14, h="left")

    sections = [
        ("HOW TO REFRESH", C["col_hdr"], [
            ("Monthly workflow",
             "1. Replace PDFs in 'ISM and MNI Reports/' with the latest month's reports.  "
             "2. Run: py -3 economic_tracker.py  from the project folder.  "
             "3. The script extracts PMI values, updates the data log and regenerates "
             "US_Economic_Tracker.xlsx."),
            ("Automation",
             "See READ-ME.txt for full instructions on automating via Windows Task Scheduler, "
             "macOS launchd, or cron."),
            ("Portability",
             "This project is fully self-contained — all paths are relative to the script.  "
             "Copy the entire folder to any machine with Python installed and it will run "
             "without modification."),
        ]),
        ("PDF NAMING CONVENTIONS", C["ea"], [
            ("ISM Manufacturing PMI",
             "Released 1st business day of month.  Source: ismworld.org.  "
             "Filename must contain 'manuf'  (e.g. ISM Manufacturing.pdf)"),
            ("ISM Non-Manufacturing (Services) PMI",
             "Released 3rd business day of month.  Source: ismworld.org.  "
             "Filename must contain 'serv' or 'non-manuf'  (e.g. ISM Services.pdf)"),
            ("Chicago PMI  (MNI Chicago Business Barometer)",
             "Released last business day of month.  Source: MNI Indicators.  "
             "Filename must contain 'mni' or 'barometer'  "
             "(e.g. 20260228.MNI_Chicago_Bus_Barometer_Press_2026_02.pdf)"),
            ("Manual override",
             "If auto-extraction fails, open ISM_Manual_Input.xlsx and enter the value "
             "manually.  The script reads this file on every run."),
        ]),
        ("SIGNAL INTERPRETATION", C["jm"], [
            ("EXPANSIONARY  (green)",
             "The metric is consistent with a growing, healthy economy."),
            ("NEUTRAL  (amber)",
             "The metric is within the target or expected range.  No clear policy signal."),
            ("CONTRACTIONARY  (red)",
             "The metric signals economic weakness or above-target inflation."),
            ("Overall Assessment",
             "The overall signal is determined by majority: if 50%+ of metrics are "
             "contractionary the overall reading is contractionary, and vice versa."),
        ]),
        ("FRED DATA SOURCES", C["inf"], [
            ("Unemployment Rate",     "FRED: UNRATE         |  Bureau of Labor Statistics  |  Monthly"),
            ("Initial Jobless Claims","FRED: IC4WSA (4-wk avg) |  Dept of Labor           |  Weekly"),
            ("Nonfarm Payrolls",      "FRED: PAYEMS (MoM change)  |  BLS                  |  Monthly"),
            ("CPI YoY %",             "FRED: CPIAUCSL (YoY % computed)  |  BLS             |  Monthly"),
            ("Core CPI YoY %",        "FRED: CPILFESL (YoY % computed)  |  BLS             |  Monthly"),
            ("PPI MoM %",             "FRED: PPIACO (MoM % computed)  |  BLS               |  Monthly"),
            ("Consumer Confidence",   "FRED: UMCSENT  |  University of Michigan             |  Monthly"),
            ("Federal Funds Rate",    "FRED: FEDFUNDS  |  Federal Reserve                   |  Monthly  (chart only)"),
        ]),
    ]

    r = 3
    for sec_title, sec_bg, rows in sections:
        rh(ws, r, 28)
        M(ws, r, 2, r, 3, f"  {sec_title}",
          bg=sec_bg, fc="FFFFFF", bold=True, size=11, h="left")
        r += 1
        for lbl, desc in rows:
            rh(ws, r, 44)
            W(ws, r, 2, lbl,  bg=C["lgray"],   bold=True, size=9, h="left", v="top")
            W(ws, r, 3, desc, bg=C["note_bg"], size=9,    h="left", v="top", wrap=True)
            box_border(ws, r, 2, r, 3)
            r += 1
        rh(ws, r, 8); r += 1

    print("  Built: Guide & Sources")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  US Economic Indicators Tracker  v6")
    print(f"  {BASE_DIR}")
    print("=" * 60)

    print("\nStep 1: Parsing PDF reports...")
    mfg_date, mfg_val, svc_date, svc_val, chi_date, chi_val = parse_reports_folder()

    print("\nStep 2: Updating data log...")
    update_manual_input(mfg_date, mfg_val, svc_date, svc_val, chi_date, chi_val)

    print("\nStep 3: Fetching FRED data...")
    data    = fetch_all()
    ff_hist = fetch_fedfunds(48)

    print("\nStep 4: Building workbook...")
    global _TEMPLATE_MODE

    if os.path.exists(TEMPLATE_FILE):
        # Template mode — load the user's formatted workbook and update data cells only.
        # Charts update automatically since they pull from _Chart Data (rebuilt below).
        # Formatting, colours, column widths, fonts — all untouched.
        print("  Template found — updating data cells, preserving your formatting...")
        _TEMPLATE_MODE = True
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        if "_Chart Data" in wb.sheetnames:
            del wb["_Chart Data"]   # always rebuild this so charts get fresh data
    else:
        # No template found — build the whole workbook from scratch as normal.
        # To use template mode: run once, copy the output as US_Economic_Tracker_TEMPLATE.xlsx,
        # format it however you like, and subsequent runs will update into your template.
        print("  No template found — building workbook from scratch...")
        _TEMPLATE_MODE = False
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    ch_ws = build_chart_data(wb, data, ff_hist)
    build_dashboard(wb, data, ff_hist, ch_ws)
    build_category_sheet(wb, "Job Market",         "Job Market",          data, C["jm"],  ch_ws)
    build_category_sheet(wb, "Inflation",           "Inflation",           data, C["inf"], ch_ws)
    build_category_sheet(wb, "Economic Activities", "Economic Activities", data, C["ea"],  ch_ws)
    if not _TEMPLATE_MODE:
        build_notes(wb)

    wb.save(OUTPUT_FILE)
    print(f"\n{'='*60}")
    print(f"  Done.  Saved: {OUTPUT_FILE}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
