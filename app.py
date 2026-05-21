import re
from io import BytesIO
from datetime import date, timedelta, datetime

import pandas as pd
import streamlit as st
from dateutil.parser import parse as parse_date
from pptx import Presentation
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ===========================================================================
#  1) EXTRACTION ENGINE  (was docket_extract.py)
# ===========================================================================
"""
docket_extract.py
-----------------
Extraction engine for patent "Case Structure" PowerPoint decks.

Design notes (derived from the actual deck + the deck's own "Key" slide):
  * Each case lives in its own text box. A box's first docket-looking token is
    the ATTY REF / docket number, e.g. ``01394-0005-00EP``.
  * The 2-4 trailing letters of the docket's last hyphen segment give the
    jurisdiction (country) -> ``EP``.
  * After the docket comes the CLIENT CODE (``15080-101EP1`` / ``107USP1``),
    then the application number, filing date, grant number, due dates, notes.

Because the client code and docket both contain digits, we strip those tokens
out *before* searching for the application number, then apply a
country-specific pattern. Every pattern below was checked against the real
formats found in the deck.
"""


# --------------------------------------------------------------------------- #
#  Docket / client-code patterns
# --------------------------------------------------------------------------- #

# 01394-0005-00EP   or   01394-0003-00MO-01CN
DOCKET_RE = re.compile(
    r"\b\d{5}-\d{4}-\d{2}[A-Z]{2,4}(?:-\d{2}[A-Z]{2,4})?\b"
)

# Client code tokens that we want to ignore when hunting for an app number:
#   15080-101EP1, 15080-004US3-CON2, 15080-101WO1US2, 15040-004CN2-MO, 107USP1
#
# Key constraint: a client code is <4-5 digits>-<3 digits> followed by a LETTER
# (the country/type segment). Requiring that trailing letter is what stops the
# pattern from swallowing pure-numeric application numbers such as the Japanese
# "2016-502307" or the Korean "...-2025-7005473".
CLIENT_CODE_RES = [
    re.compile(r"\b\d{4,5}-\d{3}[A-Za-z][A-Za-z0-9-]*"),   # 15080-101EP1, 15040-004CN2-MO
    re.compile(r"\b\d{3}[A-Z]{2,3}P?\d*\b"),               # 107USP1, 107USP2
]


# --------------------------------------------------------------------------- #
#  Country-specific APPLICATION-number patterns (validated against the deck)
# --------------------------------------------------------------------------- #
COUNTRY_APP_RES = {
    "US": re.compile(r"\b\d{2}/\d{3},?\d{3}\b"),          # 61/793,993 ; 14/211,002
    "EP": re.compile(r"\b\d{8}\.\d\b"),                   # 14764430.6
    "JP": re.compile(r"\b\d{4}-\d{6}\b"),                 # 2016-502307
    "KR": re.compile(r"\b10-\d{4}-\d{7}\b"),              # 10-2017-7008850
    "CN": re.compile(r"(?<!\d)\d{12}\.\s?[\dX](?!\d)"),   # 201580054350.9 / .X / ZL-prefixed
    "IN": re.compile(r"\b\d{12}\b"),                      # 201747008733
    "AU": re.compile(r"\b\d{10}\b"),                      # 2015317972
    "TW": re.compile(r"\b\d{9}\b"),                       # 112127315
    "CL": re.compile(r"\b\d{9}\b"),                       # 202500166
    "EA": re.compile(r"\b\d{9}\b"),                       # 201790628
    "CA": re.compile(r"\b\d,?\d{3},?\d{3}\b"),            # 2,961,200 / 3262284
    "IL": re.compile(r"\b\d{6}\b"),                       # 250836
    "NZ": re.compile(r"\b\d{6,7}\b"),                     # 818923
    "SG": re.compile(r"\b\d{11}[A-Z]\b"),                 # 11201701957X
    "ZA": re.compile(r"\b\d{4}/\d{5}\b"),                 # 2025/01545
    "MX": re.compile(r"\bMX/a/\d{4}/\d{6}\b"),            # MX/a/2017/002789
    "CO": re.compile(r"\bNC\d{4}/\d{6}\b"),               # NC2025/001854
    "AR": re.compile(r"\bP\d{9}\b"),                      # P230101911
    "HK": re.compile(r"(?<!\d)\d{8,11}\.\d(?!\d)"),       # 17113734.5 / 62024096696.5
    "BR": re.compile(r"BR[\d\s]+?\.\d"),                  # BR112017005111.7 / BR 12 2022 023284.1
    "MO": re.compile(r"\bJ/\d+\b"),                       # J/008517 (Macau)
}

# Order used for the "try everything" fallback when the docket country is
# unknown or its pattern misses. More-specific patterns come first so they win.
FALLBACK_ORDER = [
    "MX", "CO", "AR", "BR", "KR", "ZA", "SG", "US", "MO",
    "CN", "HK", "EP", "JP", "IN", "AU", "TW", "CA", "IL", "NZ",
]

PCT_RE = re.compile(r"\bPCT/[A-Z]{2}\d{4}/\d{6}\b")
# WO2014/143643 ; WO2024020127 (no slash) ; WO 2025/160227 A1
WIPO_RE = re.compile(r"\bWO\s?\d{4}/?\d{6}(?:\s?A\d)?\b")

# US grant numbers like 9,629,860 / 10,195,222 (used for the bonus "patent no." column)
US_GRANT_RE = re.compile(r"\b\d{1,2},\d{3},\d{3}\b")

DATE_RE = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b")
DUE_LINE_RE = re.compile(r"\b(due|by)\b", re.IGNORECASE)

STATUS_RE = re.compile(r"\b(ABN|ABANDONED|WITHDRAWN|PENDING|GRANTED|ISSUED|EXPIRED)\b",
                       re.IGNORECASE)

# Boxes that are pure headers / page chrome (no real case data)
SKIP_EXACT = {"structure", "case structures", "key", "priorities",
              "appendix \u2013 blow-ups", "appendix - blow-ups"}


# --------------------------------------------------------------------------- #
#  Helpers
# --------------------------------------------------------------------------- #
def get_country(docket):
    """EP from 01394-0005-00EP ; CN from 01394-0003-02CN ; MO special-cased."""
    if not docket:
        return None
    last = docket.split("-")[-1]
    m = re.match(r"\d*([A-Z]{2,4})$", last)
    country = m.group(1) if m else None
    # 01394-0003-00MO-01CN -> Macau application even though last segment is CN
    if "MO" in docket:
        return "MO"
    return country


def _strip_known_tokens(text, docket):
    """Remove docket + client-code tokens so they don't masquerade as app #s."""
    out = text
    if docket:
        out = out.replace(docket, " ")
    for rx in CLIENT_CODE_RES:
        out = rx.sub(" ", out)
    return out


def _clean_match(value):
    if value is None:
        return None
    # collapse internal whitespace (e.g. CN "201580054350. 9" -> "201580054350.9")
    return re.sub(r"\s+", "", value.strip())


def find_application_number(text, docket):
    """Return (application_number, country)."""
    search_text = _strip_known_tokens(text, docket)
    country = get_country(docket)

    # PCT national-phase parent: the PCT number is the identifier, not a
    # separate application serial. Reported in the PCT column instead.
    if country == "PCT":
        return None, country

    # 1) Use the jurisdiction's own pattern. For a known country we trust only
    #    that pattern -- guessing with another country's (looser) pattern tends
    #    to grab grant numbers, so we'd rather leave the field blank.
    if country in COUNTRY_APP_RES:
        m = COUNTRY_APP_RES[country].search(search_text)
        return (_clean_match(m.group(0)) if m else None), country

    # 2) Unknown jurisdiction -> try every pattern in priority order.
    for key in FALLBACK_ORDER:
        m = COUNTRY_APP_RES[key].search(search_text)
        if m:
            return _clean_match(m.group(0)), country

    return None, country


def find_pct(text):
    m = PCT_RE.search(text)
    return m.group(0) if m else None


def find_wipo(text):
    m = WIPO_RE.search(text)
    if not m:
        return None
    # Normalize to "WO YYYY/NNNNNN" (keep optional kind code like A1)
    raw = re.sub(r"\s+", "", m.group(0))           # WO2025/160227A1
    raw = re.sub(r"^WO", "", raw)
    kind = ""
    km = re.search(r"(A\d)$", raw)
    if km:
        kind = " " + km.group(1)
        raw = raw[: km.start()]
    return f"WO {raw}{kind}".strip()


def _norm_date(raw):
    """Return MM/DD/YYYY or None."""
    try:
        d = parse_date(raw, dayfirst=False, fuzzy=False)
        return d.strftime("%m/%d/%Y")
    except Exception:
        return None


def find_dates(lines):
    """
    Split dates into a single filing date and a list of due-date deadlines.

    * A *due date* is any date that sits on a line containing 'due'/'by'.
    * The *filing date* is the first non-due date (typically next to the app #).
    Returns (filing_date, [ {action, date}, ... ], [undated_action, ...]).
    """
    filing = None
    deadlines = []
    undated = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        dates_on_line = DATE_RE.findall(line)
        is_due = bool(DUE_LINE_RE.search(line))

        if is_due:
            if dates_on_line:
                for raw in dates_on_line:
                    nd = _norm_date(raw)
                    if not nd:
                        continue
                    # action = text up to the date, tidied
                    idx = line.find(raw)
                    action = line[:idx].strip(" :-\u2013") or line.strip()
                    deadlines.append({"action": action, "date": nd})
            else:
                # Undated action item, e.g. "Assignment due".
                # Require the explicit word "due" so a stray "...by..." with no
                # date (common in narrative notes) isn't logged as an action.
                if re.search(r"\bdue\b", line, re.IGNORECASE):
                    undated.append(line)
        else:
            for raw in dates_on_line:
                nd = _norm_date(raw)
                if nd and filing is None:
                    filing = nd
                    break

    return filing, deadlines, undated


# --------------------------------------------------------------------------- #
#  Box-level parsing
# --------------------------------------------------------------------------- #
def parse_box(text, slide_num):
    """Return a case dict for one text box, or None if it holds no case data."""
    raw = text.strip()
    if raw.lower() in SKIP_EXACT:
        return None

    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if not lines:
        return None

    docket_m = DOCKET_RE.search(raw)
    docket = docket_m.group(0) if docket_m else None

    pct = find_pct(raw)
    wipo = find_wipo(raw)

    # Skip boxes that have no identifiers at all (titles, page numbers, etc.)
    if not docket and not pct and not wipo:
        return None

    app_no, country = (None, None)
    if docket:
        app_no, country = find_application_number(raw, docket)
    elif pct:
        country = "PCT"

    filing, deadlines, undated = find_dates(lines)

    status_m = STATUS_RE.search(raw)
    status = status_m.group(0).upper() if status_m else ""

    return {
        "slide": slide_num,
        "docket": docket,
        "country": country,
        "application_number": app_no,
        "pct_number": pct,
        "wipo_number": wipo,
        "filing_date": filing,
        "status": status,
        "deadlines": deadlines,
        "undated_actions": undated,
        "raw_text": raw,
    }


# --------------------------------------------------------------------------- #
#  Shape walking
# --------------------------------------------------------------------------- #
def iter_box_texts(shape):
    """Yield text for a shape, recursing into groups."""
    if shape.shape_type == 6:  # GroupShape
        for child in shape.shapes:
            yield from iter_box_texts(child)
    elif getattr(shape, "has_text_frame", False) and shape.text.strip():
        yield shape.text.strip()


def extract_cases(pptx_source):
    """Parse a pptx (path or file-like) -> list of case dicts."""
    prs = Presentation(pptx_source)
    cases = []
    for slide_num, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            for text in iter_box_texts(shape):
                case = parse_box(text, slide_num)
                if case:
                    cases.append(case)
    return cases

# ===========================================================================
#  2) EXCEL BUILDER  (was build_excel.py)
# ===========================================================================
"""
build_excel.py
--------------
Turn extracted cases into a formatted two-sheet workbook:
  * "Deadlines" – one row per dated deadline (calendar-ready), sorted by date.
  * "All Cases" – one row per case with every identifier we pulled.

Used by both the CLI (run_extract.py) and the Streamlit app.
"""


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
BODY_FONT = Font(name="Arial", size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _join_deadlines(case):
    parts = [f"{d['action']}: {d['date']}" for d in case["deadlines"]]
    parts += [u for u in case["undated_actions"]]
    return "; ".join(parts)


def cases_to_rows(cases, client):
    """Build (deadline_rows, case_rows) as lists of dicts."""
    case_rows, deadline_rows, seen = [], [], set()

    for c in cases:
        case_rows.append({
            "Client": client,
            "Slide": c["slide"],
            "Docket Number": c["docket"],
            "Country": c["country"],
            "Application Number": c["application_number"],
            "PCT Number": c["pct_number"],
            "WIPO Number": c["wipo_number"],
            "Filing Date": c["filing_date"],
            "Status": c["status"],
            "Due Dates / Actions": _join_deadlines(c),
        })
        for d in c["deadlines"]:
            key = (client, c["docket"], d["action"], d["date"])
            if key in seen:
                continue
            seen.add(key)
            deadline_rows.append({
                "Due Date": d["date"],
                "Action": d["action"],
                "Docket Number": c["docket"],
                "Country": c["country"],
                "Application Number": c["application_number"]
                                      or c["pct_number"] or c["wipo_number"],
                "Client": client,
                "Slide": c["slide"],
            })

    deadline_rows.sort(key=lambda r: datetime.strptime(r["Due Date"], "%m/%d/%Y"))
    return deadline_rows, case_rows


def _write_sheet(ws, rows, columns, date_cols=()):
    ws.append(columns)
    for col_idx, _ in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for row in rows:
        ws.append([row.get(c, "") for c in columns])

    # body styling + column widths
    widths = {c: len(c) for c in columns}
    for r in range(2, ws.max_row + 1):
        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=col_idx)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(col_name == "Due Dates / Actions"))
            val = cell.value
            if val not in (None, ""):
                widths[col_name] = max(widths[col_name], min(len(str(val)), 60))

    for col_idx, col_name in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = widths[col_name] + 2

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"


def build_workbook(cases, client):
    deadline_rows, case_rows = cases_to_rows(cases, client)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Deadlines"
    _write_sheet(ws1, deadline_rows,
                 ["Due Date", "Action", "Docket Number", "Country",
                  "Application Number", "Client", "Slide"])

    ws2 = wb.create_sheet("All Cases")
    _write_sheet(ws2, case_rows,
                 ["Client", "Slide", "Docket Number", "Country",
                  "Application Number", "PCT Number", "WIPO Number",
                  "Filing Date", "Status", "Due Dates / Actions"])

    return wb


def workbook_bytes(cases, client):
    buf = BytesIO()
    build_workbook(cases, client).save(buf)
    buf.seek(0)
    return buf

# ===========================================================================
#  3) STREAMLIT UI  (was app.py)
# ===========================================================================
"""
DocketPoint — PowerPoint case-structure extractor
==================================================
Run locally with:   streamlit run app.py

Keep app.py, docket_extract.py, and build_excel.py in the same folder.

Upload one or more "Case Structure" .pptx decks; the app extracts every case's
docket / application / PCT / WIPO numbers, filing dates, and dated deadlines,
shows them in two tables, and lets you download a formatted Excel workbook
(Deadlines sheet + All Cases sheet).
"""


st.set_page_config(page_title="DocketPoint", page_icon="\U0001F4CA", layout="wide")
st.title("\U0001F4CA DocketPoint")

st.sidebar.markdown("### About DocketPoint")
st.sidebar.markdown(
    "Extracts docket, application, PCT and WIPO numbers plus filing dates and "
    "dated deadlines from patent *Case Structure* PowerPoint decks, and exports "
    "them to Excel.\n\n"
    "**Deadlines** sheet = one row per dated deadline (calendar-ready).\n\n"
    "**All Cases** sheet = one row per case with every identifier."
)
st.sidebar.markdown("---")

ppt_files = st.file_uploader(
    "Upload one or more PowerPoint (.pptx) files",
    type="pptx",
    accept_multiple_files=True,
)

months_back = st.slider(
    "On the Deadlines view, include deadlines due this many months in the past:",
    0, 36, 0,
)

if not ppt_files:
    st.info("Upload a .pptx case-structure deck to begin.")
    st.stop()

cutoff = date.today() - timedelta(days=int(30.4 * months_back))

all_cases = []
for f in ppt_files:
    client = f.name.replace(".pptx", "")
    cases = extract_cases(f)
    if not cases:
        st.warning(f"\u26A0\uFE0F No extractable cases found in {f.name}.")
        continue
    for c in cases:
        c["_client"] = client
    all_cases.append((client, cases))

if not all_cases:
    st.stop()

# Build combined tables across all uploaded files
deadline_rows, case_rows = [], []
for client, cases in all_cases:
    d_rows, c_rows = cases_to_rows(cases, client)
    deadline_rows += d_rows
    case_rows += c_rows

deadlines_df = pd.DataFrame(deadline_rows)
cases_df = pd.DataFrame(case_rows)

# Apply the past-deadline filter to the Deadlines view only
if not deadlines_df.empty:
    parsed = pd.to_datetime(deadlines_df["Due Date"], format="%m/%d/%Y", errors="coerce")
    deadlines_df = deadlines_df[parsed.dt.date >= cutoff].reset_index(drop=True)

st.success(
    f"\u2705 Extracted {len(cases_df)} cases and "
    f"{len(deadlines_df)} deadlines from {len(all_cases)} file(s)."
)

tab1, tab2 = st.tabs([f"\U0001F5D3\uFE0F Deadlines ({len(deadlines_df)})",
                      f"\U0001F4C1 All Cases ({len(cases_df)})"])
with tab1:
    st.dataframe(deadlines_df, use_container_width=True, hide_index=True)
with tab2:
    st.dataframe(cases_df, use_container_width=True, hide_index=True)

# Build a single workbook covering all uploaded files
combined = []
for _, cases in all_cases:
    combined.extend(cases)
client_label = all_cases[0][0] if len(all_cases) == 1 else "Combined"

buf = BytesIO()
build_workbook(combined, client_label).save(buf)
buf.seek(0)

st.download_button(
    "\U0001F4E5 Download Excel",
    buf,
    file_name=f"{client_label}_Docket_Extract.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
