import os
import re
from datetime import datetime, date
from dateutil.parser import parse
from pptx import Presentation
import pandas as pd
import streamlit as st
from io import BytesIO

# === Regex Patterns ===
PATTERNS = {
    "docket": re.compile(r"\b\d{4,}\.\d{4}(?:-\w+)?\b"),
    "application": re.compile(r"\d{2}/\d{3},\d{3}"),
    "pct": re.compile(r"PCT/US\d{2}/\d{6}"),
    "wipo": re.compile(r"[A-Z]{2}\d{4}/\d{6}"),
    "date": re.compile(r"\b(?:\d{1,2}[/-])?\d{1,2}[/-]\d{2,4}\b")
}

# === Date & Action Extraction ===
def extract_action_and_dates_from_raw_text(raw_text):
    raw_lines = raw_text.splitlines()
    due_date = None
    extension_date = None
    action = ""

    for line in raw_lines:
        line_lower = line.lower()

        if "due" in line_lower:
            try:
                due_pos = line_lower.find("due")
                after_due = line[due_pos:]
                due_match = PATTERNS["date"].search(after_due)
                if due_match:
                    due_date_raw = due_match.group()
                    parsed_due = parse(due_date_raw, dayfirst=False, fuzzy=True)

                    if parsed_due.date() >= date.today():
                        due_date = parsed_due.strftime("%m/%d/%Y")

                        # Action is only the content before the due date on the same line
                        action = line.split(due_date_raw)[0].strip()

                        # Handle extension
                        if "ext" in line_lower or "extension" in line_lower:
                            ext_pos = line_lower.find("ext")
                            after_ext = line[ext_pos:]
                            ext_match = PATTERNS["date"].search(after_ext)
                            if ext_match:
                                extension_date_raw = ext_match.group()
                                parsed_ext = parse(extension_date_raw, dayfirst=False, fuzzy=True)
                                extension_date = parsed_ext.strftime("%m/%d/%Y")

            except:
                continue

    return action, due_date, extension_date

# === Entry Extraction per Textbox ===
def extract_entries_from_textbox(text, slide_index, file_name):
    entries = []
    text = text.replace("\u2028", "\n")  # Handle soft line breaks

    docket_match = PATTERNS["docket"].search(text)
    if not docket_match:
        return []

    docket_number = docket_match.group()
    application_match = PATTERNS["application"].search(text)
    application_number = application_match.group() if application_match else ""

    pct_match = PATTERNS["pct"].search(text)
    pct_number = pct_match.group() if pct_match else ""

    wipo_match = PATTERNS["wipo"].search(text)
    wipo_number = wipo_match.group() if wipo_match else ""

    action, due_date, extension_date = extract_action_and_dates_from_raw_text(text)
    if not due_date:
        return []

    entries.append({
        "Slide": slide_index,
        "Textbox Content": text,
        "Docket Number": docket_number,
        "Application Number": application_number,
        "PCT Number": pct_number,
        "WIPO Number": wipo_number,
        "Action": action,
        "Due Date": due_date,
        "Extension Date": extension_date,
        "File Name": file_name
    })

    return entries

# === PPTX Parser ===
def extract_from_pptx(file):
    all_entries = []
    prs = Presentation(file)

    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                raw_text = shape.text.strip()
                entries = extract_entries_from_textbox(raw_text, i + 1, file.name)
                all_entries.extend(entries)

    return pd.DataFrame(all_entries)

# === Streamlit UI ===
st.set_page_config(layout="wide")
st.title("\U0001F4CA DocketPoint")

st.sidebar.image("firm_logo.png", use_container_width=True)
st.sidebar.markdown("---")
st.sidebar.markdown(
    """
    **About DocketPoint**

    This tool extracts docket numbers, application numbers, and due dates from PowerPoint files.  
    It helps organize patent prosecution data and export it to Excel for streamlined docket tracking.  
    Use the slider to filter by due date range.
    """
)
st.sidebar.markdown("---")

ppt_files = st.file_uploader("Upload one or more PowerPoint (.pptx) files", type="pptx", accept_multiple_files=True)
months_back = st.slider("Include due dates up to this many months in the past:", 0, 24, 0)

if ppt_files:
    all_dfs = []
    for ppt_file in ppt_files:
        df = extract_from_pptx(ppt_file)
        if df.empty:
            st.warning(f"⚠️ No extractable data found in {ppt_file.name}.")
            continue
        df["Filename"] = ppt_file.name
        all_dfs.append(df)

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)

        # Filter by due date range
        final_df["Due Date Parsed"] = pd.to_datetime(final_df["Due Date"], errors="coerce")
        cutoff = pd.Timestamp.now() - pd.DateOffset(months=months_back)
        final_df = final_df[final_df["Due Date Parsed"] >= cutoff].copy()
        final_df = final_df.sort_values(by="Due Date Parsed").drop(columns=["Due Date Parsed"])

        st.success(f"✅ Extracted {len(final_df)} entries from {len(all_dfs)} file(s).")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("\U0001F4E5 Download Excel", output, file_name="extractedDocket.xlsx")
