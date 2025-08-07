import os 
import re
from datetime import datetime, date
from dateutil.parser import parse
from pptx import Presentation
import pandas as pd
import streamlit as st
from io import BytesIO

PATTERNS = {
    "docket": re.compile(r"\b\d{4,}\.\d{4}(?:-\w+)?\b"),
    "application": re.compile(r"\d{2}/\d{3},\d{3}"),
    "pct": re.compile(r"PCT/US\d{2}/\d{6}"),
    "wipo": re.compile(r"[A-Z]{2}\d{4}/\d{6}"),
    "date": re.compile(r"\b(?:\d{1,2}[/-])?\d{1,2}[/-]\d{2,4}\b")
}

def extract_action_and_dates_from_raw_text(raw_text):
    raw_lines = raw_text.splitlines()
    future_dates = []
    action = ""
    extension_date = None

    for line in raw_lines:
        for date_match in PATTERNS["date"].findall(line):
            try:
                parsed = parse(date_match, dayfirst=False, fuzzy=True)
                if parsed.date() >= date.today():
                    future_dates.append((parsed.strftime("%m/%d/%Y"), line, date_match))
            except:
                continue

    if len(future_dates) == 0:
        return "", None, None

    # If multiple dates found, only treat those with 'due' as due dates
    if len(future_dates) > 1:
        for formatted_date, line, date_match in future_dates:
            line_lower = line.lower()
            if "due" in line_lower:
                due_pos = line_lower.find("due")
                after_due = line[due_pos:]
                if date_match in after_due:
                    due_date = formatted_date
                    action = line.split(date_match)[0].strip()

                    # Check for extension after this date
                    if "ext" in line_lower or "extension" in line_lower:
                        ext_pos = line_lower.find("ext")
                        after_ext = line[ext_pos:]
                        ext_match = PATTERNS["date"].search(after_ext)
                        if ext_match:
                            try:
                                ext_parsed = parse(ext_match.group(), dayfirst=False, fuzzy=True)
                                extension_date = ext_parsed.strftime("%m/%d/%Y")
                            except:
                                pass

                    return action, due_date, extension_date

    # Fallback: just take the first future-looking date
    first_date, first_line, first_raw = future_dates[0]
    action = first_line.split(first_raw)[0].strip()
    return action, first_date, None

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

def extract_from_pptx(pptx_file):
    all_entries = []
    prs = Presentation(pptx_file)

    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                raw_text = shape.text.strip()
                entries = extract_entries_from_textbox(raw_text, i + 1, pptx_file.name)
                all_entries.extend(entries)

    return pd.DataFrame(all_entries)

# === Streamlit UI ===
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
        final_df["Earliest Due Date"] = pd.to_datetime(final_df["Due Date"], errors="coerce")
        final_df = final_df.sort_values(by="Earliest Due Date", ascending=True).drop(columns=["Earliest Due Date"])

        st.success(f"✅ Extracted {len(final_df)} entries from {len(all_dfs)} file(s).")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("\U0001F4E5 Download Excel", output, file_name="extractedDocket.xlsx")
