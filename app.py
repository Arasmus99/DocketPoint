import re
import pandas as pd
from pptx import Presentation
from datetime import date, timedelta
from dateutil.parser import parse
import streamlit as st
from io import BytesIO

# === Regex Patterns ===
PATTERNS = {
    "docket_number": re.compile(
        r"\b\d{4}-[A-Z]{2,}-\d{5}-\d{2}\b"
        r"|\b\d{5}-\d{2}\b"
        r"|\b\d{5}-\d{4}-\d{2}[A-Z]{2,4}\b"
        r"|\b\d{4}[.-]\d{4}-?[A-Z]{2}\d*\b"
        r"|\b\d{4}-\d{4}-[A-Z]{3}\b"
    ),
    "application_number": re.compile(r"\b\d{2}/\d{3}[,]?\d{3}\s+[A-Z]{2}\b"),
    "alt_application_number": re.compile(
        r"\b[Pp]\d{11}\s+[A-Z]{2}-\w{1,4}\b"
        r"|\b\d{5,12}(?:[.,]\d+)?\s+[A-Z]{2,3}\b"
    ),
    "pct_number": re.compile(r"PCT/[A-Z]{2}\d{4}/\d{6}"),
    "wipo_number": re.compile(r"\bWO\d{4}/\d{6}\b"),
    "date": re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b")
}

SKIP_PHRASES = ["PENDING", "ABANDONED", "WITHDRAWN", "GRANTED", "ISSUED", "STRUCTURE"]

def extract_texts_from_shape_recursive(shape):
    texts = []
    if shape.shape_type == 6:
        for shp in shape.shapes:
            texts.extend(extract_texts_from_shape_recursive(shp))
    else:
        text = extract_text_from_shape(shape)
        if text:
            texts.append(text)
    return texts

def extract_text_from_shape(shape):
    if shape.has_text_frame:
        return shape.text.strip()
    return ""

def should_include(text):
    upper_text = text.upper()
    return not any(phrase in upper_text for phrase in SKIP_PHRASES)

def get_earliest_due_date(dates_str):
    if not isinstance(dates_str, str):
        return pd.NaT
    try:
        dates = [parse(d.strip(), dayfirst=False, fuzzy=True) for d in dates_str.split(";") if d.strip()]
        return min(dates) if dates else pd.NaT
    except:
        return pd.NaT

def process_extensions(text):
    ext_keywords = ["ext", "extension"]
    for line in text.splitlines():
        lower = line.lower()
        if any(kw in lower for kw in ext_keywords):
            dates = PATTERNS["date"].findall(line)
            if len(dates) >= 2:
                return parse(dates[1], dayfirst=False, fuzzy=True).strftime("%m/%d/%Y")
    return None

def extract_entries_from_textbox(text):
    entries = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    entry = {
        "docket_number": None,
        "application_number": None,
        "pct_number": None,
        "wipo_number": None,
        "due_dates": [],
        "raw_text": "\n".join(lines),
        "Extension": process_extensions(text)
    }

    for line in lines:
        clean_line = re.sub(r"[^0-9A-Za-z/,\.\s-]", "", line.replace(" /,", "/").replace("/", "/").replace(",,", ",").replace(" /", "/"))
        clean_line = clean_line.replace(",", "")

        if not entry["docket_number"] and PATTERNS["docket_number"].search(clean_line):
            entry["docket_number"] = PATTERNS["docket_number"].search(clean_line).group(0)
        if not entry["pct_number"] and PATTERNS["pct_number"].search(clean_line):
            entry["pct_number"] = PATTERNS["pct_number"].search(clean_line).group(0)
        if not entry["application_number"] and PATTERNS["application_number"].search(clean_line):
            entry["application_number"] = PATTERNS["application_number"].search(clean_line).group(0)
        elif not entry["application_number"] and PATTERNS["alt_application_number"].search(clean_line):
            entry["application_number"] = PATTERNS["alt_application_number"].search(clean_line).group(0)
        if not entry["wipo_number"] and PATTERNS["wipo_number"].search(clean_line):
            entry["wipo_number"] = PATTERNS["wipo_number"].search(clean_line).group(0)

        for match in PATTERNS["date"].findall(clean_line):
            try:
                parsed = parse(match, dayfirst=False, fuzzy=True)
                entry["due_dates"].append(parsed.strftime("%m/%d/%Y"))
            except:
                continue

    if not (entry["docket_number"] or entry["application_number"] or entry["pct_number"] or entry["wipo_number"]):
        return []

    entry["Textbox Content"] = entry["raw_text"]

    if entry["due_dates"]:
        entries.append(entry)

    return entries

def split_due_dates(entry_dict):
    due_dates = entry_dict["due_dates"]
    if not due_dates:
        return []

    results = []
    for d in due_dates:
        new_entry = {
            "Slide": entry_dict["Slide"],
            "Textbox Content": entry_dict["Textbox Content"],
            "Docket Number": entry_dict["docket_number"],
            "Application Number": entry_dict["application_number"],
            "PCT Number": entry_dict["pct_number"],
            "WIPO Number": entry_dict["wipo_number"],
            "Extension": entry_dict["Extension"],
            "Due Dates": "; ".join(due_dates),
            "Due Date": d
        }
        results.append(new_entry)
    return results

def extract_from_pptx(upload):
    prs = Presentation(upload)
    raw_entries = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            texts = extract_texts_from_shape_recursive(shape)
            for text in texts:
                if not should_include(text):
                    continue
                for entry in extract_entries_from_textbox(text):
                    entry["Slide"] = slide_num
                    raw_entries.append(entry)

    all_split_entries = []
    for entry in raw_entries:
        all_split_entries.extend(split_due_dates(entry))

    df = pd.DataFrame(all_split_entries)
    if df.empty:
        return pd.DataFrame(columns=["Slide", "Textbox Content", "Docket Number", "Application Number", "PCT Number", "WIPO Number", "Due Date", "Extension"])

    today = date.today()
    df["Due Date Parsed"] = pd.to_datetime(df["Due Date"], errors="coerce")
    df = df[df["Due Date Parsed"] >= pd.to_datetime(today)]
    df = df.drop(columns=["Due Date Parsed"])
    return df

# === Streamlit UI ===
st.title("\U0001F4CA DocketPoint")
st.sidebar.image("firm_logo.png", use_container_width=True)
st.sidebar.markdown("---")
st.sidebar.markdown("""
**About DocketPoint**

This tool extracts docket numbers, application numbers, and due dates from PowerPoint files.  
It helps organize patent prosecution data and export it to Excel for streamlined docket tracking.  
Use the slider to filter by due date range.
""")
st.sidebar.markdown("---")

ppt_files = st.file_uploader("Upload one or more PowerPoint (.pptx) files", type="pptx", accept_multiple_files=True)
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
        final_df["Earliest Due Date"] = final_df["Due Date"].apply(get_earliest_due_date)
        final_df = final_df.sort_values(by="Earliest Due Date", ascending=True).drop(columns=["Earliest Due Date"])

        st.success(f"✅ Extracted {len(final_df)} entries from {len(all_dfs)} file(s).")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("\U0001F4E5 Download Excel", output, file_name="combined_extracted_data.xlsx")
