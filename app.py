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

def extract_extension_and_due_dates(text):
    due_dates = []
    extension = ""
    lines = text.splitlines()
    for line in lines:
        if "ext" in line.lower() or "extension" in line.lower():
            all_dates = PATTERNS["date"].findall(line)
            if len(all_dates) >= 2:
                try:
                    due = parse(all_dates[0], dayfirst=False, fuzzy=True)
                    ext = parse(all_dates[1], dayfirst=False, fuzzy=True)
                    due_dates.append(due.strftime("%m/%d/%Y"))
                    extension = ext.strftime("%m/%d/%Y")
                    continue
                except:
                    continue
        for match in PATTERNS["date"].findall(line):
            try:
                parsed = parse(match, dayfirst=False, fuzzy=True)
                due_dates.append(parsed.strftime("%m/%d/%Y"))
            except:
                continue
    return due_dates, extension

def date_split(due_dates_str, raw_text, base_entry):
    if not isinstance(due_dates_str, str):
        return [base_entry]

    due_dates = [d.strip() for d in due_dates_str.split(";") if d.strip()]
    if len(due_dates) <= 1:
        return [base_entry]

    results = []
    lines = raw_text.splitlines()
    for due_date in due_dates:
        action = ""
        for line in lines:
            if due_date in line:
                action = line.strip()
                break
        entry_copy = base_entry.copy()
        entry_copy["Due Date"] = due_date
        entry_copy["Action"] = action
        results.append(entry_copy)
    return results

def extract_entries_from_textbox(text):
    entries = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    entry = {
        "docket_number": None,
        "application_number": None,
        "pct_number": None,
        "wipo_number": None,
        "due_dates": [],
        "extension": "",
        "raw_text": "\n".join(lines)
    }

    for line in lines:
        clean_line = line.replace(" /,", "/").replace("/", "/").replace(",,", ",").replace(" /", "/")
        clean_line = re.sub(r"[^0-9A-Za-z/,.\s-]", "", clean_line)
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

    entry["due_dates"], entry["extension"] = extract_extension_and_due_dates(entry["raw_text"])

    if not (entry["docket_number"] or entry["application_number"] or entry["pct_number"] or entry["wipo_number"]):
        return []

    if entry["due_dates"]:
        entries.append(entry)

    return entries

def extract_from_pptx(upload):
    prs = Presentation(upload)
    results = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            texts = extract_texts_from_shape_recursive(shape)
            for text in texts:
                if not should_include(text):
                    continue
                entries = extract_entries_from_textbox(text)
                for entry in entries:
                    base_entry = {
                        "Slide": slide_num,
                        "Textbox Content": entry["raw_text"],
                        "Docket Number": entry["docket_number"],
                        "Application Number": entry["application_number"],
                        "PCT Number": entry["pct_number"],
                        "WIPO Number": entry["wipo_number"],
                        "Due Dates": "; ".join(entry["due_dates"]),
                        "Extension": entry["extension"]
                    }
                    split_entries = date_split(base_entry["Due Dates"], base_entry["Textbox Content"], base_entry)
                    for e in split_entries:
                        if get_earliest_due_date(e["Due Date"]) >= date.today():
                            results.append(e)

    if not results:
        return pd.DataFrame(columns=["Slide", "Textbox Content", "Docket Number", "Application Number", "PCT Number", "WIPO Number", "Due Date", "Extension", "Action"])

    df = pd.DataFrame(results)
    df["Earliest Due Date"] = df["Due Date"].apply(get_earliest_due_date)
    df = df.sort_values(by="Earliest Due Date", ascending=True)
    df = df.drop(columns=["Earliest Due Date"])
    return df

# === Streamlit UI ===
st.title("\U0001F4CA DocketPoint")
st.sidebar.image("firm_logo.png", use_container_width=True)
st.sidebar.markdown("---")
st.sidebar.markdown("""
**About DocketPoint**

This tool extracts docket numbers, application numbers, and due dates from PowerPoint files.  
It helps organize patent prosecution data and export it to Excel for streamlined docket tracking.
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
