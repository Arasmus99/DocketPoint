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

# === Recursive text extraction ===
def extract_texts_from_shape_recursive(shape):
    texts = []
    if shape.shape_type == 6:  # GroupShape
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

def extract_action_from_raw_text(entry):
    """Extracts the action (text before date) from raw_text for a given due_date."""
    raw_lines = entry["raw_text"].splitlines()
    for line in raw_lines:
        if entry["due_date"] in line:
            return line.split(entry["due_date"])[0].strip()
        # Try fuzzy matching if reformatted date isn't found
        for match in PATTERNS["date"].findall(line):
            try:
                parsed = parse(match, dayfirst=False, fuzzy=True).strftime("%m/%d/%Y")
                if parsed == entry["due_date"]:
                    return line.split(match)[0].strip()
            except:
                continue
    return ""

def extract_entries_from_textbox(text, months_back=0):
    entries = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    cutoff_date = date.today() - timedelta(days=30 * months_back)

    meta = {
        "docket_number": None,
        "application_number": None,
        "pct_number": None,
        "wipo_number": None,
        "due_dates": [],
        "raw_text": "\n".join(lines)
    }

    for line in lines:
        clean_line = line.replace(" /,", "/").replace("/", "/").replace(",,", ",").replace(" /", "/")
        clean_line = re.sub(r"[^0-9A-Za-z/,.\s-]", "", clean_line)
        clean_line = clean_line.replace(",", "")

        if not meta["docket_number"] and PATTERNS["docket_number"].search(clean_line):
            meta["docket_number"] = PATTERNS["docket_number"].search(clean_line).group(0)

        if not meta["pct_number"] and PATTERNS["pct_number"].search(clean_line):
            meta["pct_number"] = PATTERNS["pct_number"].search(clean_line).group(0)

        if not meta["application_number"] and PATTERNS["application_number"].search(clean_line):
            meta["application_number"] = PATTERNS["application_number"].search(clean_line).group(0)
        elif not meta["application_number"] and PATTERNS["alt_application_number"].search(clean_line):
            meta["application_number"] = PATTERNS["alt_application_number"].search(clean_line).group(0)

        if not meta["wipo_number"] and PATTERNS["wipo_number"].search(clean_line):
            meta["wipo_number"] = PATTERNS["wipo_number"].search(clean_line).group(0)

        for match in PATTERNS["date"].findall(clean_line):
            try:
                parsed = parse(match, dayfirst=False, fuzzy=True)
                if parsed.date() >= cutoff_date:
                    meta["due_dates"].append(parsed.strftime("%m/%d/%Y"))
            except:
                continue

    if not (meta["docket_number"] or meta["application_number"] or meta["pct_number"] or meta["wipo_number"]):
        return []

    for due_date in meta["due_dates"]:
        entries.append({
            "slide": None,
            "raw_text": meta["raw_text"],
            "docket_number": meta["docket_number"],
            "application_number": meta["application_number"],
            "pct_number": meta["pct_number"],
            "wipo_number": meta["wipo_number"],
            "due_date": due_date
        })

    return entries

def extract_from_pptx(upload, months_back):
    prs = Presentation(upload)
    results = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        for shape_num, shape in enumerate(slide.shapes, start=1):
            texts = extract_texts_from_shape_recursive(shape)
            for text in texts:
                if not should_include(text):
                    continue
                entries = extract_entries_from_textbox(text, months_back)
                for entry in entries:
                    entry["slide"] = slide_num
                    entry["Action"] = extract_action_from_raw_text(entry)
                    results.append({
                        "Slide": entry["slide"],
                        "Textbox Content": entry["raw_text"],
                        "Docket Number": entry["docket_number"],
                        "Application Number": entry["application_number"],
                        "PCT Number": entry["pct_number"],
                        "WIPO Number": entry["wipo_number"],
                        "Due Date": entry["due_date"],
                        "Action": entry["Action"]
                    })

    if not results:
        return pd.DataFrame(columns=[
            "Slide", "Textbox Content", "Docket Number", "Application Number",
            "PCT Number", "WIPO Number", "Due Date", "Action"
        ])

    df = pd.DataFrame(results)
    df["Earliest Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce")
    df = df.sort_values(by="Earliest Due Date", ascending=True).drop(columns=["Earliest Due Date"])
    return df

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
        df = extract_from_pptx(ppt_file, months_back)
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
        st.download_button("\U0001F4E5 Download Excel", output, file_name="combined_extracted_data.xlsx")
