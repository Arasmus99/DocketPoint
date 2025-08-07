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

# === Extract from PowerPoint Shapes ===
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

# === Core Parsing Logic ===
def extract_action_and_dates_from_raw_text(entry):
    line = entry.get("line", "")
    line_lower = line.lower()
    due_date = None
    extension_date = None
    action = ""

    try:
        date_match = PATTERNS["date"].search(line)
        if date_match:
            action = line.split(date_match.group())[0].strip()

        # If 'ext' or 'extension' exists, allow due + extension
        if "ext" in line_lower or "extension" in line_lower:
            due_pos = line_lower.find("due")
            after_due = line[due_pos:]
            due_match = PATTERNS["date"].search(after_due)
            if due_match:
                due_date = parse(due_match.group(), dayfirst=False, fuzzy=True).strftime("%m/%d/%Y")

            ext_pos = line_lower.find("ext")
            after_ext = line[ext_pos:]
            ext_match = PATTERNS["date"].search(after_ext)
            if ext_match:
                extension_date = parse(ext_match.group(), dayfirst=False, fuzzy=True).strftime("%m/%d/%Y")

        # Otherwise, only assign due_date if trusted keywords are present
        elif any(kw in line_lower for kw in ["last day to pay", "final date", "deadline"]):
            trusted_match = PATTERNS["date"].search(line)
            if trusted_match:
                due_date = parse(trusted_match.group(), dayfirst=False, fuzzy=True).strftime("%m/%d/%Y")

    except:
        pass

    return action, due_date, extension_date

def extract_entries_from_textbox(text):
    entries = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    meta = {
        "docket_number": None,
        "application_number": None,
        "pct_number": None,
        "wipo_number": None,
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

        # Only treat lines with "due" as deadline candidates
        if "due" in line.lower():
            entries.append({
                "slide": None,
                "raw_text": meta["raw_text"],
                "line": line,
                "docket_number": meta["docket_number"],
                "application_number": meta["application_number"],
                "pct_number": meta["pct_number"],
                "wipo_number": meta["wipo_number"],
            })

    return entries

def extract_from_pptx(upload, months_back):
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
                    entry["slide"] = slide_num
                    action, corrected_due, extension = extract_action_and_dates_from_raw_text(entry)
                    entry["Action"] = action
                    entry["Extension"] = extension
                    entry["due_date"] = corrected_due

                    # Final cutoff filtering
                    if not entry["due_date"]:
                        continue
                    try:
                        if parse(entry["due_date"]).date() < date.today() - timedelta(days=30 * months_back):
                            continue
                    except:
                        continue

                    results.append({
                        "Slide": entry["slide"],
                        "Textbox Content": entry["raw_text"],
                        "Docket Number": entry["docket_number"],
                        "Application Number": entry["application_number"],
                        "PCT Number": entry["pct_number"],
                        "WIPO Number": entry["wipo_number"],
                        "Due Date": entry["due_date"],
                        "Action": entry["Action"],
                        "Extension": entry["Extension"]
                    })

    if not results:
        return pd.DataFrame(columns=[
            "Slide", "Textbox Content", "Docket Number", "Application Number",
            "PCT Number", "WIPO Number", "Due Date", "Action", "Extension"
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
