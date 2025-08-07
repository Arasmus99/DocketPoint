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

# === Helpers ===

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

def extract_dates_in_future(text):
    future_dates = []
    for match in PATTERNS["date"].findall(text):
        try:
            parsed = parse(match, dayfirst=False, fuzzy=True)
            if parsed.date() >= date.today():
                future_dates.append((match, parsed.strftime("%m/%d/%Y")))
        except:
            continue
    return future_dates

# === Extract Action + Dates From Line ===
def extract_action_and_dates_from_raw_text(entry):
    raw_lines = entry["raw_text"].splitlines()
    due_date = None
    extension_date = None
    action = ""

    for line in raw_lines:
        line_lower = line.lower()

        if "due" in line_lower:
            try:
                # === Extract due date ===
                due_pos = line_lower.find("due")
                after_due = line[due_pos:]
                due_match = PATTERNS["date"].search(after_due)
                if due_match:
                    due_date_raw = due_match.group()
                    parsed_due = parse(due_date_raw, dayfirst=False, fuzzy=True)

                    if parsed_due.date() >= date.today():
                        due_date = parsed_due.strftime("%m/%d/%Y")

                        # === Action is just from the line with the due date ===
                        action = line.split(due_date_raw)[0].strip()

                        # === Extract extension date if present ===
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

# === Main Extractor ===
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

    # Extract meta info once
    for line in lines:
        clean_line = re.sub(r"[^0-9A-Za-z/,.\s-]", "", line).replace(",", "")
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

    # Count number of future dates in total
    total_future_dates = extract_dates_in_future(text)

    if len(total_future_dates) > 1:
        # Do line-by-line parsing
        for line in lines:
            parsed = extract_action_and_dates_from_raw_text(line)
            if parsed:
                entries.append({
                    "slide": None,
                    "raw_text": meta["raw_text"],
                    "docket_number": meta["docket_number"],
                    "application_number": meta["application_number"],
                    "pct_number": meta["pct_number"],
                    "wipo_number": meta["wipo_number"],
                    "Action": parsed["Action"],
                    "Due Date": parsed["Due Date"],
                    "Extension": parsed["Extension"]
                })
    else:
        parsed = extract_action_and_dates_from_raw_text(text)
        if parsed:
            entries.append({
                "slide": None,
                "raw_text": meta["raw_text"],
                "docket_number": meta["docket_number"],
                "application_number": meta["application_number"],
                "pct_number": meta["pct_number"],
                "wipo_number": meta["wipo_number"],
                "Action": parsed["Action"],
                "Due Date": parsed["Due Date"],
                "Extension": parsed["Extension"]
            })

    return entries

# === Full PPTX Extractor ===
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

                    # Final cutoff filter
                    try:
                        if parse(entry["Due Date"]).date() < date.today() - timedelta(days=30 * months_back):
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
                        "Action": entry["Action"],
                        "Due Date": entry["Due Date"],
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
        st.download_button("\U0001F4E5 Download Excel", output, file_name="extractedDocket.xlsx")
