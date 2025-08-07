import os
import re
from datetime import datetime, date
from dateutil.parser import parse
from pptx import Presentation
import pandas as pd

PATTERNS = {
    "docket": re.compile(r"\b\d{4,}\.\d{4}(?:-\w+)?\b"),
    "application": re.compile(r"\d{2}/\d{3},\d{3}"),
    "pct": re.compile(r"PCT/US\d{2}/\d{6}"),
    "wipo": re.compile(r"[A-Z]{2}\d{4}/\d{6}"),
    "date": re.compile(r"\b(?:\d{1,2}[/-])?\d{1,2}[/-]\d{2,4}\b")
}

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

def extract_from_pptx(pptx_folder):
    all_entries = []

    for filename in os.listdir(pptx_folder):
        if filename.endswith(".pptx"):
            file_path = os.path.join(pptx_folder, filename)
            prs = Presentation(file_path)

            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        raw_text = shape.text.strip()
                        entries = extract_entries_from_textbox(raw_text, i + 1, filename)
                        all_entries.extend(entries)

    return pd.DataFrame(all_entries)
