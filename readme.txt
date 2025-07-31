# 📊 PowerPoint-to-Excel Parser

**Automated extraction of legal and patent docket data from PowerPoint slides into structured Excel spreadsheets.**

This tool parses PowerPoint (.pptx) presentations containing legal or IP case tracking information and exports a clean, filterable Excel sheet containing structured data such as:

- 📁 Docket Numbers  
- 🧾 Application Numbers  
- 🌐 Country Codes  
- 📅 Due Dates  
- 📌 Document Types or Descriptions

---

## 🚀 Features

- ✅ Reads `.pptx` slide text, including multi-line legal entries
- ✅ Extracts key identifiers using intelligent parsing (regex + structure rules)
- ✅ Identifies country codes, application numbers, and US docket patterns
- ✅ Automatically parses and reformats due dates into Excel
- 📥 Exports to `.xlsx` for downstream tracking or reporting

---

## 🧰 How to Use

### 1. Install Requirements

```bash
pip install -r requirements.txt
