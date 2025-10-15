
```md
# Word to JSON/TSV Converter

This tool extracts data from tables in `.docx` documents and saves it in a structured format:  
→ **JSON** — for programmatic processing  
→ **TSV** — for safe viewing in spreadsheet programs (LibreOffice Calc, Excel)

Especially useful for marking up texts (e.g., literary works) in tabular form.

---

## Features

- Processes any number of `.docx` files in the project folder.  
- Preserves exact table structure: **exactly 5 columns in fixed order**:  
  `event`, `status`, `psychophysiological_state`, `circumstances`, `pattern_of_behavior`.  
- Each **physical table row → one record (JSON object)**.  
- Empty cells are preserved as empty strings (not skipped).  
- Exports to **TSV** (not CSV) to safely handle data containing:
  - line breaks within cells,  
  - quotation marks, commas, dashes, and other special characters.

---

## Requirements

- **OS**: Linux (tested on **Linux Mint 22.1**).  
- **Python 3.10–3.12**.  
- **LibreOffice** (for viewing TSV files).  
- System dependencies sometimes required to build `lxml` from source (not needed if pip uses prebuilt wheels):

  ```bash
  sudo apt install libxml2-dev libxslt1-dev python3-dev
  ```

> On Ubuntu/Debian, you may also need `python3-venv` that matches your installed Python version.

---

## Installation and Usage

### 1. Install system dependencies (if needed)
```bash
sudo apt update
sudo apt install python3-venv libxml2-dev libxslt1-dev python3-dev
```

> **Note**: The `lib*-dev` packages are only required when `lxml` is built from source. On many systems, pip installs prebuilt wheels and these packages are unnecessary.

### 2. Clone the repository
```bash
git clone https://github.com/generalllist/word-to-json.git
cd word-to-json
```

### 3. Set up a virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

### 4. Prepare your data
Place your `.docx` files in the project root (the `word-to-json` folder). For example:

```bash
# Run this command from the project root (folder word-to-json)
cp ~/path/to/your/files/*.docx ./
```

### 5. Run extraction
```bash
python word_to_json.py
```

> The script verifies that at least one `.docx` file exists and that each contains **exactly one table with 5 columns in the required order**. Clear error messages are shown on failure.

### 6. Convert to TSV (recommended)
```bash
# Single file:
python json_to_tsv.py your_file.json

# All JSON files in the folder (bash):
shopt -s nullglob
for f in *.json; do
    python json_to_tsv.py "$f"
done
```

> **Note**: `shopt -s nullglob` prevents the loop from running when no `.json` files exist (bash only). Adjust or omit for other shells.

### 7. Open the result
- Double-click the `.tsv` file → opens in **LibreOffice Calc**.  
- Or open manually via **File → Open**, setting **Delimiter: Tab** and **Encoding: UTF-8**.  
- Any TSV-capable editor (e.g., VS Code, Notepad++) can also be used.

---

## Important!

- **Do not use CSV** — line breaks and quotes will corrupt the format.  
- **TSV (Tab-Separated Values)** is robust for this type of data.  
- Input `.docx` files must contain **exactly one table** with **5 columns** in the specified order.

---

## Project Structure

```
word-to-json/
├── word_to_json.py   # .docx → JSON
├── json_to_tsv.py    # JSON → TSV
├── requirements.txt  # Python dependencies
├── README.md         # This document
└── .gitignore        # Excludes venv, .docx, and temporary files
```

---

## Contents of `requirements.txt`

```text
# XML parsing (required by python-docx)
lxml==6.0.2

# DOCX handling
python-docx==1.2.0

# Backport of typing features for older Python versions
typing_extensions==4.15.0
```

---

## License

Distributed under the **MIT License** — see the `LICENSE` file.

---

> Made with ❤️ for researchers, linguists, and anyone working with textual markup in tabular form.
```
