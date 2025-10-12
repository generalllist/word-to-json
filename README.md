# Word to JSON/TSV Converter

This tool extracts data from tables in .docx documents and saves it in a structured format:
→ **JSON** — for programmatic processing
→ **TSV** — for safe viewing in spreadsheet programs (LibreOffice Calc, Excel)

Especially useful for marking up texts (e.g., literary works) in tabular form.

---

## Features

- Support for any number of .docx files in a project folder
- Preservation of the exact table structure:

- 5 columns in a fixed order:

`event`, `status`, `psychophysiological_state`, `circumstances`, `pattern_of_behavior`
- Each **physical table row = one record**
- Empty cells → empty rows (not skipped!)
- Safe saving to **TSV** (not CSV!) to avoid distortions due to:
- line breaks in cells
- quotation marks, commas, dashes, and other special characters.

---

## Requirements

- **Operating System**: Linux (tested on **Linux Mint 22.1**)
- **Python 3.10+**
- **LibreOffice** (for viewing TSVs)

> On Ubuntu/Debian systems, you may need to install `python3-venv`.

---

## Installation and Run

### 1. Install system dependencies
```bash
sudo apt update
sudo apt install python3.12-venv
```

### 2. Clone the repository
```bash
git clone https://github.com/your-name/word-to-json.git
cd word-to-json
```

### 3. Set up a virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 4. Prepare data
Place your .docx files in the project folder:
```bash
cp ~/path/to/your/files/*.docx ./
```

### 5. Run Extract
```bash
python word_to_json.py
```
→ For each `.docx`, a `.json` will be created.

### 6. Convert to TSV (recommended!)
```bash
# For a single file:
python json_to_tsv.py your_file.json

# For all JSON in a folder:
for f in *.json; do python json_to_tsv.py "$f"; done
```

### 7. Open the result
- Double-click the `.tsv` file → it will open in **LibreOffice Calc**
- Or open it manually via **File → Open**, specifying:
- **Delimiter: Tab**
- **Encoding: UTF-8**

---

## Important!

- **Don't use CSV** — text often contains line breaks and quotation marks, which break the CSV structure.
- **TSV (Tab-Separated Values)** is a reliable format for such data.
- The original .docx file must contain **exactly one table** with **5 columns** in the specified order.

---

## Project Structure

```
word-to-json/
├── word_to_json.py # Main script: .docx → JSON
├── json_to_tsv.py # Converter: JSON → TSV
├── requirements.txt # Python Dependencies
├── README.md # These instructions
└── .gitignore # Excludes venv, .docx, and other temporary files
```

---

## License

This project is open and free. Use, modify, share!

---

> Made with ❤️ for researchers, linguists, and anyone who works with text in tabular markup.