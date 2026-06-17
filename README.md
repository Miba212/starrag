# PDF to Excel Converter

Converts `Checkliste Große Wartung.pdf` into `Checkliste Große Wartung.xlsx`, preserving the visual layout of each page as closely as possible.

## Requirements

- Python 3.8+

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python pdf_to_excel.py
```

The script reads `Checkliste Große Wartung.pdf` from the current directory and produces `Checkliste Große Wartung.xlsx` in the same directory.

## Output

- One Excel sheet per PDF page, named `Page 1`, `Page 2`, etc.
- Text blocks are placed in merged cells at their approximate page positions.
- Images are inserted at their corresponding positions.
- Column widths and row heights approximate an A4 page layout.
