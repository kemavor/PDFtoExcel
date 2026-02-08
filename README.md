# PDFtoExcel

An application to convert PDF documents and present them in Excel format.

## Description

This tool extracts tables from PDF documents and converts them into Excel spreadsheets. It's useful for digitizing data from PDF reports, invoices, statements, and other tabular documents.

## Features

- Extract tables from PDF files
- Convert to Excel (.xlsx) format
- Support for multiple tables per PDF
- Extract from all pages or specific pages
- Command-line interface for easy usage
- Python API for programmatic use

## Installation

1. Clone this repository:
```bash
git clone https://github.com/kemavor/PDFtoExcel.git
cd PDFtoExcel
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

Note: This tool requires Java to be installed on your system (required by tabula-py).

## Usage

### Command Line Interface

Basic usage:
```bash
python pdf_to_excel.py input.pdf
```

Specify output file:
```bash
python pdf_to_excel.py input.pdf -o output.xlsx
```

Extract from specific pages:
```bash
python pdf_to_excel.py input.pdf -p 1-3
```

### Python API

```python
from pdf_to_excel import convert_pdf_to_excel

# Basic conversion
convert_pdf_to_excel('input.pdf')

# Custom output filename
convert_pdf_to_excel('input.pdf', 'output.xlsx')

# Extract from specific pages
convert_pdf_to_excel('input.pdf', pages='1-3')
```

### Examples

Run the example script to see usage patterns:
```bash
python example.py
```

## Requirements

- Python 3.7+
- tabula-py
- pandas
- openpyxl
- Java Runtime Environment (JRE)

## How It Works

1. The tool uses tabula-py to extract tables from PDF files
2. Extracted data is processed with pandas DataFrames
3. Data is written to Excel format using openpyxl
4. Each table is placed in a separate sheet (if multiple tables exist)

## Limitations

- Works best with PDFs containing structured tables
- Requires Java to be installed
- May not work well with scanned PDFs (use OCR first)
- Complex table layouts may require manual adjustment

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
