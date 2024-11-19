Bilingual Glossary Converter

This project extracts bilingual glossary content from a PDF file, converts it to a formatted Word document, and exports the structured content to an Excel file with labeled columns for "Finnish Term," "Swedish Term," and "Description."

Project Overview

1. Extract Text from PDF: 
   - Read each page of the PDF using the `PyMuPDF` library to capture text, preserving the glossary’s original formatting.
   
2. Convert to Word Document:
   - Using `python-docx`, write the extracted text to a Word document, with optional formatting for headings or specific terms as needed.

3. Export to Excel:
   - Parse the Word content and use `pandas` and `openpyxl` to populate an Excel file with labeled columns for "Finnish Term," "Swedish Term," and "Description."

Prerequisites

- Python 3.x
- Libraries: `pymupdf`, `python-docx`, `pandas`, `openpyxl`

Install dependencies:
```bash
pip install pymupdf python-docx pandas openpyxl
```

Step-by-Step Instructions

1.  **PDF to Word**: Run the `pdf_to_word.py` script to extract text from the PDF and save it in a Word document.
   
2. **Word to Excel**: Run the `word_to_excel.py` script to read the Word file and export glossary entries into an Excel sheet with labeled columns.

## File Structure

- `pdf_to_word.py`: Script to convert PDF content to Word.
- `word_to_excel.py`: Script to parse Word content and export to Excel.
- `output_glossary.docx`: The intermediate Word document generated from the PDF.
- `glossary_output.xlsx`: The final Excel file with structured terms.
