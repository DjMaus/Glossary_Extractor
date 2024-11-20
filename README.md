# PDF Text Extractor and Converter

This project provides a user-friendly interface for converting PDF documents to Word format, with the ability to later process the content into Excel. The tool features page range selection and real-time conversion progress tracking.

## Features

- GUI-based PDF to Word converter
- Page range selection for partial document conversion
- Real-time conversion progress tracking
- Support for converting entire documents
- Excel export functionality (coming soon)

## Prerequisites

- Python 3.x
- Required libraries:
  ```bash
  pip install pdf2docx pdfplumber python-docx pandas openpyxl tkinter
  ```

## Usage

1. **PDF to Word Conversion**:
   - Run `improved_pdf_to_word.py`
   - Select your PDF file using the browse button
   - Choose between:
     - Converting specific pages (enter start and end page numbers)
     - Converting the entire document (check "Convert Whole Document")
   - Select output location for the Word file
   - Click "Convert" and monitor the progress

2. **Word to Excel** (Feature in development):
   - Will allow structured export of Word content to Excel format
   - Support for custom column mapping
   - Data organization and categorization

## Project Structure

- `improved_pdf_to_word.py`: GUI application for PDF to Word conversion
- `word_to_excel.py`: Script for converting Word content to Excel (upcoming)

## Future Enhancements

- Custom formatting options for Word output
- Batch processing capabilities
- Advanced Excel export features
- OCR support for scanned documents
