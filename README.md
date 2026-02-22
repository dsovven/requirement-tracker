# Requirements Tracker

A desktop application for capturing and managing requirements from PDF documents. Draw selections around content in a PDF, automatically extract text, stamp numbered labels, and export structured tracking documents in Word and Excel formats.

## Features

- **PDF Viewer** -- Open and navigate PDF documents with zoom and pan controls
- **Requirement Capture** -- Draw rectangles around content to capture requirements with auto-numbered labels
- **Text Extraction** -- Extracts text natively from PDFs, with OCR fallback for scanned documents
- **Screenshot Editing** -- Annotate captured screenshots with highlight and white-out tools
- **Hierarchical Numbering** -- Main requirements (1, 2, 3) and sub-requirements (1.1, 1.2, 1.3)
- **Export** -- Generate Word (.docx) and Excel (.xlsx) tracking documents with embedded screenshots
- **Save Markup** -- Save the annotated PDF with stamped requirement numbers and outlines

## Requirements

- Python 3.11+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (optional, for OCR on scanned PDFs)

## Installation

```bash
pip install -r requirements.txt
```

To enable OCR for scanned documents, install Tesseract:

- **Linux:** `sudo apt-get install tesseract-ocr`
- **macOS:** `brew install tesseract`
- **Windows:** Download from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)

## Usage

```bash
python requirements_tracker.py
```

On Windows, you can also run:

```
run.bat
```

### Workflow

1. **Open a PDF** (Ctrl+O)
2. **Draw rectangles** around requirement text to capture them
3. **Edit screenshots** if needed (highlight or white-out regions)
4. **Save the markup** (Ctrl+S) to produce a stamped PDF
5. **Export** (Ctrl+E) to generate Word and Excel tracking documents

### Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+O | Open PDF |
| Ctrl+S | Save markup |
| Ctrl+E | Export documents |
| Ctrl+Plus/Minus | Zoom in/out |
| PageUp/PageDown | Navigate pages |
| F | Fit width |
| Delete | Delete selected requirement |

## Dependencies

| Package | Purpose |
|---|---|
| PyMuPDF (fitz) | PDF rendering and manipulation |
| PyQt5 | GUI framework |
| python-docx | Word document generation |
| openpyxl | Excel workbook generation |
| pytesseract | OCR text extraction |
| Pillow | Image processing |
