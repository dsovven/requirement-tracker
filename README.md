# Requirements Tracker

A desktop application for capturing and documenting requirements from PDF documents. Select regions on a PDF, automatically number them, annotate with highlights or white-outs, extract text, and export to Word, Excel, or marked-up PDF.

## Features

- **PDF Viewing** — Open, navigate, and zoom PDF files
- **Requirement Capture** — Draw rectangles on PDF pages to capture requirement areas with automatic sequential numbering (supports hierarchical numbering like 1, 1.1, 1.2, 2, etc.)
- **Screenshot Annotation** — Highlight or white-out regions using brush or rectangle tools with full undo support
- **Text Extraction** — Extract text from captured regions via native PDF extraction or OCR (Tesseract) fallback
- **Export** — Generate marked-up PDFs, Word documents (`.docx`), or Excel spreadsheets (`.xlsx`)

## Requirements

- Python 3.7+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (optional, for OCR text extraction)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

**Windows:**

```bash
run.bat
```

**Manual:**

```bash
python requirements_tracker.py
```

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+O | Open PDF |
| Ctrl+S | Save markup |
| Ctrl+E | Export document |
| Ctrl+Z | Undo last delete |
| Ctrl+G | Go to page |
| Ctrl+/- | Zoom in/out |
| Page Up/Down or Arrow Keys | Navigate pages |
| Delete | Delete selected requirement(s) |
| F | Fit width |

## Dependencies

| Package | Purpose |
|---|---|
| PyQt5 | Desktop GUI framework |
| PyMuPDF (fitz) | PDF rendering and manipulation |
| python-docx | Word document export |
| openpyxl | Excel export |
| Pillow | Image processing |
| pytesseract | OCR text extraction (optional) |
