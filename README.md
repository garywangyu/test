# Document Search Tool

This tool provides a basic GUI for searching text across multiple document formats.
Supported formats include PDF, Word (.docx), PowerPoint (.pptx), Excel (.xlsx) and
plain text files. Results show snippets with highlighted keywords.

## Requirements

```
pip install PyMuPDF python-docx python-pptx openpyxl pillow
```

## Usage

```
python document_search_tool.py
```

Select a root directory when prompted, enter a keyword and click `Search` to
scan files under that directory.
