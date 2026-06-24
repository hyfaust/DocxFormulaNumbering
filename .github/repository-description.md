# DocxFormulaNumbering

Automatic equation numbering tool for Word documents.

## Description

DocxFormulaNumbering (DFN) is a Python tool that automatically adds equation numbers and bookmarks to single-line formulas in Microsoft Word documents.

## Features

- Automatic detection of single-line formulas
- Chapter-based numbering: (1-1), (1-2), (2-1)...
- Sequential numbering: (1), (2), (3)...
- Bookmark creation for cross-referencing
- eqArr format support
- 100% Word-compatible

## Quick Start

```bash
pip install -r requirements.txt
python dfn.py input.docx output.docx
```

## License

GPL-3.0
