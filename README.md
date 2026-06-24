# DocxFormulaNumbering (DFN)

> Automatic Equation Numbering Tool for Word Documents

[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![Word 2016+](https://img.shields.io/badge/Word-2016+-green.svg)](https://products.office.com/word)
[![License: GPL v3](https://img.shields.io/badge/license-GPLv3-blue.svg)](LICENSE)

[English](README.md) | [简体中文](README_zh.md)

---

## Introduction

DocxFormulaNumbering (DFN) is a Python tool that automatically adds equation numbers and bookmarks to single-line formulas in Microsoft Word documents. It supports chapter-based numbering formats and creates bookmarks for cross-referencing.

## Features

- ✅ Automatic detection of single-line formulas in Word documents
- ✅ Chapter-based numbering format: `(1-1)`, `(1-2)`, `(2-1)`...
- ✅ Simple sequential numbering: `(1)`, `(2)`, `(3)`...
- ✅ Bookmark creation for cross-referencing
- ✅ eqArr format support (professional equation layout)
- ✅ 100% Word-compatible via Word COM automation

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Basic Usage

```bash
# Default mode: chapter-number format
python dfn.py input.docx output.docx

# Simple mode: sequential numbering
python dfn.py input.docx output.docx --simple

# eqArr format (requires t1.docx template)
python dfn.py input.docx output.docx --t1

# Overwrite original file
python dfn.py document.docx
```

## Command Line Options

| Option | Description |
|--------|-------------|
| `input` | Input docx file path (required) |
| `output` | Output docx file path (optional, overwrites input by default) |
| `--simple` | Use simple sequential numbering (no chapter numbers) |
| `--t1` | Use eqArr format (requires t1.docx template) |

## Prerequisites

### For Default Mode

Your document must have multi-level list configured:

1. Chapter headings must use "Heading 1" style
2. Apply numbering via Home → Paragraph → Multilevel List
3. Ensure heading style is linked to list level

### For eqArr Format

When using `--t1`, ensure `t1.docx` template exists in the same directory.

## Output Formats

### Default Format (t2)

```
E = mc²#(1-1)
```

Number appears as plain text on the same line as the formula.

### eqArr Format (t1)

```xml
<m:eqArr>
  <m:t>E = mc²#</m:t>
  <m:d>
    <m:t>1</m:t>
  </m:d>
</m:eqArr>
```

Number appears in a delimiter element, integrated with the formula for professional display.

## Bookmark Naming

- **Default mode**: `eq{chapter}_{number}` e.g., `eq1_1`, `eq1_2`, `eq2_1`
- **Simple mode**: `eq{number}` e.g., `eq1`, `eq2`, `eq3`

Cross-reference via Insert → Cross-reference → Bookmark in Word.

## Project Structure

```
DocxFormulaNumbering/
├── dfn.py              # Main script (entry point)
├── README.md           # English documentation
├── README_zh.md        # Chinese documentation
├── USAGE.md            # Detailed usage guide
├── requirements.txt    # Python dependencies
├── LICENSE             # GPL v3 License
├── .gitignore          # Git ignore rules
├── t1.docx             # eqArr format template
└── example.docx        # Example document for testing
```

## Technical Details

- **Language**: Python 3.6+
- **Dependencies**: pywin32 (Word COM automation)
- **Core**: Word Object Model + OMML (Office Math Markup Language)

## Example

### Input Document

```
Chapter 1 Introduction

E = mc²

F = ma

Chapter 2 Relativity

p = mv
```

### Output (Default Mode)

```
Chapter 1 Introduction

E = mc²#(1-1)  [Bookmark: eq1_1]

F = ma#(1-2)   [Bookmark: eq1_2]

Chapter 2 Relativity

p = mv#(2-1)   [Bookmark: eq2_1]
```

## Troubleshooting

### Q: "Document not configured for multi-level list" error

A: Your chapter headings don't have multi-level list applied. Use `--simple` parameter to skip chapter detection.

### Q: Formula numbers display as garbled text

A: Press `F9` to update all fields, or right-click the number and select "Update Field".

### Q: How to reference equation numbers?

A: Use Insert → Cross-reference → Bookmark in Word, select the corresponding `eq*` bookmark.

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## Author

FaustSherpad

## Version

1.0.0
