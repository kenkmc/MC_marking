# CheckMate – OMR Marking & Analysis

CheckMate is a Windows desktop app for Optical Mark Recognition (OMR). It detects filled answers from scanned PDF answer sheets, supports optional OCR for text fields (name/class/student ID), and exports results and analysis to Excel.

## What the program does

- Import scanned PDFs of answer sheets
- Define regions for multiple-choice bubbles and text fields
- Recognize filled options and (optionally) OCR text fields
- Review and edit results in-app
- Export to Excel with scores, per-question statistics, and topic analysis
- Batch process multiple PDFs using saved templates

## How it works (high level)

1. **Import PDF** and navigate pages.
2. **Mark regions**:
	- Options (bubble groups)
	- Text fields (name/class/student ID, etc.)
	- Alignment reference (optional but improves consistency)
3. **Recognize all pages**: auto-deskew and auto-align (optional), then detect filled bubbles or OCR text fields.
4. **Review results** in the right-side table. You can edit detected values directly.
5. **Export to Excel**: results, scores, summary stats, and topic analysis.

## Features

- Auto-deskew and auto-align for scanned pages
- Flexible template marking (save/load templates)
- Manual student info entry or paste from Excel
- Per-question % correct and overall summary statistics
- Topic mapping and topic-level analysis
- Export images with answer overlays (optional)

## Requirements

- **OS**: Windows 10/11
- **Python**: 3.10+ (for running from source)
- **Dependencies**: see [requirements.txt](requirements.txt)
- **OCR** (optional for text fields):
  - EasyOCR (recommended) — downloads models on first run (internet required)
  - Tesseract — install separately if you prefer this engine

## Accuracy and limitations

Accuracy depends on scan quality and marking consistency. Typical issues include faint marks, heavy erasures, skewed scans, and inconsistent lighting.

To improve accuracy:

- Use clear, dark fill marks (not ticks)
- Scan at consistent DPI with good contrast
- Enable auto-deskew and auto-align
- Mark an alignment reference region on the template
- Review and correct outliers in the results table

OCR accuracy for text fields varies by handwriting/print quality, font, and language. Manual editing is supported for corrections.

## Build a Windows EXE

The steps below build a standalone Windows executable (.exe) that runs without preinstalled Python.

### One-click build (recommended)

- [build_exe.bat](build_exe.bat)

Or use PowerShell:

- [build_exe.ps1](build_exe.ps1)

Output:

- [dist/CheckMate.exe](dist/CheckMate.exe)

### Important notes

- **EasyOCR** downloads models on first run (internet required).
- If you use **Tesseract**, install Tesseract OCR separately; otherwise use EasyOCR.
- To reduce size, remove unused OCR modules and rebuild.
- Custom icon: place [app.ico](app.ico) at the project root and rebuild.

### Advanced: manual build

1. Install
	- `pip install pyinstaller`
2. Build
	- `pyinstaller mc_marking.spec`

Edit [mc_marking.spec](mc_marking.spec) to customize the build.
