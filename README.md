# MC Marking - Multiple Choice Answer Sheet Marker

A desktop application for automatically marking multiple choice answer sheets using optical mark recognition (OMR) and OCR technology.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![PyQt5](https://img.shields.io/badge/GUI-PyQt5-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## Features

- **PDF Support**: Load and process multi-page PDF answer sheets
- **Visual Mark Definition**: Draw rectangles to define answer regions and text fields
- **Bubble Detection**: Automatically detect filled bubbles (A, B, C, D options)
- **Color Mark Support**: Detects both pencil (gray) and colored (blue) marks
- **OCR Text Recognition**: Extract text from defined regions using EasyOCR
- **Answer Key Comparison**: Set correct answers and automatically calculate scores
- **Batch Processing**: Process all pages in a PDF with one click
- **Excel Export**: Export results to Excel spreadsheet
- **Template Saving**: Save and load mark templates for reuse

## Installation

1. Clone the repository:
```bash
git clone https://github.com/kenkmc/MC_marking.git
cd MC_marking
```

2. Create a virtual environment:
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# or
source .venv/bin/activate  # Linux/Mac
```

3. Install dependencies:
```bash
pip install PyQt5 PyMuPDF easyocr openpyxl Pillow numpy opencv-python torch torchvision
```

## Usage

1. Run the application:
```bash
python main.py
```

2. **Load PDF**: Click "Load PDF" to open an answer sheet PDF

3. **Define Marks**:
   - Select mark type: "Text" for name/ID fields, "Option" for A/B/C/D bubbles
   - Draw rectangles on the answer sheet to define recognition areas
   - Each mark is labeled with a question number (Q1, Q2, etc.)

4. **Set Answer Key**: Right-click on option marks to set the correct answer

5. **Recognize**:
   - Click "Recognize Current Page" for single page
   - Click "Recognize All Pages" to batch process entire PDF

6. **Export**: Click "Export to Excel" to save results

## Mark Types

| Type | Description | Recognition Method |
|------|-------------|-------------------|
| Text | Student name, ID, class | OCR (EasyOCR) |
| Option | Multiple choice A/B/C/D | Bubble detection (darkness + saturation analysis) |

## Detection Algorithm

The bubble detection uses a combined scoring system:
- **Darkness Score**: Measures grayscale intensity
- **Saturation Score**: Detects colored marks (blue pens)
- **Blue Channel Score**: Additional weight for blue ink

Formula: `combined_score = darkness * 1.0 + saturation * 0.5 + blue * 0.3`

## Requirements

- Python 3.10+
- PyQt5
- PyMuPDF (fitz)
- EasyOCR
- OpenPyXL
- Pillow
- NumPy
- OpenCV
- PyTorch (CPU version)

## License

MIT License
