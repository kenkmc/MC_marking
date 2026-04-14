# Copilot instructions for CheckMate (MC_marking)

## Big picture architecture
- Single-process PyQt5 desktop app. UI + processing live in [omr_software.py](omr_software.py); entry point is [main.py](main.py) (note the EasyOCR import order to avoid Windows DLL conflicts).
- Core concepts:
  - `MarkItem` + `MarkingView` manage interactive regions (text/option/alignment) on a `QGraphicsScene`.
  - `OMRSoftware` orchestrates PDF rendering (PyMuPDF), corrections (OpenCV deskew/align), recognition, and exports.
  - Results are stored per page in `self.results[page_idx] = {"options":{}, "text":{}, "option_crops":{}, "text_crops":{}}` and drive the right-side results table and exports.
  - Student absence is stored in `self.student_absence = {page_idx: bool}`.
- Data flow: PDF page → image (PyMuPDF → PIL/NumPy) → optional deskew/align → crop by scene marks (using `page_offsets`) → detect bubbles/OCR → update `self.results` → export Excel/images.

## i18n / Translation system
- All UI strings use `tr(key, **kwargs)` from the module-level translation dict `_TRANSLATIONS`.
- Two languages: `"en"` (English, default) and `"zh"` (繁體中文).
- Language is switched at runtime via `set_language(lang)` which triggers `_switch_language()` to rebuild the UI while preserving application state.
- Add new translatable strings by adding entries to both `en` and `zh` dicts in `_TRANSLATIONS`.
- Version constant: `APP_VERSION` at module level.

## Key workflows / commands
- Run from source: `python omr_software.py` (or `python main.py`) after installing [requirements.txt](requirements.txt).
- Build Windows EXE (PyInstaller onedir): [build_exe.bat](build_exe.bat) or [build_exe.ps1](build_exe.ps1) which use [mc_marking.spec](mc_marking.spec).

## Project-specific patterns & conventions
- Import order matters: `easyocr` is imported before PyQt5 in [main.py](main.py) to avoid DLL conflicts on Windows.
- Marks are stored in **scene coordinates**; converting to image coordinates always subtracts the page’s `page_offsets` (see `load_page()` and recognition loops in [omr_software.py](omr_software.py)).
- Alignment/deskew are optional and controlled by UI checkboxes (`check_auto_deskew`, `check_auto_align`). Alignment uses a user-marked region (`MARK_TYPE_ALIGN`) when present, otherwise falls back to table-boundary detection. When template matching fails due to large shifts, a phase-correlation fallback (`_align_phase_correlation_fallback`) attempts full-page alignment.
- Bubble detection is implemented in `detect_filled_option()` with combined grayscale/saturation heuristics and writes debug records to `self.debug_records` for export; debug crops go to `debug_crops/`.
- Templates are JSON-serializable mark definitions (`get_all_marks_data()` / `load_marks_from_data()`), and batch processing loads templates per PDF.
- Student info dialog and topic dialog both support Ctrl+V paste from spreadsheets (tab-separated).
- Recognition can be run on all pages (`run_recognition_all`), or selectively on specific pages (`run_recognition_selected`).

## External dependencies / integrations
- PDF rendering: PyMuPDF (`fitz`), image ops: PIL + OpenCV (`cv2`), GUI: PyQt5, Excel export: openpyxl, OCR: EasyOCR or Tesseract.
- Tesseract is optional and must be installed externally; EasyOCR downloads models on first run.

## Where to look first
- UI + behavior: [omr_software.py](omr_software.py)
- App entrypoint / import order: [main.py](main.py)
- Build pipeline: [build_exe.ps1](build_exe.ps1), [build_exe.bat](build_exe.bat), [mc_marking.spec](mc_marking.spec)
