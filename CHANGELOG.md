# Changelog

## 1.7.1 — 2026-07-26

### Blank-form accuracy

- Fixed false answers caused by yellow, green, or purple scanner fringes around
  hollow rectangular answer boxes.
- Added per-option contour alignment so legacy templates measure the writable
  rectangle interior instead of the geometric centre of an offset crop.
- Added regression coverage for blank colour-fringed boxes and thin blue check marks.
- Verified all 30 answer regions from the reported 300 dpi blank form remain blank.

### Template safety

- Prevented saving templates that contain no option or text regions.
- Rejected empty recognition templates without clearing the currently loaded valid
  template, and showed the loaded option/text/alignment counts in the status bar.
- Added page-size metadata to the bundled 30-question, four-alignment template.

## 1.7.0 — 2026-07-26

### Recognition accuracy

- Rebuilt the OMR decision core with stable option boundaries and combined grayscale,
  contrast, colour, centre-fill, and blue-ink evidence.
- Improved recognition of faint marks, multiple selections, and fully marked rows.
  Blank or ambiguous answers are flagged instead of guessed.
- Changed deskew to preserve the original image dimensions and template coordinates.
- Added robust multi-anchor similarity alignment for rotation, scale, and translation,
  with rejection of an incorrect anchor.
- Ensured selected-page recognition uses page 1 as its alignment reference.

### Review and usability

- Added per-answer confidence, review reasons, colour highlighting, and `F8` navigation
  to the next blank, multiple, invalid, or low-confidence result.
- Added in-app answer-crop preview and Save As.
- Added PDF drag-and-drop and shortcuts for opening PDFs, loading/saving templates,
  and starting recognition.
- Made answer-key edits on the first page update scoring immediately.
- Persisted key recognition and export preferences between sessions.

### Speed and reliability

- Disabled diagnostic image writes by default and moved optional diagnostics to a
  temporary session folder.
- Added a switch to skip text OCR when only OMR is required.
- Unified full-document, selected-page, and batch recognition paths.
- Fixed stale results when opening a new PDF, cross-file batch state, batch OCR calls,
  page offsets, old document cleanup, and pure-rotation alignment.
- Added template schema v2 with reference dimensions and automatic scaling.

### Distribution

- Split minimal runtime, OCR, and build dependencies.
- Added deterministic OMR regression tests, syntax checks, and GUI/EXE smoke tests.
- Added reproducible Windows portable ZIP and per-user installer builds.
- Added SHA-256 checksums and automated GitHub Releases.

## 1.6.4

- Previous public release.
