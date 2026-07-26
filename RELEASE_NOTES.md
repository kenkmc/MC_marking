# CheckMate 1.7.0

This release improves OMR accuracy, makes uncertain answers much faster to review,
and replaces the previous ad-hoc package with tested Windows installer and portable
downloads.

## Highlights

- Improved detection of dark or faint pencil, blue ink, blanks, multiple answers,
  and fully marked rows without forced guessing.
- Fixed-canvas deskew plus robust multi-anchor rotation, scale, and translation
  alignment.
- Confidence percentages, clear review reasons, visual highlighting, `F8` review
  navigation, and an in-app answer-crop preview.
- PDF drag-and-drop and shortcuts: `Ctrl+O`, `Ctrl+L`, `Ctrl+Shift+S`, and `Ctrl+R`.
- Optional text OCR and diagnostic images for a faster default recognition path.
- Fixed stale cross-document state, batch OCR failures, answer-key edits, template
  scaling, and selected-page alignment.
- Automated regression tests, executable smoke tests, installer tests, and checksums.

## Downloads

- **`CheckMate_Setup_v1.7.0.exe`** — recommended per-user Windows installer.
- **`CheckMate_v1.7.0.zip`** — portable edition; extract the complete archive before
  starting `CheckMate.exe`.
- **`SHA256SUMS.txt`** — integrity checks for both packages.

Existing JSON templates remain compatible. Re-saving a template in 1.7.0 adds page
size metadata so coordinates can be scaled automatically.

The binaries are not commercially code-signed, so Windows SmartScreen may show a
warning. EasyOCR may require internet access to download its model on first use.
