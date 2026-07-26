# CheckMate 1.7.1

This patch fixes the reported case where a completely blank 300 dpi answer sheet
produced false answers.

## Fixes

- Rejects paper-bright yellow, green, and purple scanner fringes around hollow answer
  rectangles instead of treating them as coloured pen strokes.
- Locates each printed response rectangle and measures its writable interior, fixing
  legacy templates whose answer boxes sit above/right of the crop centre.
- Preserves genuine dark marks, faint pencil marks, multiple answers, and thin blue
  check marks.
- Prevents saving a template with zero option/text regions.
- Rejects an empty JSON template before it can replace the currently loaded valid
  template, and shows region counts after a successful load.
- Refreshes the bundled 30-question/four-alignment template with v2 page metadata.

The reported blank form was checked against all 30 configured regions: 0 false
answers after the fix. The automated suite includes blank colour-fringe and thin
blue-check regressions.

## Downloads

- **`CheckMate_Setup_v1.7.1.exe`** — recommended per-user Windows installer.
- **`CheckMate_v1.7.1.zip`** — portable edition; extract the complete archive before
  starting `CheckMate.exe`.
- **`SHA256SUMS.txt`** — integrity checks for both packages.

The supplied `ver1.7_30題MC範本_4align_mark.json` contains no regions and is therefore
not usable. Use the corrected bundled template at
`_internal\template\30題MC範本_4align_mark.json`.

The binaries are not commercially code-signed, so Windows SmartScreen may show a
warning. EasyOCR may require internet access to download its model on first use.
