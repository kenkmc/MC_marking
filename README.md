# CheckMate – The definitive OMR software

Windows EXE packaging guide.

The steps below build a Windows executable (.exe) that runs without a preinstalled Python or libraries.

## One-click build (recommended)

Double-click:

- [build_exe.bat](build_exe.bat)

Or use PowerShell:

- [build_exe.ps1](build_exe.ps1)

Output:

- [dist/CheckMate.exe](dist/CheckMate.exe)

## Important notes

- **EasyOCR** downloads models on first run (internet required).
- If you use **Tesseract**, install Tesseract OCR separately; otherwise use EasyOCR.
- To reduce size, remove unused OCR modules and rebuild.
- Custom icon: place [app.ico](app.ico) at the project root and rebuild.

## Advanced: manual build

Install the build tool and run:

1. Install
	- `pip install pyinstaller`
2. Build
	- `pyinstaller mc_marking.spec`

Edit [mc_marking.spec](mc_marking.spec) to customize the build.
