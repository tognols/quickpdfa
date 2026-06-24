# QuickPDFA

A tiny PowerShell front-end for Ghostscript that converts one PDF into PDF/A-2b.

## Features
- Drag-and-drop one `.pdf` file
- One-click conversion with **Convert to PDF/A**
- Output file name is `<original>_PDFA.pdf` in the same folder

## Requirements
- Windows PowerShell 5.1+
- Ghostscript installed (recommended 64-bit)
  - Download: https://ghostscript.com/releases/gsdnld.html

## Run
From this folder, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\QuickPDFA.ps1
```

## Notes
- The app tries to find Ghostscript automatically (`gswin64c.exe`/`gswin32c.exe`).
- If Ghostscript is not detected, install it and relaunch the app.
- Conversion uses Ghostscript PDF/A-2b settings with strict compatibility policy and OutputIntent ICC profile embedding.


###
HAPPILY VIBE-CODED BY MATTEO TOGNOLO