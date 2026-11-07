# File Converter

A small Windows desktop app (WPF, .NET 10) for personal file conversion — drag files in,
pick a target format, convert. Uses your installed Microsoft Office for high-fidelity
Office↔PDF conversions.

## Supported conversions

| From | To | Engine |
|---|---|---|
| PPTX / PPT | PDF | PowerPoint (must be installed) |
| DOCX / DOC | PDF | Word (must be installed) |
| XLSX / XLS | PDF | Excel (must be installed) |
| PDF | DOCX | Word opens the PDF and saves it as an editable document |
| PNG / JPG / WEBP / BMP / TIFF / GIF | any of the others, or PDF | Magick.NET |
| multiple PDFs | one merged PDF | PDFsharp |
| PDF | page-range PDF (split) | PDFsharp |
| PDF | optimized PDF | PDFsharp (modest savings — heavy compression needs Ghostscript) |

## Run (development)

```
dotnet run
```

## Install on your machine

Double-click **`installer\FileConverterSetup.msi`** (or build a fresh one with
`build-installer.bat`). This installs the real thing:

- App location: `%LocalAppData%\Programs\File Converter\FileConverter.exe`
- Start Menu shortcut + Desktop shortcut ("File Converter")
- Lists in Windows Settings → Apps, with a working uninstaller
- Installing a newer build automatically upgrades in place (same `UpgradeCode`)

The MSI needs the WiX v6 tool to build: `dotnet tool install --global wix --version 6.0.2`
(stay on v6 — WiX v7 requires accepting its Open Source Maintenance Fee EULA).

The app icon lives in `assets\app.ico` (regenerate with `assets\make-icon.ps1`; also wired
into the exe and window title bar via `FileConverter.csproj` / `MainWindow.xaml`).

## Build the single .exe (no installer)

```
publish.bat
```

→ `publish\FileConverter.exe` (self-contained; no .NET install needed on the target machine).

## Command line

```
FileConverter slides.pptx                     # convert with the recommended target
FileConverter photo.png --to webp --out D:\out
FileConverter --merge out.pdf in1.pdf in2.pdf
FileConverter --split in.pdf --pages 1-3,5 [--out out.pdf]
FileConverter --optimize in.pdf [--out out.pdf]
FileConverter --help
```

Office-call timeout (default 120 s) can be overridden with the environment variable
`FILECONVERTER_OFFICE_TIMEOUT_MS` (milliseconds).

## Notes & troubleshooting

- **Office conversions** drive PowerPoint/Word/Excel via COM. If an Office app stops
  responding, the app closes it after the timeout, marks the job failed with a hint, and
  continues with the rest of the queue.
- If Word/Excel/PowerPoint conversions keep timing out: run **Office Quick Repair**
  (Windows Settings → Apps → Microsoft Office → Modify → Quick Repair) and make sure
  Office is signed in/activated.
- PDF → DOCX quality depends on the PDF: text-based PDFs convert well; scanned PDFs
  (images of text) will not become editable text.
- `Optimize PDF` rebuilds the file structure (small savings). For big scanned PDFs,
  install Ghostscript for real compression.
- Test suite (development): `bash testdata/run-tests.sh` after generating test files with
  `testdata/make-test-files.ps1`.

## Machine notes (DSH sandbox)

- Files inside the DSH workspace carry a Low integrity label and package-scoped grants.
  Word/Excel cannot open them (`Workbooks.Open` fails / Word hangs until the watchdog
  fires), and PowerPoint COM may be refused outright. Copy inputs to a folder outside the
  workspace (e.g. `%TEMP%`) and convert there — verified working end-to-end 2026-11-07.
- A watchdog-killed Word instance can poison the next run (exit 1, no output, stray
  WINWORD.EXE). Kill stray WINWORD/EXCEL first, then retry; the same input converts fine.
- PDF → DOCX relies on Word's own PDF-import filter; on 2026-11-07 it intermittently hung
  even in direct Word COM calls (plain .docx opened fine) after a WINWORD crash + pending
  Office update — Office-side transient, not app state. If it wedges repeatedly:
  Windows Settings → Apps → Microsoft Office → Modify → **Quick Repair**, then retry.
