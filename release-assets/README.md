# Release assets (distribution-only)

Binary deliverables are committed here for convenience so the latest build is always
downloadable straight from the repo, mirroring the PDFreader release layout.

| File | What it is | Built |
|---|---|---|
| `FileConverterSetup.msi` | WiX installer: installs to `%LocalAppData%\Programs\File Converter`, Start-Menu + Desktop shortcuts, Add/Remove-Programs entry, in-place upgrades (same UpgradeCode) | 2026-10-25 |
| `FileConverter.exe` | Self-contained single-file build (no .NET install needed) | 2026-10-25 |

## Install notes

- **MSI**: double-click and run. If Windows SmartScreen appears (unsigned binary),
  choose *More info → Run anyway*.
- **Portable exe**: download and run directly; nothing is installed.

## Why this folder is not built from CI

Both artifacts are ~70–77 MB and the installer build needs WiX v6 locally
(`build-installer.bat` in the repo root rebuilds the MSI). Refresh them after a
release build by copying from `installer\` and `publish\`.
