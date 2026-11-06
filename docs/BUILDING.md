# Building PDF2MD

## Windows environment

Use Python 3.12.10 x64 for the published Windows build, Microsoft .NET Framework's C# compiler and Inno Setup 6.7+. `requirements-release.txt` records the exact lightweight package versions; `requirements.txt` contains the supported ranges for development. The published bundles contain Python, so end users do not need to install it.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-release.txt
.\.venv\Scripts\python.exe pdf2md_ui.py
```

## Tests

```powershell
$env:PDF2MD_SETTINGS_FILE = Join-Path $env:TEMP 'pdf2md-test-settings.json'
.\.venv\Scripts\python.exe -B -m unittest discover -s tests
```

GUI tests need an interactive desktop. Tests for legacy `.doc` and `.xls` need Microsoft Word and Excel; otherwise those checks may be skipped. Generated fixtures are temporary and do not require personal documents.

## Installer

```powershell
.\.venv\Scripts\python.exe -B tools\build_windows_installer.py --iscc 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe' --version 0.2.0
```

The output is `dist/PDF2MD-0.2.0-setup.exe`. Use `--output-dir` to choose another directory. The default installation is for all users; `/CURRENTUSER` selects per-user installation. Both modes create Desktop and Start menu shortcuts. The published installation tests used `/CURRENTUSER`.

## Light portable ZIP

```powershell
.\.venv\Scripts\python.exe -B tools\build_portable_light.py --version 0.2.0
```

The output is `dist/PDF2MD-0.2.0-portable.zip`. Extract it completely and run `PDF2MD/PDF2MD.exe`. The ZIP does not install shortcuts or register an uninstaller. It uses the same managed runtime as the installer, storing settings in the Windows user profile and optional Marker files in a separately selected folder. The builder refuses to overwrite an existing ZIP.

Both builders include the application modules, Python runtime, basic dependencies, user documentation and license notices. They exclude Marker, model weights, development caches and package installation URLs that could contain machine-specific paths.

## Optional Marker

Use the application's installer to obtain the pinned Marker dependencies. Its CPU/CUDA choice and dependency versions are defined in `marker_install.py`. Package and model licenses are separate; see `THIRD_PARTY.md`. Marker runs as a subprocess and is not required for the synthetic conversion tests.

## Command line

```powershell
.\.venv\Scripts\python.exe pdf2md.py document.pdf --engine auto --language en
.\.venv\Scripts\python.exe pdf2md.py workbook.xlsx --sheet 2 -o output\
```

The CLI writes a separate Markdown file. It overwrites existing output without prompting but rejects the source itself as the output, including aliases and hard links.

## Release contents

Publish the tested EXE, portable ZIP, matching source archive, `SHA256SUMS.txt` and `release-manifest.json`. Hash the final files after packaging. Record the public source commit in the manifest; source and binaries must refer to the same application version. Do not include settings, logs, temporary files or conversion inputs/outputs.
