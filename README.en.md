# PDF2MD

A local Windows application for converting **PDF, Word, Excel, CSV, TSV and TXT to Markdown**. Saves a separate output file and extracts text, tables and images. The interface supports Russian and English.

Built to prepare documents for AI workflows: Markdown can help reduce the tokens spent processing them. Actual savings depend on the document and the AI service used.

[Русский](README.md) · [FAQ](FAQ.md#english) · [Changelog](CHANGELOG.md) · [Privacy](PRIVACY.md#english)

## Download

**Version 0.2.2 · Windows 10 (1809 or later) / Windows 11 · x64**

Download from [Releases](https://github.com/RootforgeIO/PDF2MD/releases).

| File | Purpose |
| --- | --- |
| `PDF2MD-0.2.2-setup.exe` | Windows installer; automatically creates Desktop and Start menu shortcuts |
| `PDF2MD-0.2.2-portable.zip` | No installation: extract the entire folder and run `PDF2MD.exe` |
| `PDF2MD-0.2.2-source.zip` | Source code and build instructions |
| `SHA256SUMS.txt` | Download checksums |

Python and the basic libraries are included. Marker and its models are optional downloads from the Settings menu. Legacy `.doc` and `.xls` require Microsoft Word or Excel. The builds are not digitally signed.

## Quick start

1. Install the application or extract the entire portable archive.
2. Open **PDF2MD**. Choose a document with **Browse…**, or drag it into the window.
3. Select a worksheet for Excel. For PDF, keep **auto** or choose an engine manually.
4. Click **Convert**. Output is saved next to the source: `document.pdf` → `document.md`; images go into `document_images`.
5. When conversion finishes, the progress bar is replaced by **Open result** and **Open folder**. These open the Markdown file in your Windows default application or its containing folder. The **×** button beside these actions returns the window to its initial view: it dismisses the result, resets the bar and collapses the log. Saved files, the selected document and preferences are kept. Move its image folder along with it to preserve links.

If output already exists, the application asks whether to replace it, save a numbered copy or cancel. **Cancel** stops the current operation. The application does not replace the source document.

## Features

- PDF headings, lists, bordered tables and images; optional OCR for scans.
- Word tables, images, comments and supported equations as LaTeX.
- One selected Excel sheet, merged cells, comments, images and wide tables with data kept in its original rows.
- CSV/TSV/TXT encoding detection; CSV delimiter detection.
- Progress, cancellation, diagnostic log and protection against overwriting the source.
- **Settings → Language → Русский / English**, without restarting. Language, engine, options and last folder are remembered.

## PDF engines

| Engine | Use for | Limitations |
| --- | --- | --- |
| `auto` | Structured for text PDFs; Marker for detected scanned pages when installed | Heuristic selection; check complex results |
| `structured` | Text-layer manuals and books, bordered tables | No OCR; borderless tables and equations may lose structure |
| `marker` | Scans, complex layouts, equations and borderless tables | Downloads models; may change numbers or merge tables; slow on CPU |
| `pymupdf` | Fast plain-text extraction | No table structure, images or OCR |

Auto examines every page, including mixed PDFs, and ignores repeated transparent chapter artwork. Without Marker, a warning explains that scans may have no extracted text.

## Optional Marker installation

Choose **Settings → Install Marker…**, select **CPU** or **NVIDIA CUDA 12.6**, and choose a folder. CUDA needs a compatible NVIDIA driver. Downloads require several GB and an internet connection. After setup, document processing runs locally.

Complex PDFs can take about a minute per page on CPU. Model loading adds startup time. Page count, layout, OCR and hardware matter more than file size alone.

## Settings and storage

The installer defaults to Program Files for all users. Use `/CURRENTUSER` for a per-user installation. Uninstall through Windows Apps.

The light portable build needs no installation, but uses the Windows user profile: settings and logs are in `%LOCALAPPDATA%\PDF2MD`. Optional Marker files are in the selected folder, defaulting to `C:\ProgramData\PDF2MD`. Copying the portable ZIP does not transfer settings or a separately installed Marker engine.

Uninstalling keeps documents, settings and separately downloaded models.

## Limitations

Check important tables, values and equations against the source. Word headers and footers are not exported. Excel formulas use saved cached values; open and save the workbook in Excel if the cache is missing. Excel numbers retain Russian formatting regardless of interface language.

Batch processing and exporting all worksheets at once are not yet supported. The system file picker follows the Windows language; third-party diagnostics may use their original language. Wide HTML tables need a compatible Markdown viewer.

## Source, license and support

Free download; open source under **AGPL-3.0-only**. See [LICENSE.txt](LICENSE.txt), [COPYRIGHT.txt](COPYRIGHT.txt) and [THIRD_PARTY.md](THIRD_PARTY.md).

[Build instructions](docs/BUILDING.md) · [Version 0.2.2 notes](docs/releases/0.2.2.md)

[Report an issue](https://github.com/RootforgeIO/PDF2MD/issues) with the application version, Windows version, document format and reproduction steps. Use a sanitized sample instead of private documents, and review logs before sharing.

Author: [RootforgeIO](https://github.com/RootforgeIO).
