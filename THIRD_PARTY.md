# Сторонние компоненты / Third-party components

PDF2MD использует перечисленные ниже компоненты. Их авторские права и условия сохраняются отдельно от лицензии PDF2MD. Таблица описывает лёгкие Windows-сборки 0.2.0; Marker не включён и устанавливается отдельно.

PDF2MD uses the following components under their respective copyrights and licenses. This table describes the light Windows 0.2.0 builds; Marker is an optional separate download.

| Component | Version | License | Upstream |
| --- | --- | --- | --- |
| beautifulsoup4 | 4.14.3 | MIT | [Source / project](https://www.crummy.com/software/BeautifulSoup/) |
| cobble | 0.1.4 | BSD-2-Clause | [Source / project](https://github.com/mwilliamson/python-cobble) |
| et_xmlfile | 2.0.0 | MIT | [Source / project](https://foss.heptapod.net/openpyxl/et_xmlfile) |
| mammoth | 1.13.0 | BSD-2-Clause | [Source / project](https://github.com/mwilliamson/python-mammoth) |
| markdownify | 1.2.2 | MIT | [Source / project](https://github.com/matthewwithanm/python-markdownify) |
| openpyxl | 3.1.5 | MIT | [Source / project](https://openpyxl.readthedocs.io/) |
| pillow | 10.4.0 | HPND | [Source / project](https://python-pillow.org/) |
| pip | 26.1.2 | MIT | [Source / project](https://github.com/pypa/pip) |
| PyMuPDF | 1.27.2.3 | AGPL-3.0 / Artifex commercial | [Source / project](https://github.com/pymupdf/PyMuPDF) |
| pymupdf4llm | 0.3.4 | AGPL-3.0 / Artifex commercial | [Source / project](https://github.com/pymupdf/pymupdf4llm) |
| six | 1.17.0 | MIT | [Source / project](https://github.com/benjaminp/six) |
| soupsieve | 2.8.4 | MIT | [Source / project](https://github.com/facelessuser/soupsieve) |
| tabulate | 0.10.0 | MIT | [Source / project](https://github.com/astanin/python-tabulate) |
| tkinterdnd2 | 0.7.0 | MIT | [Source / project](https://github.com/Eliav2/tkinterdnd2) |
| typing_extensions | 4.15.0 | PSF-2.0 | [Source / project](https://github.com/python/typing_extensions) |
| Python | 3.12.10 | Python license and included notices | [Python](https://www.python.org/downloads/release/python-31210/) |
| Tcl/Tk | 8.6 runtime | Tcl/Tk license | [Tcl/Tk](https://www.tcl.tk/) |

## License files and source

Upstream license texts supplied in the installed packages are collected under `LICENSES/packages/`. Python and Tcl/Tk notices are included in `LICENSES/`. Package metadata and additional bundled notices also remain in `runtime/Lib/site-packages/` in the binary distributions. Pip includes separately licensed vendored components; their supplied notices are included.

PyMuPDF and PyMuPDF4LLM are offered under AGPL or an Artifex commercial license. These builds use the open-source distributions; they do not include a commercial Artifex license. The GNU AGPL v3 text is included in `LICENSE.txt` and in the PyMuPDF4LLM license directory. Upstream licensing: [PyMuPDF](https://pymupdf.io/licensing).

Cobble 0.1.4 declares BSD-2-Clause in its [upstream README](https://github.com/mwilliamson/python-cobble/blob/master/README.rst); its wheel does not ship a separate license file. The upstream attribution and reference are recorded in `LICENSES/cobble-NOTICE.txt`.

For the application's source and build scripts, use the matching source archive or release tag in [RootforgeIO/PDF2MD](https://github.com/RootforgeIO/PDF2MD). Component sources are available from the linked upstream projects; versions are listed above and pinned in `requirements-release.txt`. Upstream components are copied from their installed distributions, without modifying their source files.

## Optional Marker

The optional installer requests Marker 1.10.2, Surya OCR 0.17.1, PyTorch 2.7.1, torchvision 0.22.1, transformers 4.57.6 and pdftext 0.6.3, plus their dependencies. CPU and NVIDIA CUDA 12.6 builds are supported.

Marker and Surya code use GPL-3.0-or-later according to package metadata. PyTorch, torchvision, transformers and the remaining dependencies retain their own terms. Model weights can have different licenses from the code. Read their upstream terms when using or redistributing a downloaded engine; weights are not included in the light release archives.

- [Marker source and model licensing](https://github.com/datalab-to/marker)
- [Surya source and model licensing](https://github.com/datalab-to/surya)
- [PyTorch](https://pytorch.org/)
- [Transformers](https://github.com/huggingface/transformers)

## Packaging

The Windows installer is built with [Inno Setup](https://jrsoftware.org/isinfo.php). The launcher uses Microsoft .NET Framework; it is not a bundled copy of .NET. Microsoft runtime DLLs distributed with Python retain their vendor terms.
