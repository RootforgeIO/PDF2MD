"""Build the x64 Windows installer from a working PDF2MD Python 3.12 runtime."""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import sysconfig
import tempfile
from importlib.metadata import distributions
from pathlib import Path

from pip._vendor.packaging.requirements import Requirement
from pip._vendor.packaging.utils import canonicalize_name


ROOT = Path(__file__).resolve().parents[1]
APP_FILES = (
    "pdf2md.py", "pdf2md_ui.py", "pdf2md_tables.py", "pdf2md_theme.py",
    "office2md.py", "marker_single_portable.py", "marker_install.py",
    "app_settings.py", "i18n.py", "ui_dialogs.py", "translations_en.json",
)
DOC_FILES = (
    "README.md", "README.en.md", "FAQ.md", "CHANGELOG.md", "PRIVACY.md",
    "LICENSE.txt", "COPYRIGHT.txt", "THIRD_PARTY.md", "requirements-release.txt",
)


def copy_light_packages(source: Path, destination: Path) -> None:
    available = {canonicalize_name(d.metadata["Name"]): d for d in distributions(path=[str(source)])}
    pending = [
        Requirement(line) for line in (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]
    pending.append(Requirement("pip"))
    copied = set()
    while pending:
        requirement = pending.pop()
        if requirement.marker and not requirement.marker.evaluate({"extra": ""}):
            continue
        name = canonicalize_name(requirement.name)
        package = available.get(name)
        if package is None or package.version not in requirement.specifier or not package.files:
            raise RuntimeError(f"The build runtime needs an installed wheel for {requirement}")
        if name in copied:
            continue
        copied.add(name)
        for relative in package.files:
            if relative.name == "direct_url.json":
                continue  # installation provenance may contain local machine paths
            if set(relative.parts) & {"..", "__pycache__", "mupdf-devel"}:
                continue
            path = (source / relative).resolve()
            if not path.is_relative_to(source):
                raise RuntimeError(f"Unexpected wheel path: {relative}")
            target = destination / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
        pending.extend(Requirement(req) for req in package.requires or ())


def build_payload(payload: Path) -> None:
    python = Path(sys.base_prefix).resolve()
    runtime = payload / "runtime"
    runtime.mkdir()
    for name in ("python.exe", "pythonw.exe", "python312.dll", "python3.dll", "LICENSE.txt"):
        shutil.copy2(python / name, runtime / name)
    for dll in python.glob("vcruntime*.dll"):
        shutil.copy2(dll, runtime / dll.name)
    for name in ("Lib", "DLLs", "tcl"):
        shutil.copytree(
            python / name, runtime / name,
            ignore=shutil.ignore_patterns("site-packages", "__pycache__", "test", "turtledemo"),
        )
    source_site = Path(sysconfig.get_path("purelib")).resolve()
    copy_light_packages(source_site, runtime / "Lib" / "site-packages")
    app = payload / "app"
    app.mkdir()
    for name in APP_FILES:
        shutil.copy2(ROOT / name, app / name)
    (app / "installed.flag").touch()
    shutil.copytree(ROOT / "assets", app / "assets")
    for name in DOC_FILES:
        shutil.copy2(ROOT / name, payload / name)
    for name in ("LICENSES", "docs"):
        shutil.copytree(ROOT / name, payload / name)
    compiler = Path(os.environ["WINDIR"]) / "Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    subprocess.run([
        str(compiler), "/nologo", "/target:winexe", "/platform:x64", "/define:INSTALLED",
        f"/out:{payload / 'PDF2MD.exe'}", f"/win32icon:{ROOT / 'assets/pdf2md.ico'}",
        "/reference:System.Windows.Forms.dll", str(ROOT / "portable_launcher.cs"),
    ], check=True)
    subprocess.run([str(runtime / "python.exe"), "-I", "-B", "-c",
        "import tkinter, tkinterdnd2, pymupdf, pymupdf4llm, tabulate, "
        "mammoth, openpyxl, markdownify, bs4, PIL, pip; "
        "from importlib.util import find_spec; "
        "assert not any(find_spec(n) for n in ('marker', 'torch', 'pymupdf_layout'))",
    ], check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--iscc", required=True, type=Path, help="Path to Inno Setup's ISCC.exe")
    parser.add_argument("--version", default="0.2.0")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dist")
    args = parser.parse_args()
    if sys.platform != "win32" or sys.version_info[:2] != (3, 12) or sys.version_info.micro < 4:
        parser.error("Use Windows x64 Python 3.12.4 or later in the 3.12 series")
    if sys.maxsize <= 2**32 or sysconfig.get_platform() != "win-amd64":
        parser.error("Use a Windows x64 Python runtime")
    if not args.iscc.is_file() or not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", args.version):
        parser.error("Provide an existing ISCC.exe and a version such as 0.1.0")

    output = args.output_dir.resolve()
    output.mkdir(exist_ok=True, parents=True)
    with tempfile.TemporaryDirectory(prefix="pdf2md-setup-", dir=output) as staging:
        payload = Path(staging)
        build_payload(payload)
        subprocess.run([
            str(args.iscc.resolve()), "/Q", f"/DPayloadDir={payload}", f"/DOutputPath={output}",
            f"/DAppVersion={args.version}", str(ROOT / "installer/PDF2MD.iss"),
        ], check=True)
    installer = output / f"PDF2MD-{args.version}-setup.exe"
    print(f"{installer} ({installer.stat().st_size / 1024**2:.1f} MiB)")


if __name__ == "__main__":
    main()
