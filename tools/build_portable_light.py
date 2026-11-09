"""Package the light Windows runtime without installing or registering it."""
from __future__ import annotations

import argparse
import re
import sys
import sysconfig
import tempfile
import zipfile
from pathlib import Path

from build_windows_installer import ROOT, build_payload


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", default="0.2.2")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dist")
    args = parser.parse_args()
    if (sys.platform != "win32" or sys.version_info[:2] != (3, 12)
            or sys.version_info.micro < 4 or sysconfig.get_platform() != "win-amd64"):
        parser.error("Use Windows x64 Python 3.12.4 or later in the 3.12 series")
    if not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", args.version):
        parser.error("Invalid version")
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    archive = output / f"PDF2MD-{args.version}-portable.zip"
    if archive.exists():
        parser.error(f"Output already exists: {archive}")
    with tempfile.TemporaryDirectory(prefix="pdf2md-light-", dir=output) as staging:
        payload = Path(staging)
        # The managed-runtime flag enables optional Marker installation and
        # per-user preferences; no registry or shortcut is created by the ZIP.
        build_payload(payload)
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as bundle:
            for file in sorted(payload.rglob("*")):
                if file.is_file():
                    bundle.write(file, Path("PDF2MD") / file.relative_to(payload))
    print(f"{archive} ({archive.stat().st_size / 1024**2:.1f} MiB)")


if __name__ == "__main__":
    main()
