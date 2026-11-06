"""Generate checksums and a compact manifest for final public release files."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    parser.add_argument("--version", default="0.2.0")
    parser.add_argument("--source-commit", required=True)
    args = parser.parse_args()
    if not re.fullmatch(r"[0-9a-f]{40}", args.source_commit):
        parser.error("Provide the full public source commit SHA")
    names = [f"PDF2MD-{args.version}-{suffix}" for suffix in ("setup.exe", "portable.zip", "source.zip")]
    artifacts = [{"file": name, "bytes": (args.directory / name).stat().st_size,
                  "sha256": sha256(args.directory / name)} for name in names]
    manifest = {
        "product": "PDF2MD", "version": args.version, "publisher": "RootforgeIO",
        "platform": "Windows 10 1809+ / Windows 11 x64", "license": "AGPL-3.0-only",
        "source_repository": "https://github.com/RootforgeIO/PDF2MD",
        "source_commit": args.source_commit, "signed": False,
        "bundled_marker": False, "artifacts": artifacts,
    }
    (args.directory / "release-manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    (args.directory / "SHA256SUMS.txt").write_text(
        "".join(f"{item['sha256']}  {item['file']}\n" for item in artifacts), encoding="ascii")
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    main()
