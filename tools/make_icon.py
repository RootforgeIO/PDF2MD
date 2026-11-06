"""Generate assets/pdf2md.ico (all Windows sizes) and assets/pdf2md.png from pdf2md_theme.make_app_icon.

    python tools/make_icon.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from pdf2md_theme import make_app_icon  # noqa: E402


def main() -> int:
    assets = ROOT / "assets"
    assets.mkdir(exist_ok=True)
    big = make_app_icon(256)
    big.save(assets / "pdf2md.png")
    sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
    # small sizes drawn natively (not downscaled from 256) stay sharp
    frames = [make_app_icon(s) for s in sizes]
    frames[-1].save(assets / "pdf2md.ico", format="ICO", sizes=[(s, s) for s in sizes], append_images=frames[:-1])
    print("written:", assets / "pdf2md.ico", assets / "pdf2md.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
