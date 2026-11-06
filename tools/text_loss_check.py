"""Compare words in a PDF with words in its Markdown output.

Shows how much text a conversion lost overall and on which pages.

    python tools/text_loss_check.py "samples\\Korg MicroKORG XL.pdf" "samples\\Korg MicroKORG XL.md"
    python tools/text_loss_check.py file.pdf file.md --pages 10

Notes:
* Image links and HTML tags are ignored; text inside HTML tables counts.
* Hyphenated words joined by the converter ("out-let" -> "outlet") show up as
  missing "out"/"let" — a known false positive of this simple check.
* Removed running headers/footers also count as "missing" (on purpose).
"""

from __future__ import annotations

import argparse
import collections
import html
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

_WORD = re.compile(r"[A-Za-zА-Яа-яЁё0-9]{3,}")


def _tokens(text: str) -> list[str]:
    return [w.lower() for w in _WORD.findall(text)]


def _clean_md(text: str) -> str:
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", " ", text)
    return re.sub(r"<[^>]+>", " ", html.unescape(text))


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="PDF vs Markdown text loss check")
    ap.add_argument("pdf", type=Path)
    ap.add_argument("md", type=Path)
    ap.add_argument("--pages", type=int, default=8, help="show N worst pages")
    args = ap.parse_args(argv)

    import pdf2md

    lib = pdf2md._pymupdf()
    doc = lib.open(args.pdf)
    md_text = args.md.read_text(encoding="utf-8")

    src = collections.Counter(w for p in doc for w in _tokens(p.get_text()))
    out = collections.Counter(_tokens(_clean_md(md_text)))
    missing = {w: c - out[w] for w, c in src.items() if c > out[w]}
    total, lost = sum(src.values()), sum(missing.values())
    print(f"{args.pdf.name}: words {total}, missing {lost} ({lost / max(total, 1):.2%})")
    print("most missing:", sorted(missing.items(), key=lambda kv: -kv[1])[:15])

    chunks = re.split(r"<!-- page: (\d+) -->", md_text)
    if len(chunks) > 1:
        per_page = {int(chunks[i]): chunks[i + 1] for i in range(1, len(chunks), 2)}
        rows = []
        for p in doc:
            s = collections.Counter(_tokens(p.get_text()))
            o = collections.Counter(_tokens(_clean_md(per_page.get(p.number + 1, ""))))
            miss = sum(max(0, c - o[w]) for w, c in s.items())
            rows.append((miss, p.number + 1, [w for w, c in s.items() if c > o[w]][:8]))
        print(f"worst pages (missing, page, words):")
        for row in sorted(rows, reverse=True)[: args.pages]:
            print("  ", row)
    else:
        print("(no <!-- page: N --> markers — per-page breakdown skipped)")
    doc.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
