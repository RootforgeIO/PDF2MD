from __future__ import annotations

from i18n import tr, get_language

import argparse
import contextlib
import errno
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

ProgressCallback = Callable[["ProgressUpdate"], None]
MarkerCommand = list[str]

_TQDM_PERCENT_RE = re.compile(r"(\d+)\s*%")
_TQDM_FRACTION_RE = re.compile(r"(\d+)\s*/\s*(\d+)")


@dataclass(frozen=True)
class ProgressUpdate:
    percent: float
    message: str
    eta_seconds: Optional[float] = None
    elapsed_seconds: Optional[float] = None
    current_page: Optional[int] = None
    total_pages: Optional[int] = None


@dataclass(frozen=True)
class PdfInfo:
    page_count: int
    file_size_bytes: int
    probably_scanned: bool


class PdfOpenError(Exception):
    """PDF cannot be opened or read (damaged, not a PDF, password-protected)."""


def validate_output(source: Path, target: Path) -> None:
    """Reject aliases too: case differences, symlinks and hard links."""
    same = source.resolve() == target.resolve()
    if not same and target.exists():
        same = source.samefile(target)
    if same:
        raise OSError(errno.EINVAL, tr('Исходный файл и результат должны быть разными файлами'), str(target))


def write_markdown(target: Path, text: str) -> None:
    """Publish complete UTF-8 output; a failed write leaves the old file intact."""
    fd, name = tempfile.mkstemp(prefix=".pdf2md-", suffix=".tmp", dir=target.parent)
    pending = Path(name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as stream:
            stream.write(text)
        try:
            os.replace(pending, target)
        except OSError as exc:
            raise type(exc)(exc.errno, exc.strerror, str(target)) from exc
    finally:
        pending.unlink(missing_ok=True)


def _pymupdf():
    """Import PyMuPDF; `fitz` alias is deprecated since PyMuPDF 1.24.3 / 2.x."""
    try:
        import pymupdf  # type: ignore

        return pymupdf
    except ImportError:
        import fitz  # type: ignore

        return fitz


def _open_pdf(pdf_path: Path):
    """Open a PDF and fail early with a readable message instead of a traceback."""
    lib = _pymupdf()
    try:
        doc = lib.open(pdf_path)
    except Exception as e:
        raise PdfOpenError(tr('Не удалось открыть PDF (файл повреждён или это не PDF): {v0}').format(v0=e)) from e
    if doc.needs_pass:
        doc.close()
        raise PdfOpenError(tr('PDF защищён паролем — снимите защиту и повторите.'))
    return doc


def get_pdf_info(pdf_path: Path) -> PdfInfo:
    doc = _open_pdf(pdf_path)
    try:
        page_count = doc.page_count
    finally:
        doc.close()

    return PdfInfo(
        page_count=page_count,
        file_size_bytes=pdf_path.stat().st_size,
        probably_scanned=_is_probably_scanned_pdf(pdf_path),
    )


def _portable_root() -> Optional[Path]:
    root = os.environ.get("PDF2MD_PORTABLE_ROOT")
    if root:
        return Path(root)
    return None


def _marker_importable() -> bool:
    """True if `marker` (marker-pdf) is importable by the current interpreter.

    The wrapper marker_single_portable.py always sits next to pdf2md.py, so its
    presence alone does not mean Marker is installed. Without this check `auto`
    resolves to Marker in a plain dev venv, shows a misleading Marker estimate,
    fails and silently falls back to PyMuPDF.
    """
    try:
        return importlib.util.find_spec("marker") is not None
    except (ImportError, ValueError):
        return False


def _portable_marker_wrapper() -> Optional[Path]:
    if not _marker_importable():
        return None
    root = _portable_root()
    candidates = []
    if root is not None:
        candidates.append(root / "app" / "marker_single_portable.py")
    candidates.append(Path(__file__).with_name("marker_single_portable.py"))

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return None


def find_marker_single() -> Optional[str]:
    marker_single = os.environ.get("PDF2MD_MARKER_SINGLE")
    if marker_single and Path(marker_single).exists():
        return marker_single

    wrapper = _portable_marker_wrapper()
    if wrapper is not None:
        return str(wrapper)

    marker_single = shutil.which("marker_single")
    if marker_single:
        return marker_single

    for name in ("marker_single.exe", "marker_single"):
        local_marker = Path(sys.executable).with_name(name)
        if local_marker.exists():
            return str(local_marker)

    return None


def find_marker_command() -> Optional[MarkerCommand]:
    marker_runner = os.environ.get("PDF2MD_MARKER_RUNNER")
    if marker_runner and Path(marker_runner).exists():
        return [marker_runner, "--marker-single"]

    marker_single = os.environ.get("PDF2MD_MARKER_SINGLE")
    if marker_single and Path(marker_single).exists():
        return [marker_single]

    # Only the setup build loads an optional Marker runtime from the user's settings.
    from marker_install import installed_build, installed_marker_command

    if installed_build():
        return installed_marker_command()

    wrapper = _portable_marker_wrapper()
    if wrapper is not None:
        return [sys.executable, str(wrapper)]

    marker_single = shutil.which("marker_single")
    if marker_single:
        return [marker_single]

    for name in ("marker_single.exe", "marker_single"):
        local_marker = Path(sys.executable).with_name(name)
        if local_marker.exists():
            return [str(local_marker)]

    return None


def marker_available() -> bool:
    return find_marker_command() is not None


ENGINES = ("auto", "marker", "structured", "pymupdf")

_pymupdf4llm_rag = None


def _load_structured_backend():
    """Classic pymupdf4llm backend (find_tables + headers + images).

    The optional `pymupdf_layout` add-on is deliberately NOT used: it is licensed
    Polyform Noncommercial and in tests merged borderless tables into paragraphs,
    dropped images and silently started Tesseract OCR.
    """
    global _pymupdf4llm_rag
    if _pymupdf4llm_rag is not None:
        return _pymupdf4llm_rag
    import contextlib
    import io

    # pymupdf4llm prints an advert for pymupdf_layout to stdout on import.
    with contextlib.redirect_stdout(io.StringIO()):
        import pymupdf4llm  # type: ignore
    if hasattr(pymupdf4llm, "use_layout"):
        try:
            pymupdf4llm.use_layout(False)
        except Exception:
            pass
    from pymupdf4llm.helpers import pymupdf_rag  # type: ignore

    _patch_table_reading_order(pymupdf_rag)
    _patch_column_boxes_for_images(pymupdf_rag)
    _pymupdf4llm_rag = pymupdf_rag
    return pymupdf_rag


# pymupdf4llm emits a table before a text block only if the table lies entirely
# ABOVE it. On two-column pages a tall table in the left column is never above
# any right-column block, so it was dumped at the end of the page. Also emit
# tables that lie entirely LEFT of the block (previous column), and never emit
# a table from the right column ahead of left-column text.
_TABLE_ORDER_OLD = "[j for j in parms.tab_rects.items() if j[1].y1 <= text_rect.y0]"
_TABLE_ORDER_NEW = (
    "[j for j in parms.tab_rects.items() if (j[1].y1 <= text_rect.y0 and j[1].x0 < text_rect.x1)"
    " or j[1].x1 <= text_rect.x0]"
)
# Same idea for images: an illustration in the left column next to right-column
# text used to require horizontal overlap with the text block, so it was
# dumped at the end of the page. Allow images left of the block; still skip
# images below it or entirely to its right. The 20 pt tolerance (about one
# line) pairs with the cut placement in _patch_column_boxes_for_images.
_IMAGE_ORDER_OLD = (
    "                if img_rect.y0 > text_rect.y0:\n"
    "                    continue\n"
    "                if img_rect.x0 >= text_rect.x1 or img_rect.x1 <= text_rect.x0:\n"
    "                    continue"
)
_IMAGE_ORDER_NEW = (
    "                if img_rect.y0 > text_rect.y0 + 20:\n"
    "                    continue\n"
    "                if img_rect.x0 >= text_rect.x1:\n"
    "                    continue"
)
# Record where each saved illustration came from, so text that exists only
# inside a drawing (boxed callouts, diagram labels) can be added as text.
_SAVE_IMAGE_OLD = "            pix.save(image_filename)\n"
_SAVE_IMAGE_NEW = (
    "            pix.save(image_filename)\n"
    "            PDF2MD_SAVED_IMAGES.append((image_filename.replace(\"\\\\\", \"/\"), pymupdf.Rect(rect)))\n"
)
structured_patch_notes: list[str] = []


def _patch_table_reading_order(rag) -> None:
    """Source-level patch of a nested function (pinned pymupdf4llm 0.3.4).

    Applied only if the exact expression is found; otherwise the library is
    left untouched and a note is recorded.
    """
    import inspect

    try:
        source = inspect.getsource(rag)
    except (OSError, TypeError):
        structured_patch_notes.append("table-order patch skipped: no source")
        return
    patched = source
    for name, old, new in (
        ("table-order", _TABLE_ORDER_OLD, _TABLE_ORDER_NEW),
        ("image-order", _IMAGE_ORDER_OLD, _IMAGE_ORDER_NEW),
        ("image-text", _SAVE_IMAGE_OLD, _SAVE_IMAGE_NEW),
    ):
        if patched.count(old) == 1:
            patched = patched.replace(old, new)
        else:
            structured_patch_notes.append(f"{name} patch skipped: pymupdf4llm changed")
    rag.PDF2MD_SAVED_IMAGES = []
    if patched == source:
        return
    exec(compile(patched, getattr(rag, "__file__", "<pymupdf_rag>"), "exec"), rag.__dict__)


@contextlib.contextmanager
def _structured_table_patches():
    """During a structured conversion: drop drawing-like "tables" and render
    tables ourselves (merged cells -> HTML, never copied from the neighbour)."""
    import pdf2md_tables

    lib = _pymupdf()
    page_cls = lib.Page
    table_cls = lib.table.Table
    orig_find = page_cls.find_tables
    orig_md = table_cls.to_markdown

    def find_tables(self, *args, **kwargs):
        finder = orig_find(self, *args, **kwargs)
        pdf2md_tables.drop_fake_tables(self, finder)
        return finder

    def to_markdown(self, clean=False, fill_empty=True):
        return pdf2md_tables.render_pymupdf_table(self)

    page_cls.find_tables = find_tables
    table_cls.to_markdown = to_markdown
    try:
        yield
    finally:
        page_cls.find_tables = orig_find
        table_cls.to_markdown = orig_md


def _patch_column_boxes_for_images(rag) -> None:
    """Keep illustrations in reading order.

    pymupdf4llm emits an image/graphic before the first text block whose top
    is at or below the image top. Its layout step (column_boxes) often returns
    one tall block for a whole column, so every illustration on the page ended
    up after all the text. We cut text blocks at the top edge of each
    illustration (raster image or vector-graphics cluster) that sits in the
    same or the previous column. Cuts are placed in the gap between text lines,
    so no line is split or lost.
    """
    original = getattr(rag, "column_boxes", None)
    if original is None or getattr(original, "_pdf2md_patched", False):
        return
    lib = _pymupdf()

    def column_boxes_cut_at_illustrations(page, *args, **kwargs):
        rects = original(page, *args, **kwargs)
        try:
            anchors = [lib.Rect(i["bbox"]) for i in page.get_image_info()]
            anchors += [lib.Rect(r) for r in (kwargs.get("avoid") or [])]
            anchors = [a for a in anchors if not a.is_empty and a.height > 10]
            if not anchors or not rects:
                return rects
            line_boxes = []
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", ()):
                    line_boxes.append(lib.Rect(line["bbox"]))
        except Exception:
            return rects

        result = []
        for rect in rects:
            rect = lib.Rect(rect)
            cuts = set()
            for a in anchors:
                if a.x0 >= rect.x1:  # entirely to the right: a later column
                    continue
                if not (rect.y0 + 2 < a.y0 < rect.y1 - 2):
                    continue
                # Cut at the illustration's top edge; if that edge crosses a
                # text line, move the cut above that line (never split a line).
                cut = a.y0
                moved = True
                while moved:
                    moved = False
                    for ln in line_boxes:
                        if ln.x1 > rect.x0 and ln.x0 < rect.x1 and ln.y0 < cut < ln.y1:
                            cut = ln.y0 - 0.1
                            moved = True
                if rect.y0 + 1 < cut < rect.y1 - 1:
                    cuts.add(round(cut, 1))
            y = rect.y0
            for cut in sorted(cuts):
                if cut - y > 1:
                    result.append(lib.Rect(rect.x0, y, rect.x1, cut))
                    y = cut
            result.append(lib.Rect(rect.x0, y, rect.x1, rect.y1))
        return result

    column_boxes_cut_at_illustrations._pdf2md_patched = True  # type: ignore[attr-defined]
    rag.column_boxes = column_boxes_cut_at_illustrations


_WORD_RE = re.compile(r"[^\W_]{2,}", re.U)


def _norm_words(text: str) -> str:
    return " ".join(w.lower() for w in _WORD_RE.findall(text))


def _illustration_text(page, rect, chunk_norm: str) -> str:
    """Text that sits inside an illustration and is missing from the Markdown.

    Returned as a blockquote (one paragraph per text block) or "" when the
    region has no meaningful text or every paragraph is already in the output
    (compared as normalized word sequences, not single words).
    """
    try:
        blocks = page.get_text("blocks", clip=rect, sort=True)
    except Exception:
        return ""
    paragraphs = []
    for b in blocks:
        if len(b) > 6 and b[6] != 0:  # image block
            continue
        text = re.sub(r"-\n(?=[a-zа-я])", "", b[4])
        text = " ".join(text.split())
        if not text:
            continue
        probe = _norm_words(text)[:80]
        if probe and probe in chunk_norm:
            continue
        paragraphs.append(text)
    if sum(len(_WORD_RE.findall(p_)) for p_ in paragraphs) < 4:
        return ""
    return "\n>\n".join(f"> {p_}" for p_ in paragraphs)


def _attach_illustration_text(page, chunk: str, saved: list) -> str:
    """Insert the text of each saved illustration right under its image link."""
    if not saved:
        return chunk
    chunk_norm = _norm_words(chunk)
    for path, rect in saved:
        quote = _illustration_text(page, rect, chunk_norm)
        if not quote:
            continue
        link = f"]({path})"
        pos = chunk.find(link)
        if pos < 0:
            continue
        end = chunk.find("\n", pos)
        end = len(chunk) if end < 0 else end
        chunk = chunk[:end] + "\n\n" + quote + "\n" + chunk[end:]
        chunk_norm += " " + _norm_words(quote)
    return chunk


_EDGE_LINES = 2


def _plain(line: str) -> str:
    return " ".join(re.sub(r"[*_#`]", " ", line).split())


def _header_key(line: str, label: Optional[str] = None) -> str:
    """Normalize a line for running-header detection: no markup; the page's own
    printed number (if known) is removed, other digits are kept."""
    t = _plain(line)
    if label is not None:
        t = re.sub(rf"\b{label}\b", " ", t)
    t = re.sub(r"[|`>\-–—.:·•]", " ", t)
    return " ".join(t.lower().split()) if len(t.split()) <= 10 else ""


def _is_page_number(line: str) -> bool:
    t = re.sub(r"[*_#|\s]", "", line)
    return bool(re.fullmatch(r"(?:page|стр\.?|с\.)?\d{1,4}(?:/\d{1,4})?", t, re.I))


def _edge_line_indexes(lines: list[str]) -> list[int]:
    """Indexes of the first/last few text lines (not images, comments, tables)."""
    idx = [
        i for i, ln in enumerate(lines)
        if ln.strip() and not ln.lstrip().startswith(("![", "<!--", "|", "<t", ">"))
    ]
    return sorted(set(idx[:_EDGE_LINES] + idx[-_EDGE_LINES:]))


def _page_label_offset(pages_lines: list[list[str]]) -> Optional[int]:
    """Most common (printed page number - page index) over page-edge numbers.

    Lets us remove "12", "**12**" or "12 | Section" only when 12 really is the
    printed number of that page, never an arbitrary number in the text.
    """
    votes: dict[int, int] = {}
    for i, lines in enumerate(pages_lines):
        cand = set(_edge_line_indexes(lines)) | {
            j for j, ln in enumerate(lines) if _is_page_number(ln)
        }
        offsets = set()
        for j in cand:
            plain = _plain(lines[j])
            if len(plain.split()) > 8:
                continue
            for num in re.findall(r"\b\d{1,4}\b", plain):
                offsets.add(int(num) - (i + 1))
        for off in offsets:
            votes[off] = votes.get(off, 0) + 1
    if not votes:
        return None
    off, n = max(votes.items(), key=lambda kv: kv[1])
    return off if n >= max(3, int(len(pages_lines) * 0.3)) else None


def _labels_edge(plain: str, label: str) -> bool:
    """True for "12", "12 | Section", "Section · 12", "Page 12", "- 12 -"."""
    return bool(
        re.match(rf"^\W*(?:page|стр\.?|с\.)?\s*{label}\b", plain, re.I)
        or re.search(rf"\b{label}\W*$", plain)
    )


def strip_running_headers(chunks: list[str]) -> list[str]:
    """Remove running headers/footers and page numbers from per-page Markdown.

    * A short line among the first/last two text lines of a page that repeats
      there on >= 30 % of pages (min. 3); the page's own number is ignored in
      the comparison, other digits are not.
    * The printed page number (offset to the page index learned from the
      document): bare number lines anywhere, and short edge lines that start or
      end with it, such as "8 | Assembly Instructions".
    Body text is never touched.
    """
    if len(chunks) < 4:
        return chunks
    pages_lines = [c.split("\n") for c in chunks]
    offset = _page_label_offset(pages_lines)
    labels = [str(i + 1 + offset) if offset is not None else None for i in range(len(chunks))]

    counts: dict[str, int] = {}
    for lines, label in zip(pages_lines, labels):
        keys = {_header_key(lines[i], label) for i in _edge_line_indexes(lines)}
        for k in keys:
            if k:
                counts[k] = counts.get(k, 0) + 1
    threshold = max(3, int(len(chunks) * 0.3))
    running = {k for k, n in counts.items() if n >= threshold}

    result = []
    for lines, label in zip(pages_lines, labels):
        drop = set()
        for j in _edge_line_indexes(lines):
            key = _header_key(lines[j], label)
            plain = _plain(lines[j])
            if key and key in running:
                drop.add(j)
            elif label is not None and len(plain.split()) <= 8 and _labels_edge(plain, label):
                drop.add(j)
        if label is not None:
            for j, ln in enumerate(lines):
                if _is_page_number(ln) and re.sub(r"\D", "", ln) == label:
                    drop.add(j)
        kept = [ln for j, ln in enumerate(lines) if j not in drop]
        result.append(re.sub(r"\n{3,}", "\n\n", "\n".join(kept)).strip())
    return result


_HEADING_LINE_RE = re.compile(r"^(#{1,6}) (.+)$", re.M)


def _clean_heading(m: "re.Match[str]") -> str:
    # Headings are bold already; pymupdf4llm wraps each span in ** and can
    # leave fragments like "Panel Controls** **& Functions".
    text = re.sub(r"\s{2,}", " ", m.group(2).replace("**", "")).strip()
    return f"{m.group(1)} {text}"


def _tidy_markdown(md: str) -> str:
    """Cosmetic, meaning-preserving cleanup of pymupdf4llm output."""
    md = _HEADING_LINE_RE.sub(_clean_heading, md)
    md = re.sub(r"[ \t]+\n", "\n", md)
    # Blank line before headings and before a table that follows a paragraph line.
    md = re.sub(r"([^\n])\n(#{1,6} )", r"\1\n\n\2", md)
    md = re.sub(r"^(#{1,6} [^\n]+)\n(?=[^\n])", r"\1\n\n", md, flags=re.M)
    md = re.sub(r"(^|\n)([^|\n][^\n]*)\n(\|)", r"\1\2\n\n\3", md)
    md = re.sub(r"([^\n])\n<table>", r"\1\n\n<table>", md)
    md = re.sub(r"</table>\n(?=[^\n])", "</table>\n\n", md)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md


def structured_available() -> bool:
    try:
        return importlib.util.find_spec("pymupdf4llm") is not None
    except (ImportError, ValueError):
        return False


def resolve_engine(engine: str, info: Optional[PdfInfo] = None) -> str:
    if engine == "auto":
        if info is not None and info.probably_scanned and marker_available():
            return "marker"
        if structured_available():
            return "structured"
        return "pymupdf"
    return engine


# Marker: conservative estimates for CPU processing and complex layouts.
# Маркетинговые бенчмарки (0.2–2 с/стр. на H100) на домашнем ПК не применимы.
_MARKER_SEC_PER_PAGE_FAST = (8.0, 25.0)  # GPU, простой текст
_MARKER_SEC_PER_PAGE_TYPICAL = (25.0, 90.0)  # CPU или сложная вёрстка
_MARKER_SEC_PER_PAGE_SCAN = (50.0, 120.0)  # OCR / сканы


def estimate_duration_seconds(engine: str, info: PdfInfo) -> tuple[float, float]:
    """
    Rough min/max duration in seconds (min, max).
    Marker varies 10×+ by GPU, OCR, tables, first model load.
    """
    pages = max(info.page_count, 1)
    resolved = resolve_engine(engine, info)

    if resolved == "pymupdf":
        return (pages * 0.03, pages * 0.25)

    if resolved == "structured":
        # find_tables() + header analysis + image export; still CPU-light.
        return (pages * 0.05, pages * 0.6)

    if info.probably_scanned:
        lo, hi = _MARKER_SEC_PER_PAGE_SCAN
    else:
        lo, hi = _MARKER_SEC_PER_PAGE_TYPICAL

    return (pages * lo, pages * hi)


def format_duration(seconds: float) -> str:
    seconds = max(0, int(round(seconds)))
    if seconds < 60:
        return tr('~{v0} сек').format(v0=seconds)
    minutes, secs = divmod(seconds, 60)
    if minutes < 60:
        return tr('~{v0} мин {v1:02d} сек').format(v0=minutes, v1=secs) if secs else tr('~{v0} мин').format(v0=minutes)
    hours, minutes = divmod(minutes, 60)
    return tr('~{v0} ч {v1:02d} мин').format(v0=hours, v1=minutes)


def format_estimate_range(engine: str, info: PdfInfo) -> str:
    lo, hi = estimate_duration_seconds(engine, info)
    size_kb = info.file_size_bytes / 1024
    resolved = resolve_engine(engine, info)
    text = (
        tr('{v0} стр., {v1:.0f} КБ · движок {v2}: примерно {v3} – {v4}').format(v0=info.page_count, v1=size_kb, v2=resolved, v3=format_duration(lo), v4=format_duration(hi))
    )
    if resolved == "marker":
        text += tr(' (Marker: на CPU часто ~1 мин/стр.; GPU и без OCR — быстрее)')
    return text


def _is_probably_scanned_pdf(pdf_path: Path) -> bool:
    """Look for image-dominated pages without useful text, including mixed PDFs.

    Inspect all pages: sampling only the cover misses scans later in the file.
    Blank pages and sparse vector/title pages alone do not require OCR.
    """
    from collections import Counter

    with _open_pdf(pdf_path) as doc:
        # Reused transparent artwork under chapter headings is not a scanned page.
        uses = Counter(xref for page in doc for xref in {item[0] for item in page.get_images()})
        for page in doc:
            text = page.get_text("text").strip()
            if len(text) >= 40:
                continue
            images = page.get_image_info()
            resources = page.get_images()
            if (text and images and resources and all(item.get("has-mask") for item in images)
                    and all(uses[item[0]] >= 3 for item in resources)):
                continue
            area = abs(page.rect)
            image_area = sum(abs(_pymupdf().Rect(item["bbox"]) & page.rect)
                             for item in images)
            if area and image_area >= area * 0.5:
                return True
    return False


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".svg"}


def images_dir_for(out_md: Path) -> Path:
    """`document.md` -> `document_images/` next to it."""
    return out_md.parent / f"{out_md.stem}_images"


def _md_image_ref(images_dir: Path, filename: str) -> str:
    """Relative, forward-slash link; spaces encoded so Markdown viewers don't break."""
    return f"{images_dir.name}/{filename}".replace(" ", "%20")


def _remove_dir_if_empty(path: Path) -> None:
    try:
        if path.is_dir() and not any(path.iterdir()):
            path.rmdir()
    except OSError:
        pass


def _relocate_marker_images(md_text: str, src_dir: Path, out_md: Path) -> tuple[str, int]:
    """Copy images Marker wrote next to its .md into `<stem>_images/` and fix links."""
    images = [p for p in src_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES]
    if not images:
        return md_text, 0
    images_dir = images_dir_for(out_md)
    images_dir.mkdir(parents=True, exist_ok=True)
    for img in images:
        shutil.copy2(img, images_dir / img.name)
        ref = _md_image_ref(images_dir, img.name)
        # Marker links images by bare file name: ![](_page_0_Picture_1.jpeg)
        md_text = re.sub(
            r"(\]\()(?:\./)?" + re.escape(img.name) + r"(\))",
            lambda m: f"{m.group(1)}{ref}{m.group(2)}",
            md_text,
        )
    return md_text, len(images)


def _parse_progress_line(line: str) -> Optional[float]:
    m = _TQDM_PERCENT_RE.search(line)
    if m:
        return float(m.group(1))

    m = _TQDM_FRACTION_RE.search(line)
    if m:
        current, total = int(m.group(1)), int(m.group(2))
        if total > 0:
            return min(100.0, current / total * 100.0)
    return None


def _emit_progress(
    callback: Optional[ProgressCallback],
    *,
    percent: float,
    message: str,
    start_time: float,
    eta_seconds: Optional[float] = None,
    current_page: Optional[int] = None,
    total_pages: Optional[int] = None,
) -> None:
    if callback is None:
        return
    callback(
        ProgressUpdate(
            percent=max(0.0, min(100.0, percent)),
            message=message,
            eta_seconds=eta_seconds,
            elapsed_seconds=time.monotonic() - start_time,
            current_page=current_page,
            total_pages=total_pages,
        )
    )


class ConversionCancelled(Exception):
    """The user cancelled the conversion (nothing is written)."""


def _kill_process_tree(proc: subprocess.Popen) -> None:
    """Stop a child process and everything it started (Marker may spawn workers)."""
    try:
        if sys.platform == "win32":
            subprocess.run(
                ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                capture_output=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=15,
            )
        else:
            proc.kill()
        proc.wait(timeout=10)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def _cancellable(callback: Optional[ProgressCallback], cancel_event) -> Optional[ProgressCallback]:
    """Progress callback that raises ConversionCancelled once `cancel_event` is set.

    Every engine reports progress per page (Marker every 0.25 s), so this is where a
    conversion notices the cancel request.
    """
    if cancel_event is None:
        return callback

    def wrapped(update: ProgressUpdate) -> None:
        if cancel_event.is_set():
            raise ConversionCancelled()
        if callback is not None:
            callback(update)

    return wrapped


def _wait_marker_with_progress(
    proc: subprocess.Popen[str],
    *,
    callback: Optional[ProgressCallback],
    start_time: float,
    estimated_max: float,
    total_pages: int,
) -> int:
    parsed_percent: dict[str, Optional[float]] = {"value": None}
    stderr_tail: list[str] = []

    def read_stderr() -> None:
        if proc.stderr is None:
            return
        for line in proc.stderr:
            stripped = line.strip()
            if stripped:
                stderr_tail.append(stripped)
                if len(stderr_tail) > 12:
                    del stderr_tail[:-12]
            pct = _parse_progress_line(line)
            if pct is not None:
                parsed_percent["value"] = pct

    reader = threading.Thread(target=read_stderr, daemon=True)
    reader.start()

    last_percent = 0.0
    try:
        while proc.poll() is None:
            elapsed = time.monotonic() - start_time
            if parsed_percent["value"] is not None:
                pct = parsed_percent["value"]
                message = tr('Marker: обработка…')
            else:
                pct = min(92.0, (elapsed / max(estimated_max, 1.0)) * 100.0)
                message = tr('Marker: загрузка моделей и обработка…')

            pct = max(last_percent, min(95.0, pct))
            last_percent = pct
            _emit_progress(
                callback,
                percent=pct,
                message=message,
                start_time=start_time,
                eta_seconds=None,  # tqdm reports stages, not whole-document completion
                total_pages=total_pages,
            )
            time.sleep(0.25)

    finally:
        if proc.poll() is None:
            _kill_process_tree(proc)
        reader.join(timeout=5.0)
        if proc.stderr is not None:
            proc.stderr.close()
    if proc.returncode not in (None, 0) and stderr_tail:
        _emit_progress(
            callback,
            percent=last_percent,
            message=tr("ERROR: Сбой Marker: ") + " | ".join(stderr_tail[-4:])[:600],
            start_time=start_time,
            total_pages=total_pages,
        )
    return proc.returncode if proc.returncode is not None else 1


def convert_with_marker_cli(
    pdf_path: Path,
    out_md: Path,
    *,
    progress_callback: Optional[ProgressCallback] = None,
    pdf_info: Optional[PdfInfo] = None,
    save_images: bool = True,
) -> bool:
    """
    Uses Marker (marker-pdf) if `marker_single` is available in PATH.
    Produces high-fidelity markdown for complex layouts, tables, equations.
    """
    validate_output(pdf_path, out_md)
    marker_command = find_marker_command()
    if not marker_command:
        return False

    info = pdf_info or get_pdf_info(pdf_path)
    _est_lo, estimated_max = estimate_duration_seconds("marker", info)
    start_time = time.monotonic()

    _emit_progress(
        progress_callback,
        percent=0.0,
        message=tr('Marker: запуск…'),
        start_time=start_time,
        eta_seconds=estimated_max,
        total_pages=info.page_count,
    )

    out_md.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pdf2md_marker_", ignore_cleanup_errors=True) as tmpdir:
        tmpdir_path = Path(tmpdir)
        cmd = [
            *marker_command,
            str(pdf_path.resolve()),
            "--output_format",
            "markdown",
            "--output_dir",
            str(tmpdir_path),
        ]
        # Child Python writes stderr in UTF-8 (tqdm bars, Cyrillic paths); on
        # Windows the default would be the ANSI code page and decoding could fail.
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1", PDF2MD_LANGUAGE=get_language())
        popen_kwargs: dict = {}
        if sys.platform == "win32":
            # No console window flashing when marker_single.exe is started from pythonw.
            popen_kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
            **popen_kwargs,
        )
        try:
            returncode = _wait_marker_with_progress(
                proc,
                callback=progress_callback,
                start_time=start_time,
                estimated_max=estimated_max,
                total_pages=info.page_count,
            )
        finally:
            # Cancelled (the progress callback raised), Ctrl+C or any error: never leave Marker
            # running in the background holding the GPU.
            if proc.poll() is None:
                _kill_process_tree(proc)
        if returncode != 0:
            return False

        md_files = sorted(tmpdir_path.glob("*.md"))
        if not md_files:
            md_files = sorted(tmpdir_path.rglob("*.md"))
        if not md_files:
            return False

        md_text = md_files[0].read_text(encoding="utf-8")
        if save_images:
            md_text, n_images = _relocate_marker_images(md_text, md_files[0].parent, out_md)
            if n_images:
                _emit_progress(
                    progress_callback,
                    percent=99.0,
                    message=tr('INFO: картинок сохранено: {v0} → {v1}').format(v0=n_images, v1=images_dir_for(out_md).name),
                    start_time=start_time,
                    total_pages=info.page_count,
                )
        write_markdown(out_md, md_text)

    _emit_progress(
        progress_callback,
        percent=100.0,
        message=tr('Marker: готово'),
        start_time=start_time,
        eta_seconds=0.0,
        total_pages=info.page_count,
    )
    return True


def convert_with_pymupdf_text(
    pdf_path: Path,
    *,
    page_markers: bool = True,
    page_separators: bool = True,
    progress_callback: Optional[ProgressCallback] = None,
    pdf_info: Optional[PdfInfo] = None,
) -> str:
    """
    Lightweight, fully-local conversion: extract text and emit simple Markdown.
    This avoids inventing content, but won't preserve complex layout well.
    """
    info = pdf_info or get_pdf_info(pdf_path)
    total_pages = max(info.page_count, 1)
    _, estimated_max = estimate_duration_seconds("pymupdf", info)
    start_time = time.monotonic()

    doc = _open_pdf(pdf_path)
    try:
        parts: list[str] = []
        for page_index in range(doc.page_count):
            page = doc.load_page(page_index)
            text = page.get_text("text") or ""
            text = text.replace("\r\n", "\n").replace("\r", "\n").rstrip()
            if page_markers:
                parts.append(f"<!-- page: {page_index + 1} -->\n{text}".rstrip())
            else:
                parts.append(text.rstrip())

            done = page_index + 1
            pct = done / total_pages * 100.0
            elapsed = time.monotonic() - start_time
            eta = (elapsed / done) * (total_pages - done) if done else estimated_max
            _emit_progress(
                progress_callback,
                percent=pct,
                message=tr('PyMuPDF: страница {v0}/{v1}').format(v0=done, v1=total_pages),
                start_time=start_time,
                eta_seconds=eta,
                current_page=done,
                total_pages=total_pages,
            )

        sep = "\n\n---\n\n" if page_separators else "\n\n"
        result = sep.join(parts).strip() + "\n"
    finally:
        doc.close()

    _emit_progress(
        progress_callback,
        percent=100.0,
        message=tr('PyMuPDF: готово'),
        start_time=start_time,
        eta_seconds=0.0,
        total_pages=total_pages,
    )
    return result


def convert_with_structured(
    pdf_path: Path,
    out_md: Path,
    *,
    page_markers: bool = True,
    page_separators: bool = True,
    save_images: bool = True,
    strip_headers: bool = True,
    progress_callback: Optional[ProgressCallback] = None,
    pdf_info: Optional[PdfInfo] = None,
) -> str:
    """
    Fast structured conversion (pymupdf4llm, classic mode): headings, lists,
    bold/italic, ruled tables as Markdown tables, images saved to
    `<stem>_images/`. No ML models, seconds per hundred pages.

    Limitation: borderless tables come out as aligned text lines — use Marker
    for those.
    """
    rag = _load_structured_backend()
    info = pdf_info or get_pdf_info(pdf_path)
    total_pages = max(info.page_count, 1)
    _, estimated_max = estimate_duration_seconds("structured", info)
    start_time = time.monotonic()

    images_dir = images_dir_for(out_md)
    if save_images:
        images_dir.mkdir(parents=True, exist_ok=True)
    abs_prefix = str(images_dir).replace("\\", "/") + "/"

    doc = _open_pdf(pdf_path)
    try:
        with _structured_table_patches():
            # Header font sizes are computed once for the whole document; per-page
            # calls then reuse it (otherwise every call rescans the whole PDF).
            hdr_info = rag.IdentifyHeaders(doc)
            parts: list[str] = []
            saved_images = getattr(rag, "PDF2MD_SAVED_IMAGES", None)
            for page_index in range(doc.page_count):
                if saved_images is not None:
                    saved_images.clear()
                chunk = rag.to_markdown(
                    doc,
                    pages=[page_index],
                    hdr_info=hdr_info,
                    write_images=save_images,
                    ignore_images=not save_images,
                    image_path=str(images_dir),
                    image_format="png",
                    filename=out_md.stem,
                    dpi=150,
                    table_strategy="lines_strict",
                    force_text=True,
                    show_progress=False,
                )
                if save_images and saved_images:
                    chunk = _attach_illustration_text(doc[page_index], chunk, list(saved_images))
                if save_images:
                    chunk = chunk.replace(abs_prefix, _md_image_ref(images_dir, ""))
                parts.append(_tidy_markdown(chunk).strip())

                done = page_index + 1
                elapsed = time.monotonic() - start_time
                _emit_progress(
                    progress_callback,
                    percent=done / total_pages * 100.0,
                    message=tr('Structured: страница {v0}/{v1}').format(v0=done, v1=total_pages),
                    start_time=start_time,
                    eta_seconds=(elapsed / done) * (total_pages - done) if done else estimated_max,
                    current_page=done,
                    total_pages=total_pages,
                )
    finally:
        doc.close()

    n_images = 0
    if save_images:
        n_images = sum(1 for p in images_dir.iterdir() if p.is_file()) if images_dir.is_dir() else 0
        _remove_dir_if_empty(images_dir)

    if strip_headers:
        parts = strip_running_headers(parts)
    if page_markers:
        parts = [f"<!-- page: {i + 1} -->\n{c}".rstrip() for i, c in enumerate(parts)]

    sep = "\n\n---\n\n" if page_separators else "\n\n"
    result = sep.join(p for p in parts if p).strip() + "\n"
    _emit_progress(
        progress_callback,
        percent=100.0,
        message=tr('Structured: готово (картинок: {v0})').format(v0=n_images) if save_images else tr('Structured: готово'),
        start_time=start_time,
        eta_seconds=0.0,
        total_pages=total_pages,
    )
    return result


def convert_pdf(
    pdf_path: Path,
    out_md: Path,
    *,
    engine: str = "auto",
    page_markers: bool = True,
    page_separators: bool = True,
    save_images: bool = True,
    strip_headers: bool = True,
    progress_callback: Optional[ProgressCallback] = None,
) -> tuple[bool, str]:
    """
    Returns (success, engine_used).

    Auto uses Structured for text and Marker for image-only pages when available.
    An explicitly chosen engine never silently falls back.
    """
    validate_output(pdf_path, out_md)
    info = get_pdf_info(pdf_path)
    resolved = resolve_engine(engine, info)
    out_md.parent.mkdir(parents=True, exist_ok=True)

    if resolved == "marker":
        ok = convert_with_marker_cli(
            pdf_path,
            out_md,
            progress_callback=progress_callback,
            pdf_info=info,
            save_images=save_images,
        )
        if ok:
            return True, "marker"
        if engine == "marker":
            return False, "marker"
        resolved = "structured" if structured_available() else "pymupdf"

    if info.probably_scanned and progress_callback:
        _emit_progress(
            progress_callback,
            percent=0.0,
            message=tr('WARN: PDF похож на скан — без OCR (Marker) текста может не быть'),
            start_time=time.monotonic(),
            total_pages=info.page_count,
        )

    if resolved == "structured":
        try:
            md = convert_with_structured(
                pdf_path,
                out_md,
                page_markers=page_markers,
                page_separators=page_separators,
                save_images=save_images,
                strip_headers=strip_headers,
                progress_callback=progress_callback,
                pdf_info=info,
            )
        except ImportError:
            if engine == "structured":
                raise
            md = None
        if md is not None:
            write_markdown(out_md, md)
            return True, "structured"

    md = convert_with_pymupdf_text(
        pdf_path,
        page_markers=page_markers,
        page_separators=page_separators,
        progress_callback=progress_callback,
        pdf_info=info,
    )
    write_markdown(out_md, md)
    return True, "pymupdf"


OFFICE_SUFFIXES = (".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".txt")
SUPPORTED_SUFFIXES = (".pdf",) + OFFICE_SUFFIXES


def is_supported_file(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_SUFFIXES


def convert_file(
    src: Path,
    out_md: Path,
    *,
    engine: str = "auto",
    sheet: Optional[str] = None,
    page_markers: bool = True,
    page_separators: bool = True,
    save_images: bool = True,
    strip_headers: bool = True,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_event: Optional[threading.Event] = None,
) -> tuple[bool, str]:
    """
    Any supported file -> Markdown. Returns (success, converter_used).

    PDF goes through `convert_pdf` (engine choice applies); DOCX/XLSX/CSV/TXT through
    `office2md` (`sheet` applies to Excel; raises `office2md.OfficeError` on bad files).
    Setting `cancel_event` stops the conversion with `ConversionCancelled`; the .md is not written.
    """
    validate_output(src, out_md)
    progress_callback = _cancellable(progress_callback, cancel_event)
    suffix = src.suffix.lower()
    if suffix == ".pdf":
        return convert_pdf(
            src,
            out_md,
            engine=engine,
            page_markers=page_markers,
            page_separators=page_separators,
            save_images=save_images,
            strip_headers=strip_headers,
            progress_callback=progress_callback,
        )
    if suffix not in OFFICE_SUFFIXES:
        raise ValueError(tr('Неподдерживаемый тип файла: {v0}').format(v0=src.suffix or src.name))

    import office2md

    start_time = time.monotonic()

    def _step(percent: float, message: str) -> None:
        _emit_progress(progress_callback, percent=percent, message=message, start_time=start_time, eta_seconds=0.0)

    out_md.parent.mkdir(parents=True, exist_ok=True)
    md = office2md.convert_office(src, out_md, sheet=sheet, save_images=save_images, progress=_step)
    write_markdown(out_md, md)
    return True, office2md.office_kind(src) or suffix.lstrip(".")


def describe_write_error(e: OSError) -> str:
    """Readable text for a failed write of the result (.md / images)."""
    where = f": {e.filename}" if getattr(e, "filename", None) else ""
    if isinstance(e, PermissionError):
        return (
            tr('нет доступа к файлу{v0} — он открыт в другой программе, только для чтения или папка защищена от записи').format(v0=where)
        )
    if isinstance(e, IsADirectoryError):
        return tr('вместо файла указана папка{v0}').format(v0=where)
    return tr('не удалось записать результат{v0}: {v1}').format(v0=where, v1=e.strerror or e)


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="pdf2md",
        description="Convert PDF, Word (.docx), Excel (.xlsx), CSV/TSV or TXT to Markdown "
        "into a separate .md file.",
    )
    p.add_argument("pdf", metavar="input", type=Path, help="Input file: " + ", ".join(SUPPORTED_SUFFIXES))
    p.add_argument(
        "-o",
        "--out",
        type=Path,
        default=None,
        help="Output .md path (default: next to the input, same name)",
    )
    p.add_argument(
        "--sheet",
        default=None,
        help="Excel: sheet name or 1-based number (default: the active sheet)",
    )
    p.add_argument(
        "--engine",
        choices=list(ENGINES),
        default="auto",
        help="Conversion engine (auto: Structured for text, Marker for scans when available)",
    )
    p.add_argument("--no-images", action="store_true", help="Do not save images (marker/structured)")
    p.add_argument("--language", choices=("ru", "en"), default="ru", help="Language of application messages")
    p.add_argument("--no-page-markers", action="store_true", help="Omit <!-- page: N --> comments")
    p.add_argument("--no-page-separators", action="store_true", help="Omit --- between pages")
    p.add_argument(
        "--keep-headers", action="store_true", help="Keep running headers/footers and page numbers (structured)"
    )
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    from i18n import set_language
    set_language(args.language)
    pdf_path: Path = args.pdf
    if not pdf_path.is_file() or not is_supported_file(pdf_path):
        print(
            f"ERROR: input must be an existing file of type {', '.join(SUPPORTED_SUFFIXES)}: {pdf_path}",
            file=sys.stderr,
        )
        return 2

    out_md: Path = args.out if args.out else pdf_path.with_suffix(".md")
    if args.out and (out_md.is_dir() or str(args.out).endswith(("/", "\\"))):
        out_md = out_md / f"{pdf_path.stem}.md"  # -o DIR: same name inside that folder
    if pdf_path.suffix.lower() == ".pdf":
        try:
            info = get_pdf_info(pdf_path)
        except PdfOpenError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return 4
        print(f"Estimate: {format_estimate_range(args.engine, info)}", file=sys.stderr)

    def _cli_progress(update: ProgressUpdate) -> None:
        eta = format_duration(update.eta_seconds) if update.eta_seconds is not None else "?"
        print(tr('\r[{v0:5.1f}%] {v1} · осталось {v2}').format(v0=update.percent, v1=update.message, v2=eta), end="", file=sys.stderr)

    from office2md import OfficeError

    try:
        ok, used = convert_file(
            pdf_path,
            out_md,
            engine=args.engine,
            sheet=args.sheet,
            page_markers=not args.no_page_markers,
            page_separators=not args.no_page_separators,
            save_images=not args.no_images,
            strip_headers=not args.keep_headers,
            progress_callback=_cli_progress,
        )
    except (PdfOpenError, OfficeError) as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        return 4
    except OSError as e:
        print(f"\nERROR: {describe_write_error(e)}", file=sys.stderr)
        return 5
    except KeyboardInterrupt:
        print(tr('\nПрервано.'), file=sys.stderr)
        return 130
    print(file=sys.stderr)
    if not ok:
        print(
            "ERROR: marker engine requested, but `marker_single` failed or is not installed.",
            file=sys.stderr,
        )
        return 3

    print(f"Done ({used}): {out_md}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
