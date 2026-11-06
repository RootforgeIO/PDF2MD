"""Word / Excel / CSV / TXT -> Markdown.

* DOCX: `mammoth` -> HTML (headings by Word styles, lists, bold/italic, tables with
  colspan/rowspan, images) -> Markdown via `markdownify`. Tables are taken out of the
  HTML first and rendered by `pdf2md_tables.render_table` (regular -> pipe table,
  merged cells -> HTML), images go to `<stem>_images/`.
* XLSX/XLSM: one sheet (`openpyxl`, cached formula values), merged cells -> spans,
  empty edges trimmed, dates/percents/decimals as Excel displays them, sheet images
  saved next to the .md.
* CSV/TSV: delimiter sniffed, encoding detected -> one table.
* TXT: encoding detected, text as is (line endings normalised).

Like the PDF engines: text is only extracted and marked up, never rewritten.
"""

from __future__ import annotations

from i18n import tr

import csv
import datetime as dt
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from functools import lru_cache
from pathlib import Path
from typing import Callable, Optional

from pdf2md_tables import Cell, Grid, render_table

# suffix -> converter kind
OFFICE_SUFFIXES = {
    ".docx": "docx",
    ".doc": "doc",  # Word 97-2003: converted to .docx by the installed Microsoft Word
    ".xlsx": "xlsx",
    ".xlsm": "xlsx",
    ".xls": "xls",  # Excel 97-2003: converted to .xlsx by the installed Microsoft Excel
    ".csv": "csv",
    ".tsv": "csv",
    ".txt": "txt",
}

# (percent, message) — pdf2md wraps it into its ProgressUpdate
StepCallback = Callable[[float, str], None]


class OfficeError(Exception):
    """Readable error for files that cannot be converted (bad file, missing library)."""


def office_kind(path: Path) -> Optional[str]:
    return OFFICE_SUFFIXES.get(path.suffix.lower())


def _require(module: str, package: str):
    try:
        return __import__(module)
    except ImportError as e:
        raise OfficeError(
            tr('Не установлена библиотека {v0}: pip install {v1}').format(v0=package, v1=package)
        ) from e


def _step(callback: Optional[StepCallback], percent: float, message: str) -> None:
    if callback:
        callback(percent, message)


def _tidy(md: str) -> str:
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"[ \t]+\n", "\n", md)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"


# --- Text encodings ----------------------------------------------------------

_BOMS = (
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32"),
    (b"\x00\x00\xfe\xff", "utf-32"),
    (b"\xff\xfe", "utf-16"),
    (b"\xfe\xff", "utf-16"),
)


def decode_text(data: bytes) -> tuple[str, str]:
    """Bytes -> (text, encoding): BOM -> UTF-8 -> BOM-less UTF-16 -> CP1251."""
    for bom, enc in _BOMS:
        if data.startswith(bom):
            return data.decode(enc), enc
    # BOM-less UTF-16 first: ASCII in UTF-16 is also "valid" UTF-8 (with NULs), which real
    # text never contains. High bytes of Latin/Cyrillic (0x00/0x04) sit at odd offsets in LE.
    sample = data[:4096]
    if len(sample) >= 4 and len(data) % 2 == 0 and sample.count(0) >= len(sample) // 8:
        high_odd = sum(1 for b in sample[1::2] if b in (0, 4))
        high_even = sum(1 for b in sample[0::2] if b in (0, 4))
        enc = "utf-16-le" if high_odd >= high_even else "utf-16-be"
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            pass
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    return data.decode("cp1251", errors="replace"), "cp1251"


def read_text_file(path: Path) -> tuple[str, str]:
    text, enc = decode_text(path.read_bytes())
    return text.replace("\r\n", "\n").replace("\r", "\n"), enc


# --- TXT -----------------------------------------------------------------------


def convert_txt(path: Path, *, progress: Optional[StepCallback] = None) -> str:
    text, enc = read_text_file(path)
    _step(progress, 100.0, tr('TXT: готово (кодировка {v0})').format(v0=enc))
    return text.rstrip("\n") + "\n" if text.strip() else ""


# --- CSV / TSV -----------------------------------------------------------------

_CSV_DELIMITERS = ",;\t|"


def _sniff_delimiter(text: str, suffix: str) -> str:
    if suffix == ".tsv":
        return "\t"
    sample = "\n".join(text.split("\n")[:50])
    try:
        return csv.Sniffer().sniff(sample, delimiters=_CSV_DELIMITERS).delimiter
    except csv.Error:
        first = sample.split("\n", 1)[0]
        counts = {d: first.count(d) for d in _CSV_DELIMITERS}
        best = max(counts, key=counts.get)  # type: ignore[arg-type]
        return best if counts[best] else ","


def convert_csv(path: Path, *, progress: Optional[StepCallback] = None) -> str:
    text, enc = read_text_file(path)
    delimiter = _sniff_delimiter(text, path.suffix.lower())
    rows = [r for r in csv.reader(io.StringIO(text), delimiter=delimiter) if any(c.strip() for c in r)]
    shown = {"\t": "TAB", ",": ",", ";": ";", "|": "|"}.get(delimiter, delimiter)
    _step(progress, 100.0, tr('CSV: строк {v0}, разделитель «{v1}», кодировка {v2}').format(v0=len(rows), v1=shown, v2=enc))
    if not rows:
        return ""
    grid: Grid = [[Cell(c.strip()) for c in r] for r in rows]
    return render_table(grid)


# --- XLSX ----------------------------------------------------------------------


# Rows: data separated by at least this many empty rows becomes separate tables.
# Columns: a run of at least this many empty columns inside a table is not shown (a few
# cells far to the right must stay in their rows, without hundreds of empty columns between).
BLOCK_GAP = 5


def _load_workbook(path: Path, *, data_only: bool = True, read_only: bool = True):
    """read_only streams one sheet's XML: ~10x faster than a full load on big books,
    but has no merged cells / images — those come from `_sheet_xml_extras` / a full load."""
    openpyxl = _require("openpyxl", "openpyxl")
    try:
        return openpyxl.load_workbook(path, data_only=data_only, read_only=read_only)
    except Exception as e:
        raise OfficeError(tr('Не удалось открыть Excel-файл (повреждён или не .xlsx): {v0}').format(v0=e)) from e


def _worksheet_names(wb) -> tuple[list[str], int]:
    """Data sheets only (chart sheets have no cells) and the index of the active one."""
    names = [ws.title for ws in wb.worksheets]
    if not names:
        raise OfficeError(tr('В книге нет листов с данными (только диаграммы)'))
    active = wb.active.title if wb.active is not None else names[0]
    return names, names.index(active) if active in names else 0


def list_sheets(path: Path) -> tuple[list[str], int]:
    """(sheet names, index of the active sheet) — for the UI sheet selector."""
    if path.suffix.lower() == ".xls":
        path = xls_as_xlsx(path)
    wb = _load_workbook(path)
    try:
        return _worksheet_names(wb)
    finally:
        wb.close()


def resolve_sheet(names: list[str], active: int, sheet: Optional[str]) -> str:
    """`None`/"" -> active sheet; exact name; else a 1-based number."""
    if not sheet:
        return names[active]
    if sheet in names:
        return sheet
    lowered = {n.casefold(): n for n in names}
    if sheet.casefold() in lowered:
        return lowered[sheet.casefold()]
    if sheet.strip().isdigit() and 1 <= int(sheet) <= len(names):
        return names[int(sheet) - 1]
    raise OfficeError(tr('Нет листа «{v0}». Листы: ').format(v0=sheet) + ", ".join(names))


_DATE_TOKEN_RE = re.compile(r"yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|am/pm", re.I)


def _strip_format_noise(fmt: str) -> str:
    fmt = re.sub(r"\[[^\]]*\]", "", fmt)  # [$-419], [Red], [h]
    fmt = re.sub(r'"[^"]*"', "", fmt)
    return fmt.replace("\\", "").split(";")[0]


def _format_datetime(value, fmt: str) -> str:
    """Excel date format -> strftime-like output; ISO when the format is unusual."""
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S" if value.second else "%H:%M")
    if isinstance(value, dt.timedelta):
        total = int(value.total_seconds())
        return f"{total // 3600}:{total % 3600 // 60:02d}:{total % 60:02d}"
    if not isinstance(value, dt.datetime):  # dt.date
        value = dt.datetime(value.year, value.month, value.day)
    clean = _strip_format_noise(fmt).lower()
    has_date = any(t in clean for t in ("y", "d")) or clean in ("", "general")
    has_time = any(t in clean for t in ("h", "s"))
    if not has_date and not has_time:
        has_date = True
    if clean in ("", "general", "mm-dd-yy"):  # built-in short date: locale-dependent, use ISO
        out = value.strftime("%Y-%m-%d")
        if value.time() != dt.time(0, 0):
            out += value.strftime(" %H:%M")
        return out

    out, pos = [], 0
    prev = None
    for m in _DATE_TOKEN_RE.finditer(clean):
        out.append(clean[pos : m.start()])
        tok = m.group(0)
        if tok in ("m", "mm"):
            # minutes after hours / before seconds, otherwise month
            following = clean[m.end() :].lstrip(":.")
            is_minute = prev in ("h", "hh") or following.startswith("s")
            num = value.minute if is_minute else value.month
            out.append(f"{num:02d}" if tok == "mm" else str(num))
        elif tok == "yyyy":
            out.append(f"{value.year:04d}")
        elif tok == "yy":
            out.append(f"{value.year % 100:02d}")
        elif tok == "mmmm":
            out.append(value.strftime("%B"))
        elif tok == "mmm":
            out.append(value.strftime("%b"))
        elif tok == "dd":
            out.append(f"{value.day:02d}")
        elif tok == "d":
            out.append(str(value.day))
        elif tok == "hh":
            out.append(f"{value.hour:02d}")
        elif tok == "h":
            out.append(str(value.hour))
        elif tok == "ss":
            out.append(f"{value.second:02d}")
        elif tok == "s":
            out.append(str(value.second))
        else:  # am/pm
            out.append("AM" if value.hour < 12 else "PM")
        prev = tok
        pos = m.end()
    out.append(clean[pos:])
    text = "".join(out).strip()
    return text or value.strftime("%Y-%m-%d")


# Numbers are shown the way Russian-locale Excel shows them: 1 234 567,89 / 12,5%.
DECIMAL_SEP = ","
THOUSANDS_SEP = " "

_GROUPING_RE = re.compile(r"[#0?],[#0?]")  # "#,##0" — digit grouping
_SCALING_RE = re.compile(r"[#0?](,+)(?![#0?,])")  # "#,##0," — divide by 1000 per trailing comma
_DIGIT = "0#?"


@dataclass(frozen=True)
class _NumSection:
    """One section of an Excel number format: literal text around a numeric core."""

    prefix: str
    core: str  # format codes only: 0 # ? . , % E+ / …
    suffix: str
    conditional: bool  # "[>100]"-style condition — sections then cannot be picked by sign

    @property
    def is_general(self) -> bool:
        return not any(ch in _DIGIT for ch in self.core)


def _parse_section(section: str) -> _NumSection:
    """Split a format section into literal prefix / numeric core / literal suffix.

    Literals: "quoted", \\x, [$₽-419] currency, plain non-code characters. `_x` (padding) and
    `*x` (fill) only align columns in Excel — skipped.
    """
    tokens: list[tuple[str, str]] = []  # ("lit"|"code", text)
    conditional = False
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            end = section.find('"', i + 1)
            end = len(section) if end < 0 else end
            tokens.append(("lit", section[i + 1 : end]))
            i = end + 1
        elif ch == "\\" and i + 1 < len(section):
            tokens.append(("lit", section[i + 1]))
            i += 2
        elif ch in "_*" and i + 1 < len(section):
            i += 2
        elif ch == "[":
            end = section.find("]", i)
            end = len(section) if end < 0 else end
            inner = section[i + 1 : end]
            if inner.startswith("$"):
                symbol = inner[1:].split("-", 1)[0]
                if symbol:
                    tokens.append(("lit", symbol))
            elif inner[:1] in "<>=":
                conditional = True
            i = end + 1
        elif ch in _DIGIT + ".,%/" or (ch in "Ee" and section[i + 1 : i + 2] in ("+", "-")):
            if ch in "Ee":
                tokens.append(("code", section[i : i + 2]))
                i += 2
            else:
                tokens.append(("code", ch))
                i += 1
        elif ch == "@" or section[i : i + 7].lower() == "general":
            n = 1 if ch == "@" else 7
            tokens.append(("code", "General"))
            i += n
        else:
            tokens.append(("lit", ch))
            i += 1

    digit_idx = [k for k, (kind, text) in enumerate(tokens) if kind == "code" and (text[0] in _DIGIT or text == "General")]
    if not digit_idx:
        lits = "".join(t for kind, t in tokens if kind == "lit")
        codes = "".join(t for kind, t in tokens if kind == "code")
        return _NumSection(lits, codes, "", conditional)
    first, last = digit_idx[0], digit_idx[-1]
    # codes right after the last digit ("%", trailing "," scaling) belong to the core
    while last + 1 < len(tokens) and tokens[last + 1][0] == "code":
        last += 1
    prefix = "".join(t for kind, t in tokens[:first] if kind == "lit")
    core = "".join(t for kind, t in tokens[first : last + 1] if kind == "code")
    suffix = "".join(t for kind, t in tokens[last + 1 :] if kind == "lit")
    # a "%" before the digits still makes a percent format
    if any(kind == "code" and t == "%" for kind, t in tokens[:first]) and "%" not in core:
        core += "%"
    return _NumSection(prefix, core, suffix, conditional)


@lru_cache(maxsize=1024)
def _parse_number_format(number_format: str) -> tuple[_NumSection, ...]:
    return tuple(_parse_section(s) for s in number_format.split(";"))


def _round_half_up(value: Decimal, decimals: int) -> Decimal:
    """Round half away from zero on the shortest repr, like Excel (2.675 -> 2.68, not 2.67)."""
    return value.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)


def _general(value) -> str:
    if isinstance(value, float) and not value.is_integer():
        # Excel keeps 15 significant digits: 0.1+0.2 shows as 0.3, not 0.30000000000000004
        return format(value, ".15g").replace(".", DECIMAL_SEP)
    return str(int(value))


def _fraction(value: float, core: str) -> str:
    """'# ?/?' -> '1 1/2'; '?/8' fixed denominator; '#' before the space = whole part."""
    from fractions import Fraction

    num_part, _, den_part = core.partition("/")
    mixed = bool(re.search(r"[#0?]+\s*[#0?]+$", num_part)) and any(ch in "#0" for ch in num_part[:1])
    whole = int(value) if mixed else 0
    frac = Fraction(value - whole).limit_denominator(10 ** max(1, sum(ch in _DIGIT for ch in den_part)) - 1)
    if den_part.isdigit():
        frac = Fraction(round(float(value - whole) * int(den_part)), int(den_part))
    if frac == 1:
        whole, frac = whole + 1, Fraction(0)
    if frac == 0:
        return str(whole)
    return f"{whole} {frac.numerator}/{frac.denominator}" if whole else f"{frac.numerator}/{frac.denominator}"


def _format_number_core(value, core: str) -> str:
    """|value| through the numeric core of one section (no sign)."""
    if "E+" in core.upper() or "E-" in core.upper():
        mant = core.upper().split("E")[0]
        decimals = len(mant.split(".", 1)[1]) if "." in mant else 0
        text = f"{float(value):.{decimals}E}"
        mantissa, exp = text.split("E")
        return f"{mantissa}E{exp[0]}{abs(int(exp)):02d}".replace(".", DECIMAL_SEP)
    if "/" in core:
        return _fraction(float(value), core)

    number = Decimal(repr(float(value))) if isinstance(value, float) else Decimal(value)
    if "%" in core:
        number *= 100
    scaling = _SCALING_RE.search(core)
    if scaling:
        number /= 1000 ** len(scaling.group(1))
    int_codes, _, frac_codes = core.replace("%", "").partition(".")
    max_dec = sum(ch in _DIGIT for ch in frac_codes)
    min_dec = sum(ch == "0" for ch in frac_codes)
    number = _round_half_up(number, max_dec)
    text = f"{number:f}"
    if max_dec > min_dec and "." in text:  # optional decimals: "0.0#" shows 1.5 and 1.25
        head, frac = text.split(".")
        frac = frac.rstrip("0").ljust(min_dec, "0")
        text = f"{head}.{frac}" if frac else head
    if number == 0 and "0" not in int_codes and not min_dec:
        text = ""  # "#,###" / "??" show nothing for zero
    elif _GROUPING_RE.search(int_codes):
        text = _group_digits(text)
    return text.replace(".", DECIMAL_SEP) + ("%" if "%" in core else "")


def format_number(value, number_format: str = "General") -> str:
    """int/float as Excel shows it: sections by sign, literal text, currency, %, E+, fractions."""
    sections = _parse_number_format(number_format or "General")
    if any(s.conditional for s in sections):
        sections = sections[:1]  # "[>100]" conditions: keep it simple — the first section
    if value < 0 and len(sections) >= 2:
        section, sign = sections[1], ""  # Excel: the 2nd section shows |value|, its own sign text
    elif value == 0 and len(sections) >= 3:
        section, sign = sections[2], ""
    else:
        section, sign = sections[0], "-" if value < 0 else ""
    if section.is_general:
        text = _general(abs(value))
    else:
        text = _format_number_core(abs(value), section.core)
        if sign and not any(ch in "123456789" for ch in text):
            sign = ""  # no "-0,00"
    return f"{section.prefix}{sign}{text}{section.suffix}".strip()


def _group_digits(text: str) -> str:
    """'-1234567.89' -> '-1 234 567.89' (decimal point still '.')."""
    sign = "-" if text.startswith("-") else ""
    int_part, dot, frac = text.lstrip("-").partition(".")
    groups = []
    while len(int_part) > 3:
        groups.insert(0, int_part[-3:])
        int_part = int_part[:-3]
    groups.insert(0, int_part)
    return sign + THOUSANDS_SEP.join(groups) + dot + frac


def format_cell_value(value, number_format: str = "General") -> str:
    """Cell value as Excel shows it: dates, percents, fixed decimals, digit grouping."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        return _format_datetime(value, number_format or "General")
    if isinstance(value, (int, float)):
        return format_number(value, number_format or "General")
    return str(value).strip()


Coord = tuple[int, int]  # (row, column), 1-based


def _iter_cells(ws):
    """Non-empty cells of a (read-only) sheet. Dimensions in the file may be wrong — reset them."""
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()
    for row in ws.iter_rows():
        for cell in row:
            if getattr(cell, "value", None) is not None:
                yield cell


@dataclass
class SheetXmlInfo:
    merges: list[tuple[int, int, int, int]]  # (min_row, min_col, max_row, max_col)
    has_drawing: bool
    formulas_without_value: int  # <c><f>…</f></c> with no cached <v>: shown empty
    hidden_rows: set[int]
    hidden_cols: set[int]


def _sheet_xml_extras(wb, ws) -> SheetXmlInfo:
    """What openpyxl's read-only sheets do not expose, read from the sheet XML in one pass:
    merged ranges, drawings, formulas without a cached value, hidden rows and columns."""
    from xml.etree.ElementTree import iterparse

    from openpyxl.utils.cell import range_boundaries

    info = SheetXmlInfo([], False, 0, set(), set())
    with wb._archive.open(ws._worksheet_path) as src:
        for _event, el in iterparse(src):
            tag = el.tag.rsplit("}", 1)[-1]
            if tag == "c":
                kinds = {child.tag.rsplit("}", 1)[-1]: child for child in el}
                if "f" in kinds:
                    v = kinds.get("v")
                    # no <v>, or an empty <v> that is not a formula returning "" (t="str")
                    if v is None or (not (v.text or "") and el.get("t") != "str"):
                        info.formulas_without_value += 1
            elif tag == "row":
                if el.get("hidden") in ("1", "true") and el.get("r"):
                    info.hidden_rows.add(int(el.get("r")))
                el.clear()  # keep memory flat on big sheets
            elif tag == "col":
                if el.get("hidden") in ("1", "true"):
                    info.hidden_cols.update(range(int(el.get("min")), int(el.get("max")) + 1))
            elif tag == "mergeCell":
                min_col, min_row, max_col, max_row = range_boundaries(el.get("ref"))
                info.merges.append((min_row, min_col, max_row, max_col))
            elif tag == "drawing":
                info.has_drawing = True
    return info


def _runs_text(indexes: list[int], name) -> str:
    """[3, 7, 8, 9] -> '3, 7–9' (with `name` turning an index into its label)."""
    parts = []
    for a, b in _split_runs(sorted(indexes), 1):
        parts.append(name(a) if a == b else f"{name(a)}–{name(b)}")
    return ", ".join(parts)


def _split_runs(indexes: list[int], gap: int) -> list[tuple[int, int]]:
    """Sorted indexes -> (first, last) runs; a run breaks at >= `gap` missing indexes."""
    runs: list[tuple[int, int]] = []
    for i in indexes:
        if runs and i - runs[-1][1] - 1 < gap:
            runs[-1] = (runs[-1][0], i)
        else:
            runs.append((i, i))
    return runs


def _sheet_comments(wb, ws) -> dict[Coord, str]:
    """Cell notes (Excel comments) of a sheet, read from the package XML (read-only sheets lack them)."""
    import posixpath
    from xml.etree.ElementTree import fromstring

    from openpyxl.utils.cell import coordinate_to_tuple

    sheet_path = ws._worksheet_path
    rels_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
    try:
        rels = fromstring(wb._archive.read(rels_path))
    except KeyError:
        return {}
    notes: dict[Coord, str] = {}
    for rel in rels:
        if not rel.get("Type", "").endswith("/comments"):
            continue
        target = rel.get("Target", "")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
            posixpath.join(posixpath.dirname(sheet_path), target)
        )
        try:
            root = fromstring(wb._archive.read(part))
        except KeyError:
            continue
        for el in root.iter():
            if el.tag.rsplit("}", 1)[-1] != "comment":
                continue
            text = "".join(t.text or "" for t in el.iter() if t.tag.rsplit("}", 1)[-1] == "t").strip()
            if text and el.get("ref"):
                notes[coordinate_to_tuple(el.get("ref"))] = text
    return notes


@dataclass
class SheetBlock:
    """Rows min_row..max_row; `cols` — shown columns (ascending; wide empty runs left out)."""

    min_row: int
    max_row: int
    cols: list[int]

    def skipped(self) -> list[tuple[int, int]]:
        """Runs of columns not shown between shown ones: [(first, last), …]."""
        return [(a + 1, b - 1) for a, b in zip(self.cols, self.cols[1:]) if b - a > 1]

    def range_name(self) -> str:
        from openpyxl.utils import get_column_letter

        return (
            f"{get_column_letter(self.cols[0])}{self.min_row}:"
            f"{get_column_letter(self.cols[-1])}{self.max_row}"
        )


def sheet_blocks(
    values: dict[Coord, str],
    merges: list[tuple[int, int, int, int]],
    *,
    gap: int = BLOCK_GAP,
    extra: set[Coord] | frozenset = frozenset(),
) -> list[SheetBlock]:
    """Tables of a sheet, top to bottom.

    Rows split at >= `gap` empty rows. Inside a table every column with data is kept — also far
    to the right, in its own row — only runs of >= `gap` empty columns are left out.
    `extra` — positions that count as data although they have no value (e.g. notes).
    """
    occupied: set[Coord] = set(values) | set(extra)
    for r0, c0, r1, c1 in merges:  # a filled merged cell occupies its whole span
        if (r0, c0) in values:
            occupied.update((r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1))
    if not occupied:
        return []
    blocks = []
    for r_lo, r_hi in _split_runs(sorted({r for r, _ in occupied}), gap):
        used_cols = sorted({c for r, c in occupied if r_lo <= r <= r_hi})
        cols = [c for c_lo, c_hi in _split_runs(used_cols, gap) for c in range(c_lo, c_hi + 1)]
        blocks.append(SheetBlock(r_lo, r_hi, cols))
    return blocks


def block_grid(
    values: dict[Coord, str],
    merges: list[tuple[int, int, int, int]],
    block: SheetBlock,
    *,
    column_letters: bool = False,
) -> Grid:
    """Grid of one block; merged-cell spans clipped to the block and its shown columns.

    `column_letters` adds a first row with Excel column letters (A, B, …, ZE).
    """
    shown = set(block.cols)
    spans: dict[Coord, Coord] = {}  # top-left -> (rowspan, colspan)
    covered: dict[Coord, bool] = {}  # covered cell -> its top-left is shown
    for r0, c0, r1, c1 in merges:
        rows = range(max(r0, block.min_row), min(r1, block.max_row) + 1)
        cols = [c for c in block.cols if c0 <= c <= c1]
        if not rows or not cols:
            continue
        anchor_shown = block.min_row <= r0 and c0 in shown
        if anchor_shown:
            spans[(r0, c0)] = (len(rows), len(cols))
        for r in rows:
            for c in cols:
                if (r, c) != (r0, c0):
                    covered[(r, c)] = anchor_shown

    grid: Grid = []
    if column_letters:
        from openpyxl.utils import get_column_letter

        grid.append([Cell(get_column_letter(c)) for c in block.cols])
    for r in range(block.min_row, block.max_row + 1):
        out_row: list[Optional[Cell]] = []
        for c in block.cols:
            if (r, c) in covered:
                # A merge whose top-left is not shown -> plain empty cell.
                out_row.append(None if covered[(r, c)] else Cell(""))
                continue
            cell = Cell(values.get((r, c), ""))
            if (r, c) in spans:
                cell.rowspan, cell.colspan = spans[(r, c)]
            out_row.append(cell)
        grid.append(out_row)
    return grid


def _save_sheet_images(ws, out_md: Path, sheet_index: int) -> list[str]:
    """Save pictures anchored on the sheet; returns Markdown lines."""
    from pdf2md import _md_image_ref, images_dir_for  # late import: pdf2md imports this module

    from openpyxl.utils import get_column_letter

    lines = []
    images = list(getattr(ws, "_images", []))
    if not images:
        return lines
    images_dir = images_dir_for(out_md)
    for n, img in enumerate(images, start=1):
        try:
            data = img._data()
        except Exception:
            continue
        ext = (getattr(img, "format", None) or "png").lower().replace("jpeg", "jpg")
        name = f"sheet{sheet_index}_image{n}.{ext}"
        images_dir.mkdir(parents=True, exist_ok=True)
        (images_dir / name).write_bytes(data)
        where = ""
        marker = getattr(getattr(img, "anchor", None), "_from", None)
        if marker is not None:
            where = f" (ячейка {get_column_letter(marker.col + 1)}{marker.row + 1})"
        lines.append(f"![Картинка {n}{where}]({_md_image_ref(images_dir, name)})")
    return lines


def convert_xlsx(
    path: Path,
    out_md: Path,
    *,
    sheet: Optional[str] = None,
    save_images: bool = True,
    progress: Optional[StepCallback] = None,
) -> str:
    _step(progress, 5.0, tr('Excel: чтение книги…'))
    wb = _load_workbook(path)
    try:
        names, active = _worksheet_names(wb)
        name = resolve_sheet(names, active, sheet)
        ws = wb[name]
        _step(progress, 20.0, tr('Excel: лист «{v0}»').format(v0=name))

        xml_info = _sheet_xml_extras(wb, ws)
        merges = xml_info.merges
        covered = {
            (r, c)
            for r0, c0, r1, c1 in merges
            for r in range(r0, r1 + 1)
            for c in range(c0, c1 + 1)
            if (r, c) != (r0, c0)
        }
        values: dict[Coord, str] = {}
        for cell in _iter_cells(ws):
            pos = (cell.row, cell.column)
            if pos in covered:  # Excel shows only the top-left value of a merged cell
                continue
            text = format_cell_value(cell.value, cell.number_format)
            if text:
                values[pos] = text
        notes = _sheet_comments(wb, ws)
    finally:
        wb.close()
    _step(progress, 60.0, tr('Excel: лист «{v0}», ячеек с данными: {v1}, примечаний: {v2}').format(v0=name, v1=len(values), v2=len(notes)))

    if xml_info.formulas_without_value:
        _step(
            progress,
            70.0,
            tr('WARN: формул без сохранённого значения: {v0} — эти ячейки будут пустыми; откройте и сохраните файл в Excel').format(v0=xml_info.formulas_without_value),
        )

    from openpyxl.utils import get_column_letter

    # A note on a merged cell belongs to its top-left cell; on a covered cell — move it there.
    anchors = {
        (r, c): (r0, c0)
        for r0, c0, r1, c1 in merges
        for r in range(r0, r1 + 1)
        for c in range(c0, c1 + 1)
        if (r, c) != (r0, c0)
    }
    notes = {anchors.get(pos, pos): text for pos, text in sorted(notes.items())}

    blocks = sheet_blocks(values, merges, extra=set(notes))
    parts = [f"## {name}", ""]
    if not blocks:
        parts.append("*(лист пуст)*")
    note_no = 0
    for block in blocks:
        skipped = block.skipped()
        caption = []
        if len(blocks) > 1 or skipped:
            caption.append(f"Ячейки {block.range_name()}")
        if skipped:
            runs = ", ".join(
                get_column_letter(a) if a == b else f"{get_column_letter(a)}–{get_column_letter(b)}"
                for a, b in skipped
            )
            caption.append(f"пустые столбцы {runs} не показаны; первая строка — буквы столбцов Excel")
        hidden_rows = [r for r in xml_info.hidden_rows if block.min_row <= r <= block.max_row]
        hidden_cols = [c for c in xml_info.hidden_cols if c in block.cols]
        if hidden_rows or hidden_cols:
            hidden = []
            if hidden_rows:
                hidden.append("строки " + _runs_text(hidden_rows, str))
            if hidden_cols:
                hidden.append("столбцы " + _runs_text(hidden_cols, get_column_letter))
            caption.append("в Excel скрыты (здесь показаны): " + ", ".join(hidden))
        if caption:
            parts += ["*" + "; ".join(caption) + "*", ""]

        shown = dict(values)
        block_notes = []
        for (r, c), text in notes.items():
            if block.min_row <= r <= block.max_row and c in block.cols:
                note_no += 1
                shown[(r, c)] = f"{shown.get((r, c), '')} [прим. {note_no}]".strip()
                block_notes.append(f"{note_no}. **{get_column_letter(c)}{r}**: " + " ".join(text.split()))
        grid = block_grid(shown, merges, block, column_letters=bool(skipped))
        parts += [render_table(grid).rstrip("\n"), ""]
        if block_notes:
            parts += ["Примечания к ячейкам:", "", *block_notes, ""]
    if len(blocks) > 1:
        _step(progress, 85.0, tr('INFO: таблиц на листе: {v0} (разделены пустыми строками)').format(v0=len(blocks)))
    wide = [b for b in blocks if b.skipped()]
    if wide:
        _step(
            progress,
            86.0,
            tr('INFO: широкая таблица — данные далеко справа оставлены в своих строках, пустые столбцы между ними не показаны'),
        )

    if save_images and xml_info.has_drawing:
        # Images need a full load (slow on big books) — only for sheets that have drawings.
        _step(progress, 90.0, tr('Excel: картинки листа…'))
        wb_full = _load_workbook(path, read_only=False)
        try:
            image_lines = _save_sheet_images(wb_full[name], out_md, names.index(name) + 1)
        finally:
            wb_full.close()
        if image_lines:
            parts += ["", *image_lines]
    rows = sum(b.max_row - b.min_row + 1 for b in blocks)
    _step(progress, 100.0, tr('Excel: готово (строк {v0}, таблиц {v1}, примечаний {v2})').format(v0=rows, v1=len(blocks), v2=len(notes)))
    return _tidy("\n".join(parts))


# --- DOCX ----------------------------------------------------------------------

# mammoth's defaults already map "heading 1".."heading 6" (case-insensitive, so the
# Russian Word UI "Заголовок 1" — stored as "heading 1" — works too).
_DOCX_STYLE_MAP = """
p[style-name='Title'] => h1:fresh
p[style-name='Subtitle'] => h2:fresh
p[style-name='Quote'] => blockquote > p:fresh
p[style-name='Intense Quote'] => blockquote > p:fresh
comment-reference => sup
"""

_TABLE_TOKEN = "PDFMDTABLETOKEN{}X"

_IMAGE_EXT = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tif",
    "image/webp": "webp",
    "image/svg+xml": "svg",
    "image/x-emf": "emf",
    "image/x-wmf": "wmf",
}


_MATH_TOKEN = "PDFMDMATHTOKEN{}X"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_NARY = {"∑": r"\sum", "∏": r"\prod", "∫": r"\int", "∬": r"\iint", "∭": r"\iiint", "∮": r"\oint", "⋃": r"\bigcup", "⋂": r"\bigcap"}
_LATEX_ESCAPE = str.maketrans({"{": r"\{", "}": r"\}", "%": r"\%", "#": r"\#", "&": r"\&", "$": r"\$"})


def _omml_to_latex(el) -> str:
    """Word equation (OMML) -> LaTeX for the common constructs; unknown ones keep their text."""

    def tag(e) -> str:
        return e.tag.rsplit("}", 1)[-1]

    def child(e, name):
        return next((c for c in e if tag(c) == name), None)

    def sub(e, name) -> str:
        c = child(e, name)
        return conv(c) if c is not None else ""

    def prop(e, pr, name, default=""):
        p = child(e, pr)
        c = child(p, name) if p is not None else None
        return c.get(f"{{{_M_NS}}}val", default) if c is not None else default

    def conv(e) -> str:
        t = tag(e)
        if t == "t":
            return (e.text or "").translate(_LATEX_ESCAPE)
        if t.endswith("Pr") or t in ("ctrlPr",):
            return ""
        if t == "f":
            return rf"\frac{{{sub(e, 'num')}}}{{{sub(e, 'den')}}}"
        if t == "sSup":
            return f"{{{sub(e, 'e')}}}^{{{sub(e, 'sup')}}}"
        if t == "sSub":
            return f"{{{sub(e, 'e')}}}_{{{sub(e, 'sub')}}}"
        if t == "sSubSup":
            return f"{{{sub(e, 'e')}}}_{{{sub(e, 'sub')}}}^{{{sub(e, 'sup')}}}"
        if t == "sPre":
            return f"{{}}_{{{sub(e, 'sub')}}}^{{{sub(e, 'sup')}}}{{{sub(e, 'e')}}}"
        if t == "rad":
            deg = sub(e, "deg")
            return rf"\sqrt[{deg}]{{{sub(e, 'e')}}}" if deg else rf"\sqrt{{{sub(e, 'e')}}}"
        if t == "d":
            beg, end = prop(e, "dPr", "begChr", "("), prop(e, "dPr", "endChr", ")")
            sep = prop(e, "dPr", "sepChr", "|")
            inner = sep.join(conv(c) for c in e if tag(c) == "e")
            return rf"\left{beg or '.'}{inner}\right{end or '.'}".replace(r"\left{", r"\left\{").replace(r"\right}", r"\right\}")
        if t == "nary":
            op = prop(e, "naryPr", "chr", "∫")
            low, up = sub(e, "sub"), sub(e, "sup")
            out = _NARY.get(op, op)
            if low:
                out += f"_{{{low}}}"
            if up:
                out += f"^{{{up}}}"
            return out + f" {sub(e, 'e')}"
        if t == "func":
            return f"{sub(e, 'fName')}{{{sub(e, 'e')}}}"
        if t == "limLow":
            return f"{sub(e, 'e')}_{{{sub(e, 'lim')}}}"
        if t == "limUpp":
            return f"{sub(e, 'e')}^{{{sub(e, 'lim')}}}"
        if t == "bar":
            return rf"\overline{{{sub(e, 'e')}}}" if prop(e, "barPr", "pos", "bot") == "top" else rf"\underline{{{sub(e, 'e')}}}"
        if t == "acc":
            return rf"\overset{{{prop(e, 'accPr', 'chr', '̂')}}}{{{sub(e, 'e')}}}"
        if t == "m":
            rows = [" & ".join(conv(c) for c in mr if tag(c) == "e") for mr in e if tag(mr) == "mr"]
            return r"\begin{matrix}" + r" \\ ".join(rows) + r"\end{matrix}"
        if t == "eqArr":
            return r"\begin{aligned}" + r" \\ ".join(conv(c) for c in e if tag(c) == "e") + r"\end{aligned}"
        return "".join(conv(c) for c in e)

    return conv(el).strip()


def _extract_word_math(document_xml: str) -> tuple[str, list[tuple[str, bool]]]:
    """Replace Word equations with text tokens (mammoth drops OMML); returns (xml, [(latex, display)])."""
    from xml.etree.ElementTree import fromstring

    if "<m:oMath" not in document_xml:
        return document_xml, []
    root_tag = re.search(r"<w:document\b[^>]*>", document_xml)
    ns_decl = " ".join(re.findall(r'xmlns:\w+="[^"]*"', root_tag.group(0))) if root_tag else ""
    formulas: list[tuple[str, bool]] = []

    def replace(match: "re.Match[str]", display: bool) -> str:
        try:
            el = fromstring(f"<root {ns_decl}>{match.group(0)}</root>")[0]
            latex = _omml_to_latex(el)
        except Exception:
            latex = re.sub(r"<[^>]+>", "", match.group(0))  # keep at least the text
        if not latex:
            return ""
        formulas.append((latex, display))
        token = _MATH_TOKEN.format(len(formulas) - 1)
        return f'<w:r><w:t xml:space="preserve">{token}</w:t></w:r>'

    xml = re.sub(r"<m:oMathPara\b.*?</m:oMathPara>", lambda m: replace(m, True), document_xml, flags=re.S)
    xml = re.sub(r"<m:oMath\b.*?</m:oMath>", lambda m: replace(m, False), xml, flags=re.S)
    return xml, formulas


def _docx_with_math_tokens(path: Path) -> tuple[io.BytesIO | None, list[tuple[str, bool]]]:
    """A copy of the .docx (in memory) with equations replaced by tokens, if it has any."""
    with zipfile.ZipFile(path) as z:
        try:
            xml = z.read("word/document.xml").decode("utf-8")
        except KeyError:
            return None, []
        new_xml, formulas = _extract_word_math(xml)
        if not formulas:
            return None, []
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                data = new_xml.encode("utf-8") if item.filename == "word/document.xml" else z.read(item.filename)
                out.writestr(item, data)
    buf.seek(0)
    return buf, formulas


def _metafile_to_png(data: bytes) -> bytes | None:
    """EMF/WMF (Office vector pictures) -> PNG: Markdown viewers cannot show metafiles."""
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as img:
            img.load(dpi=144)
            out = io.BytesIO()
            img.convert("RGBA").save(out, "PNG")
            return out.getvalue()
    except Exception:
        return None


def _html_to_md(html_fragment: str) -> str:
    markdownify = _require("markdownify", "markdownify")
    return markdownify.markdownify(html_fragment, heading_style="ATX", bullets="-")


def _cell_text(cell) -> str:
    """Table cell HTML -> Markdown text; paragraphs become line breaks.

    A table nested in the cell becomes one line per row, cells joined by " · " (a table inside a
    Markdown table cell is not possible; this keeps every value readable and in order).
    """
    import html as html_lib

    from bs4 import BeautifulSoup

    own_table = cell.find_parent("table")
    for nested in [t for t in cell.find_all("table") if t.find_parent("table") is own_table]:
        rows = []
        for tr in nested.find_all("tr"):
            if tr.find_parent("table") is not nested:
                continue
            texts = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"], recursive=False)]
            if any(texts):
                rows.append(" · ".join(t for t in texts if t))
        fragment = "".join(f"<p>{html_lib.escape(r)}</p>" for r in rows)
        nested.replace_with(BeautifulSoup(fragment, "html.parser"))
    md = _html_to_md(cell.decode_contents()).strip()
    lines = [ln.strip() for ln in md.split("\n")]
    return "\n".join(ln for ln in lines if ln)


def html_table_grid(table) -> tuple[Grid, int]:
    """bs4 <table> -> (span-aware grid, header row count)."""
    rows = [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]
    occupied: dict[tuple[int, int], bool] = {}
    cells: dict[tuple[int, int], Cell] = {}
    ncols = 0
    for r, tr in enumerate(rows):
        c = 0
        for td in tr.find_all(["td", "th"], recursive=False):
            while occupied.get((r, c)):
                c += 1
            rowspan = max(1, int(td.get("rowspan", 1) or 1))
            colspan = max(1, int(td.get("colspan", 1) or 1))
            rowspan = min(rowspan, len(rows) - r)
            cells[(r, c)] = Cell(_cell_text(td), rowspan=rowspan, colspan=colspan)
            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied[(r + dr, c + dc)] = True
            c += colspan
            ncols = max(ncols, c)
    grid: Grid = []
    for r in range(len(rows)):
        row: list[Optional[Cell]] = []
        for c in range(ncols):
            if (r, c) in cells:
                row.append(cells[(r, c)])
            elif occupied.get((r, c)):
                row.append(None)
            else:
                row.append(Cell(""))
        grid.append(row)
    thead = table.find("thead")
    header_rows = len(thead.find_all("tr")) if thead is not None else 0
    if header_rows == 0 and rows and all(td.name == "th" for td in rows[0].find_all(["td", "th"], recursive=False)):
        header_rows = 1
    return grid, header_rows


def convert_docx(
    path: Path,
    out_md: Path,
    *,
    save_images: bool = True,
    progress: Optional[StepCallback] = None,
) -> str:
    from pdf2md import _md_image_ref, images_dir_for  # late import: pdf2md imports this module

    mammoth = _require("mammoth", "mammoth")
    bs4 = _require("bs4", "beautifulsoup4")

    images_dir = images_dir_for(out_md)
    counter = {"n": 0}

    def _image(image):
        counter["n"] += 1
        ext = _IMAGE_EXT.get((image.content_type or "").lower(), "bin")
        with image.open() as src:
            data = src.read()
        if ext in ("emf", "wmf"):
            png = _metafile_to_png(data)
            if png is not None:
                data, ext = png, "png"
        name = f"image{counter['n']}.{ext}"
        images_dir.mkdir(parents=True, exist_ok=True)
        (images_dir / name).write_bytes(data)
        return {"src": _md_image_ref(images_dir, name)}

    _step(progress, 5.0, tr('Word: чтение документа…'))
    try:
        prepared, formulas = _docx_with_math_tokens(path)
        with (prepared or path.open("rb")) as f:
            kwargs = {"style_map": _DOCX_STYLE_MAP}
            if save_images:
                kwargs["convert_image"] = mammoth.images.img_element(_image)
            result = mammoth.convert_to_html(f, **kwargs)
    except Exception as e:
        raise OfficeError(tr('Не удалось открыть Word-файл (повреждён или не .docx): {v0}').format(v0=e)) from e
    _step(progress, 50.0, tr('Word: разметка…'))

    soup = bs4.BeautifulSoup(result.value, "html.parser")
    if not save_images:
        for img in soup.find_all("img"):
            img.decompose()

    tables: list[str] = []
    # Top-level tables only, collected up front: nested ones are flattened into their outer
    # cell's text (and detached from the tree) while the outer table is processed.
    for table in [t for t in soup.find_all("table") if t.find_parent("table") is None]:
        grid, header_rows = html_table_grid(table)
        if not grid:
            table.decompose()
            continue
        if header_rows == 0:
            header_rows = 1  # Word tables rarely mark header rows; the first row is the header
        token = _TABLE_TOKEN.format(len(tables))
        tables.append(render_table(grid, header_rows=header_rows))
        placeholder = soup.new_tag("p")
        placeholder.string = token
        table.replace_with(placeholder)

    md = _html_to_md(str(soup))
    for i, table_md in enumerate(tables):
        md = md.replace(_TABLE_TOKEN.format(i), "\n" + table_md.rstrip("\n") + "\n", 1)
    for i, (latex, display) in enumerate(formulas):
        md = md.replace(_MATH_TOKEN.format(i), f"\n$$\n{latex}\n$$\n" if display else f"${latex}$", 1)
    md = _tidy(md)

    warnings = sorted({m.message for m in result.messages if m.type == "warning"})
    if warnings:
        shown = "; ".join(warnings[:3]) + (f" … (+{len(warnings) - 3})" if len(warnings) > 3 else "")
        _step(progress, 90.0, f"WARN: Word: {shown}")
    images_note = tr(', картинок {v0}').format(v0=counter['n']) if save_images else ""
    _step(progress, 100.0, tr('Word: готово (таблиц {v0}{v1})').format(v0=len(tables), v1=images_note))
    return md


# --- DOC (Word 97-2003) via Microsoft Word ------------------------------------------

# Runs in Windows PowerShell; paths come through environment variables (no quoting issues
# with spaces, "+", Cyrillic). Word opens a temporary COPY of the .doc read-only with macros
# disabled, so the original is never touched. Word's COM methods take arguments by reference
# in PowerShell, hence the [ref] everywhere.
_WORD_DOC_TO_DOCX_PS = r"""
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: no macros
    $src = $env:PDF2MD_DOC_SRC; $no = $false; $yes = $true
    # a dummy password makes a protected file fail fast instead of showing a dialog
    $pwd = 'pdf2md-no-password'
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument)
    $doc = $word.Documents.Open([ref]$src, [ref]$no, [ref]$yes, [ref]$no, [ref]$pwd)
    try {
        $dst = $env:PDF2MD_DOC_DST; $fmt = [int]$env:PDF2MD_DOC_FMT
        $doc.SaveAs2([ref]$dst, [ref]$fmt)
    } finally {
        $save = 0; $doc.Close([ref]$save)
    }
} finally {
    if ($word.Documents.Count -eq 0) { $save = 0; $word.Quit([ref]$save) }   # never close user's documents
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
"""


# Same idea for Excel (its COM methods take plain arguments in PowerShell).
_EXCEL_SAVE_AS_PS = r"""
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.AutomationSecurity = 3   # no macros
    # Open(FileName, UpdateLinks, ReadOnly, Format, Password)
    $wb = $xl.Workbooks.Open($env:PDF2MD_DOC_SRC, 0, $true, [Type]::Missing, 'pdf2md-no-password')
    try {
        $wb.SaveAs($env:PDF2MD_DOC_DST, [int]$env:PDF2MD_DOC_FMT)
    } finally {
        $wb.Close($false)
    }
} finally {
    if ($xl.Workbooks.Count -eq 0) { $xl.Quit() }   # never close the user's own workbooks
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"""

_OFFICE_APPS = {
    "word": ("Word.Application", "Microsoft Word", _WORD_DOC_TO_DOCX_PS, ".docx"),
    "excel": ("Excel.Application", "Microsoft Excel", _EXCEL_SAVE_AS_PS, ".xlsx"),
}


def _com_available(progid: str) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"))
        return True
    except OSError:
        return False


def word_available() -> bool:
    return _com_available("Word.Application")


def excel_available() -> bool:
    return _com_available("Excel.Application")


WD_FORMAT_DOC = 0  # wdFormatDocument (Word 97-2003)
WD_FORMAT_DOCX = 16  # wdFormatXMLDocument
XL_FORMAT_XLS = 56  # xlExcel8 (Excel 97-2003)
XL_FORMAT_XLSX = 51  # xlOpenXMLWorkbook


def word_save_as(src: Path, dst: Path, fmt: int, *, timeout: float = 300.0) -> None:
    """Re-save a Word file in another format with the installed Microsoft Word."""
    office_save_as("word", src, dst, fmt, timeout=timeout)


def office_save_as(app: str, src: Path, dst: Path, fmt: int, *, timeout: float = 300.0) -> None:
    """Re-save a file with the installed Microsoft Word / Excel (COM through PowerShell)."""
    progid, app_name, script, modern = _OFFICE_APPS[app]
    if not _com_available(progid):
        raise OfficeError(
            tr('Старый формат {v0}: нужен установленный {v1}. Либо откройте файл в {v2} и сохраните как {v3}.').format(v0=src.suffix.lower(), v1=app_name, v2=app_name.split()[-1], v3=modern)
        )
    with tempfile.TemporaryDirectory(prefix="pdf2md_office_") as tmp:
        copy = Path(tmp) / f"source{src.suffix.lower()}"
        shutil.copyfile(src, copy)
        env = dict(
            os.environ,
            PDF2MD_DOC_SRC=str(copy),
            PDF2MD_DOC_DST=str(dst.resolve()),
            PDF2MD_DOC_FMT=str(fmt),
        )
        try:
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-Command", script],
                env=env,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as e:
            raise OfficeError(tr('{v0} не ответил за {v1} с').format(v0=app_name, v1=int(timeout))) from e
    if proc.returncode != 0 or not dst.is_file():
        detail = (proc.stderr or proc.stdout).strip().splitlines()
        reason = next((ln for ln in detail if ln.strip()), tr('неизвестная ошибка'))
        raise OfficeError(tr('{v0} не смог открыть файл (пароль или повреждён?): {v1}').format(v0=app_name, v1=reason[:300]))


def xls_as_xlsx(src: Path) -> Path:
    """Old .xls -> cached .xlsx copy made by Excel (the sheet list and the conversion reuse it)."""
    import hashlib

    st = src.stat()
    key = hashlib.sha1(f"{src.resolve()}|{st.st_mtime_ns}|{st.st_size}".encode("utf-8")).hexdigest()[:16]
    cache = Path(tempfile.gettempdir()) / "pdf2md_xls_cache"
    dst = cache / f"{key}.xlsx"
    if not dst.is_file():
        cache.mkdir(parents=True, exist_ok=True)
        office_save_as("excel", src, dst, XL_FORMAT_XLSX)
    return dst


def doc_to_docx(src: Path, dst: Path, *, timeout: float = 300.0) -> None:
    """Old binary .doc -> .docx with the installed Microsoft Word (COM through PowerShell)."""
    word_save_as(src, dst, WD_FORMAT_DOCX, timeout=timeout)


def convert_doc(
    path: Path,
    out_md: Path,
    *,
    save_images: bool = True,
    progress: Optional[StepCallback] = None,
) -> str:
    _step(progress, 2.0, tr('Word: старый .doc → .docx через Microsoft Word…'))
    with tempfile.TemporaryDirectory(prefix="pdf2md_doc_") as tmp:
        docx = Path(tmp) / "converted.docx"
        doc_to_docx(path, docx)
        return convert_docx(docx, out_md, save_images=save_images, progress=progress)


# --- Entry point -----------------------------------------------------------------


def convert_office(
    path: Path,
    out_md: Path,
    *,
    sheet: Optional[str] = None,
    save_images: bool = True,
    progress: Optional[StepCallback] = None,
) -> str:
    """Convert a non-PDF file; returns the Markdown text (the caller writes it)."""
    kind = office_kind(path)
    if kind == "txt":
        return convert_txt(path, progress=progress)
    if kind == "csv":
        return convert_csv(path, progress=progress)
    if kind == "xlsx":
        return convert_xlsx(path, out_md, sheet=sheet, save_images=save_images, progress=progress)
    if kind == "xls":
        _step(progress, 2.0, tr('Excel: старый .xls → .xlsx через Microsoft Excel…'))
        return convert_xlsx(xls_as_xlsx(path), out_md, sheet=sheet, save_images=save_images, progress=progress)
    if kind == "docx":
        return convert_docx(path, out_md, save_images=save_images, progress=progress)
    if kind == "doc":
        return convert_doc(path, out_md, save_images=save_images, progress=progress)
    raise OfficeError(tr('Неподдерживаемый тип файла: {v0}').format(v0=path.suffix))
