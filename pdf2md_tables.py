"""Table helpers shared by the converters.

* `render_table()` turns a grid of cells into Markdown:
  - a plain GFM pipe table when the grid is regular;
  - an HTML <table> with rowspan/colspan when cells are merged — exact
    structure, nothing duplicated (GFM viewers: GitHub, Obsidian, VS Code,
    Typora render HTML tables inside .md).
* `is_fake_pdf_table()` rejects "tables" PyMuPDF finds in vector drawings
  (keyboards, block diagrams): mostly empty grids, single rows, many curves.
* `render_pymupdf_table()` adapts a PyMuPDF `Table` to `render_table()`.

Why not PyMuPDF's own Table.to_markdown(): with fill_empty=True (the default)
it fills merged cells by copying the cell to the LEFT first, which silently
puts e.g. the "LSB" value into the "Value" column — wrong data.
"""

from __future__ import annotations

import html
from dataclasses import dataclass
from typing import Optional, Sequence


@dataclass
class Cell:
    text: str
    rowspan: int = 1
    colspan: int = 1


# grid[row][col] is a Cell at the top-left corner of a (possibly merged) cell,
# or None for positions covered by a merged cell.
Grid = list[list[Optional[Cell]]]


def _norm(text: Optional[str]) -> str:
    if text is None:
        return ""
    lines = [ln.strip() for ln in str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    return "\n".join(ln for ln in lines if ln)


def _has_spans(grid: Grid) -> bool:
    return any(c is None or c.rowspan > 1 or c.colspan > 1 for row in grid for c in row)


def _pipe_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", "<br>")


def _html_cell(text: str) -> str:
    return "<br>".join(html.escape(ln) for ln in text.split("\n"))


def render_table(grid: Grid, *, header_rows: int = 1) -> str:
    """Markdown for a table grid. Pipe table if regular, HTML if cells are merged."""
    grid = [row for row in grid if row]
    if not grid:
        return ""
    ncols = max(len(r) for r in grid)
    grid = [list(r) + [Cell("")] * (ncols - len(r)) for r in grid]

    if not _has_spans(grid):
        rows = [[_pipe_cell(c.text) for c in r] for r in grid]  # type: ignore[union-attr]
        head = rows[0] if header_rows else [""] * ncols
        body = rows[1:] if header_rows else rows
        out = ["| " + " | ".join(head) + " |", "|" + "|".join(["---"] * ncols) + "|"]
        out += ["| " + " | ".join(r) + " |" for r in body]
        return "\n".join(out) + "\n"

    out = ["<table>"]
    for j, row in enumerate(grid):
        tag = "th" if j < header_rows else "td"
        cells = []
        for c in row:
            if c is None:
                continue
            attrs = ""
            if c.rowspan > 1:
                attrs += f' rowspan="{c.rowspan}"'
            if c.colspan > 1:
                attrs += f' colspan="{c.colspan}"'
            cells.append(f"<{tag}{attrs}>{_html_cell(c.text)}</{tag}>")
        out.append("<tr>" + "".join(cells) + "</tr>")
    out.append("</table>")
    return "\n".join(out) + "\n"


# --- PyMuPDF tables ---------------------------------------------------------

_TOL = 1.5  # points


def pymupdf_table_grid(table) -> Grid:
    """Build a span-aware grid from a PyMuPDF Table.

    `table.extract()` returns None for positions covered by a merged cell;
    spans are recovered from cell bounding boxes against column/row edges.
    """
    data = table.extract()
    rows = table.rows
    nrows = len(data)
    ncols = max((len(r) for r in data), default=0)

    col_x0: list[float] = []
    for i in range(ncols):
        xs = [r.cells[i][0] for r in rows if i < len(r.cells) and r.cells[i] is not None]
        col_x0.append(min(xs) if xs else float("inf"))
    row_y0 = [r.bbox[1] for r in rows]

    grid: Grid = []
    for j in range(nrows):
        out_row: list[Optional[Cell]] = []
        for i in range(ncols):
            text = data[j][i] if i < len(data[j]) else None
            box = rows[j].cells[i] if i < len(rows[j].cells) else None
            if text is None and box is None:
                out_row.append(None)
                continue
            colspan = rowspan = 1
            if box is not None:
                x1, y1 = box[2], box[3]
                colspan = 1 + sum(1 for k in range(i + 1, ncols) if col_x0[k] < x1 - _TOL)
                rowspan = 1 + sum(1 for r in range(j + 1, nrows) if row_y0[r] < y1 - _TOL)
            out_row.append(Cell(_norm(text), rowspan=rowspan, colspan=colspan))
        grid.append(out_row)

    # Positions not covered by any span but still None (rare detection gaps)
    # become empty cells so the HTML table stays rectangular.
    covered = [[False] * ncols for _ in range(nrows)]
    for j, row in enumerate(grid):
        for i, c in enumerate(row):
            if c is None:
                continue
            for dj in range(c.rowspan):
                for di in range(c.colspan):
                    if j + dj < nrows and i + di < ncols:
                        covered[j + dj][i + di] = True
    for j in range(nrows):
        for i in range(ncols):
            if not covered[j][i]:
                grid[j][i] = Cell("")
    return grid


def render_pymupdf_table(table) -> str:
    grid = pymupdf_table_grid(table)
    header_rows = 1
    try:
        header = table.header
        if header is not None and header.external and any(header.names):
            # Header text sits above the table body, outside its cells.
            ncols = len(grid[0]) if grid else len(header.names)
            names = [_norm(n) for n in header.names][:ncols]
            grid.insert(0, [Cell(n) for n in names] + [Cell("")] * (ncols - len(names)))
    except Exception:
        pass
    return render_table(grid, header_rows=header_rows)


def table_stats(table) -> tuple[int, int, float]:
    """(rows, cols, share of non-empty cells)."""
    data = table.extract()
    cells = [c for r in data for c in r]
    filled = sum(1 for c in cells if c is not None and str(c).strip())
    return len(data), max((len(r) for r in data), default=0), filled / max(len(cells), 1)


def is_fake_pdf_table(page, table, *, curve_limit: int = 50) -> bool:
    """Heuristic: is this "table" really a vector drawing?

    Calibrated on real manuals (Korg microKORG XL, Moog Mavis): genuine tables
    had 69–100 % filled cells and no curves; drawings detected as tables had
    0–25 % filled cells and hundreds of curves (keyboards, knobs).
    """
    rows, cols, filled = table_stats(table)
    if rows < 2 or cols < 2:
        return True
    if filled < 0.4:
        return True
    try:
        bbox = table.bbox
        curves = 0
        for d in page.get_drawings():
            r = d.get("rect")
            if r is None:
                continue
            if r.intersects(bbox) and any(item[0] == "c" for item in d.get("items", ())):
                curves += 1
                if curves >= curve_limit:
                    return True
    except Exception:
        pass
    return False


def drop_fake_tables(page, finder) -> None:
    """Remove drawing-like tables from a PyMuPDF TableFinder in place."""
    try:
        finder.tables = [t for t in finder.tables if not is_fake_pdf_table(page, t)]
    except Exception:
        pass


def is_regular(grid: Sequence[Sequence[Optional[Cell]]]) -> bool:
    return not _has_spans([list(r) for r in grid])
