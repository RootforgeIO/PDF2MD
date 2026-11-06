"""Tests for Word / Excel / CSV / TXT -> Markdown (office2md + pdf2md.convert_file).

DOCX fixtures are written by hand (zipfile + WordprocessingML), so no python-docx
is needed. XLSX fixtures use openpyxl, which the converter needs anyway.
Tests for a format are skipped when its library (mammoth / openpyxl) is missing.
"""

from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import office2md  # noqa: E402
import pdf2md  # noqa: E402

HAS_MAMMOTH = importlib.util.find_spec("mammoth") is not None
HAS_OPENPYXL = importlib.util.find_spec("openpyxl") is not None

# 1x1 red PNG
PNG_1X1 = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753de"
    "0000000c4944415408d763f8cfc000000301010018dd8db00000000049454e44ae426082"
)

_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_NS = (
    _W + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)


def _p(text: str, style: str = "", num: bool = False, bold: bool = False) -> str:
    ppr = ""
    if style or num:
        ppr = "<w:pPr>"
        if style:
            ppr += f'<w:pStyle w:val="{style}"/>'
        if num:
            ppr += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        ppr += "</w:pPr>"
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def _tc(text: str, extra: str = "") -> str:
    return f"<w:tc><w:tcPr>{extra}</w:tcPr>{_p(text)}</w:tc>"


def make_docx(path: Path) -> Path:
    image = (
        "<w:p><w:r><w:drawing><wp:inline>"
        '<wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Pic" descr="Красная точка"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Pic\"/><pic:cNvPicPr/></pic:nvPicPr>"
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    # 3 columns; row 1: "Группа" spans 2 columns; column 3 "Значение" merged vertically over rows 1-2.
    table = (
        "<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>"
        "<w:tr>"
        + _tc("Группа", '<w:gridSpan w:val="2"/>')
        + _tc("Значение", '<w:vMerge w:val="restart"/>')
        + "</w:tr><w:tr>"
        + _tc("A")
        + _tc("B")
        + _tc("", "<w:vMerge/>")
        + "</w:tr><w:tr>"
        + _tc("1")
        + _tc("2")
        + _tc("3")
        + "</w:tr></w:tbl>"
    )
    simple = (
        "<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
        "<w:tr>" + _tc("Параметр") + _tc("Диапазон") + "</w:tr>"
        "<w:tr>" + _tc("Cutoff") + _tc("0…127") + "</w:tr></w:tbl>"
    )
    body = (
        _p("Руководство", style="Heading1")
        + _p("Обычный абзац с ")
        + _p("жирным текстом", bold=True)
        + _p("Раздел", style="Heading2")
        + _p("первый пункт", num=True)
        + _p("второй пункт", num=True)
        + image
        + table
        + _p("Диапазоны:")  # Word merges back-to-back tables when saving .doc
        + simple
        + _p("Конец.")
    )
    document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {_NS}><w:body>{body}</w:body></w:document>'
    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {_W}>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>'
        "</w:styles>"
    )
    numbering = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {_W}>'
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/media/image1.png", PNG_1X1)
    return path


def make_xlsx(path: Path) -> Path:
    import datetime as dt

    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XlImage

    wb = Workbook()
    first = wb.active
    first.title = "Сводка"
    first["A1"] = "не тот лист"

    ws = wb.create_sheet("Данные")
    # empty first row and column — must be trimmed
    ws["B2"], ws["C2"], ws["D2"] = "Товар", "Цена", "Доля"
    ws["B3"], ws["C3"], ws["D3"] = "Синтезатор", 1234.5, 0.125
    ws["C3"].number_format = "0.00"
    ws["D3"].number_format = "0.0%"
    ws["B4"], ws["C4"] = "Дата", dt.datetime(2026, 11, 2)
    ws["C4"].number_format = "dd.mm.yyyy"
    ws["B5"] = "Итого"
    ws["C5"] = "=C3*2"  # no cached value when written by openpyxl
    ws.merge_cells("C4:D4")
    png = path.with_name("pixel.png")
    png.write_bytes(PNG_1X1)
    ws.add_image(XlImage(str(png)), "F2")
    wb.active = wb.sheetnames.index("Данные")
    wb.save(path)
    return path


class TextFormatsTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_decode_text_encodings(self) -> None:
        text = "Привет, мир — ёлка"
        for enc in ("utf-8", "utf-8-sig", "utf-16", "cp1251"):
            decoded, _used = office2md.decode_text(text.encode(enc))
            self.assertEqual(decoded, text, enc)
        for enc in ("utf-16-le", "utf-16-be"):  # no BOM
            self.assertEqual(office2md.decode_text(text.encode(enc)), (text, enc))
            ascii_text = "plain ascii line\n"
            self.assertEqual(office2md.decode_text(ascii_text.encode(enc)), (ascii_text, enc))

    def test_txt_utf16_with_cyrillic_name(self) -> None:
        src = self.tmp / "заметки.txt"
        src.write_bytes("Строка 1\r\nСтрока 2\r\n\r\n# не заголовок?\r\n".encode("utf-16"))
        out = src.with_suffix(".md")
        ok, used = pdf2md.convert_file(src, out)
        self.assertEqual((ok, used), (True, "txt"))
        self.assertEqual(out.read_text(encoding="utf-8"), "Строка 1\nСтрока 2\n\n# не заголовок?\n")

    def test_csv_cp1251_semicolon(self) -> None:
        src = self.tmp / "прайс.csv"
        src.write_bytes("Имя;Цена;Примечание\nКабель;100;a|b\n\nСтойка;2500;\n".encode("cp1251"))
        out = src.with_suffix(".md")
        pdf2md.convert_file(src, out)
        md = out.read_text(encoding="utf-8")
        self.assertIn("| Имя | Цена | Примечание |", md)
        self.assertIn("| Кабель | 100 | a\\|b |", md)
        self.assertIn("| Стойка | 2500 |  |", md)
        self.assertEqual(md.count("\n"), 4)  # header, rule, 2 rows — empty line dropped

    def test_tsv_uses_tab(self) -> None:
        src = self.tmp / "t.tsv"
        src.write_text("a\tb,c\n1\t2,3\n", encoding="utf-8")
        md = office2md.convert_office(src, src.with_suffix(".md"))
        self.assertIn("| a | b,c |", md)

    def test_cli_rejects_unsupported_and_converts_txt(self) -> None:
        bad = self.tmp / "x.pptx"
        bad.write_text("x")
        self.assertEqual(pdf2md.main([str(bad)]), 2)
        src = self.tmp / "a.txt"
        src.write_text("hello\n", encoding="utf-8")
        self.assertEqual(pdf2md.main([str(src), "-o", str(self.tmp / "o" / "a.md")]), 0)
        self.assertEqual((self.tmp / "o" / "a.md").read_text(encoding="utf-8"), "hello\n")


@unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
class XlsxTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.src = make_xlsx(self.tmp / "книга продаж.xlsx")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_list_sheets_and_resolve(self) -> None:
        names, active = office2md.list_sheets(self.src)
        self.assertEqual(names, ["Сводка", "Данные"])
        self.assertEqual(active, 1)
        self.assertEqual(office2md.resolve_sheet(names, active, None), "Данные")
        self.assertEqual(office2md.resolve_sheet(names, active, "1"), "Сводка")
        self.assertEqual(office2md.resolve_sheet(names, active, "сводка"), "Сводка")
        with self.assertRaises(office2md.OfficeError):
            office2md.resolve_sheet(names, active, "Нет такого")

    def test_active_sheet_values_merges_images_and_formula_warning(self) -> None:
        out = self.src.with_suffix(".md")
        messages: list[str] = []
        ok, used = pdf2md.convert_file(self.src, out, progress_callback=lambda u: messages.append(u.message))
        self.assertEqual((ok, used), (True, "xlsx"))
        md = out.read_text(encoding="utf-8")
        self.assertTrue(md.startswith("## Данные\n"))
        self.assertNotIn("не тот лист", md)
        # merged C4:D4 -> HTML table with colspan; formats as displayed
        self.assertIn("<table>", md)
        self.assertIn('<td colspan="2">02.11.2026</td>', md)
        self.assertIn("<td>1234,50</td>", md)  # "0.00": no grouping
        self.assertIn("<td>12,5%</td>", md)
        self.assertIn("<th>Товар</th>", md)
        # trimmed: no empty leading column/row
        self.assertNotIn("<td></td><td>Синтезатор", md)
        # image saved next to the .md, relative link with %20
        images = office2md_images(out)
        self.assertEqual(len(images), 1)
        self.assertIn("книга%20продаж_images/sheet2_image1.png", md)
        self.assertIn("(ячейка F2)", md)
        self.assertTrue(any(m.startswith("WARN:") and "формул" in m for m in messages))

    def test_sheet_by_name_via_cli(self) -> None:
        out = self.tmp / "s.md"
        self.assertEqual(pdf2md.main([str(self.src), "--sheet", "Сводка", "-o", str(out)]), 0)
        md = out.read_text(encoding="utf-8")
        self.assertIn("## Сводка", md)
        self.assertIn("| не тот лист |", md)
        self.assertEqual(pdf2md.main([str(self.src), "--sheet", "9", "-o", str(out)]), 4)

    def test_format_cell_value(self) -> None:
        import datetime as dt

        f = office2md.format_cell_value
        self.assertEqual(f(3.0), "3")
        self.assertEqual(f(0.1), "0,1")
        self.assertEqual(f(1234567), "1234567")  # General: no grouping, like Excel
        self.assertEqual(f(2.5, "0.000"), "2,500")
        self.assertEqual(f(0.5, "0%"), "50%")
        self.assertEqual(f(0.125, "0.0%"), "12,5%")
        self.assertEqual(f(True), "TRUE")
        self.assertEqual(f(dt.datetime(2026, 1, 5, 14, 30), "dd.mm.yyyy hh:mm"), "05.01.2026 14:30")
        self.assertEqual(f(dt.datetime(2026, 1, 5), "mm-dd-yy"), "2026-01-05")
        self.assertEqual(f(dt.time(9, 5), "h:mm"), "09:05")
        self.assertEqual(f(0.1 + 0.2), "0,3")  # 15 significant digits, like Excel
        # digit grouping, decimal comma, Excel rounding
        self.assertEqual(f(1234567.891, "#,##0.00"), "1 234 567,89")
        self.assertEqual(f(-1234567, "#,##0"), "-1 234 567")
        self.assertEqual(f(999, "#,##0.00"), "999,00")
        self.assertEqual(f(2.675, "0.00"), "2,68")  # half up, not banker's / binary 2.67
        self.assertEqual(f(-0.001, "0.00"), "0,00")  # no "-0,00"
        self.assertEqual(f(1234567, "#,##0.00;[Red]-#,##0.00"), "1 234 567,00")
        self.assertEqual(f(1500000, "#,##0,"), "1 500")  # trailing comma = thousands
        self.assertEqual(f(42, "@"), "42")

    def test_wide_table_keeps_far_right_cells_in_their_rows(self) -> None:
        from openpyxl import Workbook
        from openpyxl.comments import Comment

        wb = Workbook()
        ws = wb.active
        ws.title = "Лист"
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(r, c, f"r{r}c{c}")
        ws.merge_cells("A4:C4")
        ws["A4"] = "итог"
        ws["ZE2"] = "далеко"  # ~680th column, same row as r2c*
        ws["ZF3"] = 1234.5
        ws["ZF3"].number_format = "#,##0.00"
        ws["B20"] = "внизу"  # separated by > BLOCK_GAP empty rows -> own table
        ws["B2"].comment = Comment("Проверить формулу", "Автор")
        ws["ZH1"].comment = Comment("Заметка на пустой ячейке", "Автор")  # extends the table
        ws["A4"].comment = Comment("На объединённой ячейке", "Автор")  # merged A4:C4
        src = self.tmp / "wide.xlsx"
        wb.save(src)

        messages: list[str] = []
        md = office2md.convert_office(src, src.with_suffix(".md"), progress=lambda p, m: messages.append(m))
        # D..ZD (>= BLOCK_GAP empty columns) left out; the single empty ZG stays
        self.assertIn("*Ячейки A1:ZH4; пустые столбцы D–ZD не показаны;", md)
        self.assertIn("<tr><th>A</th><th>B</th><th>C</th><th>ZE</th><th>ZF</th><th>ZG</th><th>ZH</th></tr>", md)
        # far-right values stay in their own rows
        self.assertIn(
            "<tr><td>r2c1</td><td>r2c2 [прим. 2]</td><td>r2c3</td><td>далеко</td><td></td><td></td><td></td></tr>", md
        )
        self.assertIn(
            "<tr><td>r3c1</td><td>r3c2</td><td>r3c3</td><td></td><td>1 234,50</td><td></td><td></td></tr>", md
        )
        self.assertIn("<td>[прим. 1]</td></tr>", md)  # note on an empty far cell, row 1
        self.assertIn('<td colspan="3">итог [прим. 3]</td>', md)
        # notes listed under the table in reading order, with addresses
        self.assertIn("1. **ZH1**: Заметка на пустой ячейке", md)
        self.assertIn("2. **B2**: Проверить формулу", md)
        self.assertIn("3. **A4**: На объединённой ячейке", md)
        # rows separated by empty rows -> second table
        self.assertIn("*Ячейки B20:B20*", md)
        self.assertLess(md.index("ZH1"), md.index("B20:B20"))
        self.assertIn("| внизу |", md)
        self.assertEqual(md.count("<table>"), 1)
        self.assertLess(md.count("<td"), 40)  # no 680-column table
        self.assertTrue(any("широкая таблица" in m for m in messages))

    def test_sheet_blocks_and_gaps(self) -> None:
        B = office2md.SheetBlock
        values = {(1, 1): "a", (1, 3): "b", (1, 20): "c"}
        blocks = office2md.sheet_blocks(values, [])
        self.assertEqual(blocks, [B(1, 1, [1, 2, 3, 20])])
        self.assertEqual(blocks[0].skipped(), [(4, 19)])
        self.assertEqual(blocks[0].range_name(), "A1:T1")
        # a filled merge bridges the gap: nothing skipped
        self.assertEqual(office2md.sheet_blocks(values, [(1, 3, 1, 20)]), [B(1, 1, list(range(1, 21)))])
        # >= BLOCK_GAP empty rows split tables; fewer do not
        self.assertEqual(
            office2md.sheet_blocks({(1, 1): "a", (5, 1): "b", (11, 2): "c"}, []),
            [B(1, 5, [1]), B(11, 11, [2])],
        )
        # notes count as data
        self.assertEqual(office2md.sheet_blocks({(1, 1): "a"}, [], extra={(1, 2)}), [B(1, 1, [1, 2])])

    def test_empty_merge_across_skipped_columns(self) -> None:
        values = {(1, 1): "a", (2, 1): "x", (2, 30): "y"}
        merges = [(1, 2, 1, 40)]  # empty merged cell spanning the skipped run
        block = office2md.sheet_blocks(values, merges)[0]
        self.assertEqual(block.cols, [1, 30])
        grid = office2md.block_grid(values, merges, block)
        self.assertEqual([[c.text if c else None for c in row] for row in grid], [["a", ""], ["x", "y"]])


def office2md_images(out_md: Path) -> list[Path]:
    d = pdf2md.images_dir_for(out_md)
    return sorted(d.iterdir()) if d.is_dir() else []


@unittest.skipUnless(HAS_MAMMOTH, "mammoth not installed")
class DocxTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.src = make_docx(self.tmp / "Мануал v2.docx")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_headings_lists_tables_images(self) -> None:
        out = self.src.with_suffix(".md")
        ok, used = pdf2md.convert_file(self.src, out)
        self.assertEqual((ok, used), (True, "docx"))
        md = out.read_text(encoding="utf-8")
        self.assertIn("# Руководство", md)
        self.assertIn("## Раздел", md)
        self.assertIn("**жирным текстом**", md)
        self.assertIn("- первый пункт", md)
        self.assertIn("- второй пункт", md)
        # merged cells -> HTML with spans, nothing duplicated
        self.assertIn('<th colspan="2">Группа</th><th rowspan="2">Значение</th>', md)
        self.assertIn("<tr><td>A</td><td>B</td></tr>", md)
        self.assertEqual(md.count("Значение"), 1)
        # regular table -> pipe table
        self.assertIn("| Параметр | Диапазон |", md)
        self.assertIn("| Cutoff | 0…127 |", md)
        # image
        self.assertIn("![Красная точка](Мануал%20v2_images/image1.png)", md)
        self.assertEqual([p.name for p in office2md_images(out)], ["image1.png"])
        # order preserved
        self.assertLess(md.index("Руководство"), md.index("Группа"))
        self.assertLess(md.index("Cutoff"), md.index("Конец."))

    def test_no_images(self) -> None:
        out = self.src.with_suffix(".md")
        pdf2md.convert_file(self.src, out, save_images=False)
        md = out.read_text(encoding="utf-8")
        self.assertNotIn("![", md)
        self.assertFalse(pdf2md.images_dir_for(out).exists())

    def test_broken_docx_is_readable_error(self) -> None:
        bad = self.tmp / "bad.docx"
        bad.write_text("not a zip")
        with self.assertRaises(office2md.OfficeError):
            pdf2md.convert_file(bad, bad.with_suffix(".md"))
        self.assertEqual(pdf2md.main([str(bad)]), 4)


@unittest.skipUnless(HAS_MAMMOTH and office2md.word_available(), "needs mammoth and Microsoft Word")
class DocViaWordTests(unittest.TestCase):
    def test_old_doc_converted_through_word_source_untouched(self) -> None:
        import hashlib

        with tempfile.TemporaryDirectory() as tmp:
            docx = make_docx(Path(tmp) / "fixture.docx")
            src = Path(tmp) / "Старый документ.doc"
            office2md.word_save_as(docx, src, office2md.WD_FORMAT_DOC)
            sha = hashlib.sha256(src.read_bytes()).hexdigest()

            out = src.with_suffix(".md")
            ok, used = pdf2md.convert_file(src, out)
            self.assertEqual((ok, used), (True, "doc"))
            self.assertEqual(hashlib.sha256(src.read_bytes()).hexdigest(), sha)
            md = out.read_text(encoding="utf-8")
            self.assertIn("# Руководство", md)
            self.assertIn("| Cutoff | 0…127 |", md)
            self.assertIn('colspan="2"', md)
            self.assertEqual(len(office2md_images(out)), 1)
            self.assertEqual(sorted(p.name for p in Path(tmp).iterdir() if p.suffix == ".docx"), ["fixture.docx"])


@unittest.skipUnless(HAS_OPENPYXL and office2md.excel_available(), "needs openpyxl and Microsoft Excel")
class XlsViaExcelTests(unittest.TestCase):
    def test_old_xls_sheets_and_values(self) -> None:
        import hashlib

        with tempfile.TemporaryDirectory() as tmp:
            xlsx = make_xlsx(Path(tmp) / "fixture.xlsx")
            src = Path(tmp) / "Старая книга.xls"
            office2md.office_save_as("excel", xlsx, src, office2md.XL_FORMAT_XLS)
            sha = hashlib.sha256(src.read_bytes()).hexdigest()

            self.assertEqual(office2md.list_sheets(src), (["Сводка", "Данные"], 1))
            out = src.with_suffix(".md")
            ok, used = pdf2md.convert_file(src, out, sheet="Данные")
            self.assertEqual((ok, used), (True, "xls"))
            self.assertEqual(hashlib.sha256(src.read_bytes()).hexdigest(), sha)
            md = out.read_text(encoding="utf-8")
            self.assertIn("## Данные", md)
            self.assertIn("1234,50", md)
            self.assertIn("12,5%", md)
            self.assertIn("02.11.2026", md)


class EdgeCaseRegressionTests(unittest.TestCase):
    """Found by the full-program test pass (2026-11-03)."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_number_formats_keep_units_and_values(self) -> None:
        f = office2md.format_number
        self.assertEqual(f(12345.5, '#,##0.00" руб."'), "12 345,50 руб.")
        self.assertEqual(f(12345.5, "#,##0.00 [$₽-419]"), "12 345,50 ₽")
        self.assertEqual(f(1234.5, '_-* #,##0.00\\ "₽"_-;\\-* #,##0.00\\ "₽"_-;_-* "-"??\\ "₽"_-;_-@_-'), "1 234,50 ₽")
        self.assertEqual(f(-1234.5, '_-* #,##0.00\\ "₽"_-;\\-* #,##0.00\\ "₽"_-;_-* "-"??\\ "₽"_-;_-@_-'), "-1 234,50 ₽")
        self.assertEqual(f(0, '_-* #,##0.00\\ "₽"_-;\\-* #,##0.00\\ "₽"_-;_-* "-"??\\ "₽"_-;_-@_-'), "- ₽")
        self.assertEqual(f(0.0000123, "0.00E+00"), "1,23E-05")
        self.assertEqual(f(0.5, "# ?/?"), "1/2")
        self.assertEqual(f(1.25, "# ?/?"), "1 1/4")
        self.assertEqual(f(-1234, "#,##0;(#,##0)"), "(1 234)")
        self.assertEqual(f(1.25, "0.0#"), "1,25")
        self.assertEqual(f(1.5, "0.0#"), "1,5")
        self.assertEqual(f(0.25, '0%" скидка"'), "25% скидка")
        self.assertEqual(f(7, 'General" шт."'), "7 шт.")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_chart_sheets_are_skipped(self) -> None:
        from openpyxl import Workbook
        from openpyxl.chart import BarChart, Reference

        wb = Workbook()
        ws = wb.active
        ws.title = "Данные"
        for i in range(1, 4):
            ws.append([i, i * i])
        cs = wb.create_chartsheet("Диаграмма")
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=3))
        cs.add_chart(chart)
        wb.active = 1  # the chart sheet is active
        src = self.tmp / "chart.xlsx"
        wb.save(src)
        self.assertEqual(office2md.list_sheets(src), (["Данные"], 0))
        self.assertIn("| 3 | 9 |", office2md.convert_office(src, src.with_suffix(".md")))

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_hidden_rows_and_columns_are_marked(self) -> None:
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        for r in range(1, 6):
            ws.append([f"r{r}c{c}" for c in range(1, 5)])
        ws.row_dimensions[3].hidden = True
        ws.column_dimensions["B"].hidden = True
        src = self.tmp / "hidden.xlsx"
        wb.save(src)
        md = office2md.convert_office(src, src.with_suffix(".md"))
        self.assertIn("*в Excel скрыты (здесь показаны): строки 3, столбцы B*", md)
        self.assertIn("r3c2", md)

    def _docx(self, name: str, body: str) -> Path:
        path = make_docx(self.tmp / name)
        ns = _NS + ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
        with zipfile.ZipFile(path) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        parts["word/document.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ns}><w:body>{body}</w:body></w:document>'
        ).encode()
        with zipfile.ZipFile(path, "w") as z:
            for k, v in parts.items():
                z.writestr(k, v)
        return path

    @unittest.skipUnless(HAS_MAMMOTH, "mammoth not installed")
    def test_word_equations_become_latex(self) -> None:
        body = (
            "<m:oMathPara><m:oMath><m:r><m:t>E=m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e>"
            "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>"
            '<w:p><w:r><w:t xml:space="preserve">Доля </w:t></w:r><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>'
            "<m:den><m:r><m:t>b_1</m:t></m:r></m:den></m:f></m:oMath></w:p>"
        )
        src = self._docx("math.docx", body)
        md = office2md.convert_office(src, src.with_suffix(".md"))
        self.assertIn("$$\nE=m{c}^{2}\n$$", md)
        self.assertIn(r"Доля $\frac{a}{b_1}$", md)  # "_" not escaped inside math

    @unittest.skipUnless(HAS_MAMMOTH, "mammoth not installed")
    def test_nested_word_table_is_readable(self) -> None:
        inner = "<w:tbl><w:tr>" + _tc("Вложенная1") + _tc("Вложенная2") + "</w:tr></w:tbl>"
        body = (
            "<w:tbl><w:tr>" + _tc("Внешняя") + "<w:tc>" + inner + _p("") + "</w:tc></w:tr>"
            "<w:tr>" + _tc("Ещё") + _tc("Строка") + "</w:tr></w:tbl>"
        )
        src = self._docx("nested.docx", body)
        md = office2md.convert_office(src, src.with_suffix(".md"))
        self.assertIn("| Внешняя | Вложенная1 · Вложенная2 |", md)
        self.assertIn("| Ещё | Строка |", md)

    def test_cli_write_errors_are_readable(self) -> None:
        from test_pdf2md import make_text_pdf

        src = make_text_pdf(self.tmp / "a.pdf", pages=1)
        out = self.tmp / "a.md"
        out.write_text("old", encoding="utf-8")
        os.chmod(out, 0o444)
        try:
            self.assertEqual(pdf2md.main([str(src), "--engine", "pymupdf"]), 5)
        finally:
            os.chmod(out, 0o666)
        folder = self.tmp / "outdir"
        folder.mkdir()
        self.assertEqual(pdf2md.main([str(src), "--engine", "pymupdf", "-o", str(folder)]), 0)
        self.assertTrue((folder / "a.md").is_file())

    def test_cancel_stops_marker_process(self) -> None:
        import threading

        from test_pdf2md import make_text_pdf

        src = make_text_pdf(self.tmp / "m.pdf", pages=1)
        fake = self.tmp / "slow_marker.py"
        pid_file = self.tmp / "pid.txt"
        fake.write_text(
            f"import os, time\nopen(r'{pid_file}', 'w').write(str(os.getpid()))\ntime.sleep(120)\n",
            encoding="utf-8",
        )
        cancel = threading.Event()
        seen = []

        def progress(update):
            seen.append(update.percent)
            if pid_file.exists():
                cancel.set()

        from unittest import mock

        with mock.patch.object(pdf2md, "find_marker_command", return_value=[sys.executable, str(fake)]):
            with self.assertRaises(pdf2md.ConversionCancelled):
                pdf2md.convert_file(src, self.tmp / "m.md", engine="marker", progress_callback=progress, cancel_event=cancel)
        pid = int(pid_file.read_text())
        alive = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True).stdout
        self.assertNotIn(str(pid), alive)
        self.assertFalse((self.tmp / "m.md").exists())


@unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
class OfficeUiTests(unittest.TestCase):
    def test_sheet_selector_only_for_excel_and_engine_only_for_pdf(self) -> None:
        from tkinter import Tk

        import pdf2md_ui
        from test_pdf2md import make_text_pdf

        with tempfile.TemporaryDirectory() as tmp:
            xlsx = make_xlsx(Path(tmp) / "b.xlsx")
            pdf = make_text_pdf(Path(tmp) / "a.pdf", pages=1)
            root = Tk()
            try:
                app = pdf2md_ui.App(root)
                app.pdf_path.set(str(xlsx))
                app._refresh_estimate()
                root.update()
                self.assertTrue(app.sheet_row.winfo_ismapped())
                self.assertEqual(list(app.sheet_box.cget("values")), ["Сводка", "Данные"])
                self.assertEqual(app.sheet.get(), "Данные")
                self.assertEqual(str(app.engine_box.cget("state")), "disabled")
                self.assertIn("листов: 2", app.estimate_label.cget("text"))

                app.sheet.set("Сводка")
                app._start_convert()
                app._worker_thread.join(30)
                deadline = time.monotonic() + 5
                while not app.progress_status.get().startswith("100%") and time.monotonic() < deadline:
                    root.update()  # the UI drains its queue every 100 ms
                    time.sleep(0.02)
                self.assertTrue(app.progress_status.get().startswith("100%"), app.progress_status.get())
                self.assertIn("## Сводка", xlsx.with_suffix(".md").read_text(encoding="utf-8"))

                app.pdf_path.set(str(pdf))
                app._refresh_estimate()
                root.update()
                self.assertFalse(app.sheet_row.winfo_ismapped())
                self.assertEqual(str(app.engine_box.cget("state")), "readonly")
            finally:
                root.destroy()


if __name__ == "__main__":
    unittest.main()
