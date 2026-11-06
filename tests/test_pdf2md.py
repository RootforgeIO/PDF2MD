"""Smoke/regression tests for PDF2MD.

Run from the project root:
    python -m unittest discover -s tests -v

Only PyMuPDF is required. Marker is emulated by a fake `marker_single`
script, so tests run in seconds and need no models or GPU.
"""

from __future__ import annotations

import os
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import marker_install  # noqa: E402
import pdf2md  # noqa: E402

lib = pdf2md._pymupdf()


def make_text_pdf(path: Path, pages: int = 3) -> Path:
    doc = lib.open()
    for i in range(pages):
        page = doc.new_page()
        page.insert_text((72, 72), f"Page {i + 1}. " + "Some extractable text. " * 4, fontsize=11)
    doc.save(path)
    doc.close()
    return path


def make_scanned_pdf(path: Path, pages: int = 2) -> Path:
    pix = lib.Pixmap(lib.csRGB, lib.IRect(0, 0, 60, 80), False)
    pix.clear_with(200)
    doc = lib.open()
    for _ in range(pages):
        page = doc.new_page()
        page.insert_image(page.rect, pixmap=pix)
    doc.save(path)
    doc.close()
    return path


FAKE_MARKER = textwrap.dedent(
    """
    import argparse, sys, time
    from pathlib import Path
    a = argparse.ArgumentParser()
    a.add_argument("pdf"); a.add_argument("--output_format"); a.add_argument("--output_dir")
    ns = a.parse_args()
    for stage in ("Recognizing layout", "Recognizing tables"):
        for i in range(0, 5):
            sys.stderr.write(f"\\r{stage}: {i * 25:3d}%| {i}/4"); sys.stderr.flush(); time.sleep(0.05)
        sys.stderr.write("\\n")
    stem = Path(ns.pdf).stem
    out = Path(ns.output_dir) / stem
    out.mkdir(parents=True)
    (out / f"{stem}.md").write_text(
        "# Заголовок из Marker\\n\\n![](_page_0_Picture_1.jpeg)\\n", encoding="utf-8"
    )
    (out / "_page_0_Picture_1.jpeg").write_bytes(b"\\xff\\xd8\\xff")
    (out / f"{stem}_meta.json").write_text("{}", encoding="utf-8")
    """
)


class Pdf2MdTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    # --- PyMuPDF path -------------------------------------------------------

    def test_pymupdf_conversion_writes_md_with_markers(self) -> None:
        pdf = make_text_pdf(self.tmp / "doc.pdf", pages=3)
        out = self.tmp / "doc.md"
        updates = []
        ok, used = pdf2md.convert_pdf(pdf, out, engine="pymupdf", progress_callback=updates.append)
        self.assertTrue(ok)
        self.assertEqual(used, "pymupdf")
        md = out.read_text(encoding="utf-8")
        self.assertIn("<!-- page: 1 -->", md)
        self.assertIn("<!-- page: 3 -->", md)
        self.assertEqual(md.count("\n---\n"), 2)
        self.assertEqual(updates[-1].percent, 100.0)

    def test_pymupdf_without_markers_and_separators(self) -> None:
        pdf = make_text_pdf(self.tmp / "doc.pdf", pages=2)
        md = pdf2md.convert_with_pymupdf_text(pdf, page_markers=False, page_separators=False)
        self.assertNotIn("<!-- page:", md)
        self.assertNotIn("---", md)
        self.assertIn("extractable text", md)

    def test_non_ascii_path(self) -> None:
        pdf = make_text_pdf(self.tmp / "мой файл.pdf", pages=1)
        out = pdf.with_suffix(".md")
        ok, _ = pdf2md.convert_pdf(pdf, out, engine="pymupdf")
        self.assertTrue(ok and out.exists())

    def test_scanned_pdf_detected_and_warned(self) -> None:
        pdf = make_scanned_pdf(self.tmp / "scan.pdf")
        self.assertTrue(pdf2md.get_pdf_info(pdf).probably_scanned)
        updates = []
        pdf2md.convert_pdf(pdf, self.tmp / "scan.md", engine="pymupdf", progress_callback=updates.append)
        self.assertTrue(any(u.message.startswith("WARN:") for u in updates))

    # --- Bad input -----------------------------------------------------------

    def test_encrypted_pdf_raises_readable_error(self) -> None:
        src = make_text_pdf(self.tmp / "src.pdf", pages=1)
        enc = self.tmp / "enc.pdf"
        doc = lib.open(src)
        doc.save(enc, encryption=lib.PDF_ENCRYPT_AES_256, user_pw="u", owner_pw="o")
        doc.close()
        with self.assertRaises(pdf2md.PdfOpenError) as ctx:
            pdf2md.get_pdf_info(enc)
        self.assertIn("паролем", str(ctx.exception))

    def test_not_a_pdf_raises_readable_error(self) -> None:
        fake = self.tmp / "fake.pdf"
        fake.write_text("not a pdf")
        with self.assertRaises(pdf2md.PdfOpenError):
            pdf2md.get_pdf_info(fake)

    def test_cli_exit_codes(self) -> None:
        fake = self.tmp / "fake.pdf"
        fake.write_text("not a pdf")
        self.assertEqual(pdf2md.main([str(self.tmp / "missing.pdf")]), 2)
        self.assertEqual(pdf2md.main([str(fake)]), 4)
        pdf = make_text_pdf(self.tmp / "ok.pdf", pages=1)
        self.assertEqual(pdf2md.main([str(pdf), "--engine", "pymupdf", "-o", str(self.tmp / "o" / "ok.md")]), 0)
        self.assertTrue((self.tmp / "o" / "ok.md").exists())

    # --- Engine resolution ---------------------------------------------------

    def test_auto_does_not_pick_marker_when_marker_not_installed(self) -> None:
        env = {k: v for k, v in os.environ.items() if not k.startswith("PDF2MD_")}
        with mock.patch.dict(os.environ, env, clear=True), mock.patch.object(
            pdf2md, "_marker_importable", return_value=False
        ), mock.patch.object(pdf2md.shutil, "which", return_value=None):
            self.assertIsNone(pdf2md.find_marker_command())
            with mock.patch.object(pdf2md, "structured_available", return_value=False):
                self.assertEqual(pdf2md.resolve_engine("auto"), "pymupdf")
            with mock.patch.object(pdf2md, "structured_available", return_value=True):
                self.assertEqual(pdf2md.resolve_engine("auto"), "structured")

    def test_setup_build_finds_marker_namespace_in_the_selected_folder(self) -> None:
        env = {k: v for k, v in os.environ.items() if not k.startswith("PDF2MD_")}
        env["LOCALAPPDATA"] = str(self.tmp / "profile")
        with mock.patch.dict(os.environ, env, clear=True), mock.patch.object(
            marker_install, "installed_build", return_value=True
        ):
            self.assertIsNone(pdf2md.find_marker_command())
            engine = self.tmp / "Выбранная папка" / "Marker"
            for name in ("marker/scripts/convert_single.py", "torch/__init__.py"):
                module = engine / "site" / name
                module.parent.mkdir(parents=True)
                module.touch()
            marker_install.user_data_dir().mkdir(parents=True)
            marker_install._register_marker(engine, "cpu")
            self.assertEqual(pdf2md.resolve_engine("auto", pdf2md.PdfInfo(1, 100, True)), "marker")
            self.assertEqual(pdf2md.find_marker_command()[-3:], ["--managed-install", str(engine), "cpu"])

    # --- Marker path (emulated) ---------------------------------------------

    def test_marker_cli_with_fake_marker(self) -> None:
        script = self.tmp / "fake_marker.py"
        script.write_text(FAKE_MARKER, encoding="utf-8")
        pdf = make_text_pdf(self.tmp / "doc.pdf", pages=2)
        out = self.tmp / "doc.md"
        updates = []
        with mock.patch.object(pdf2md, "find_marker_command", return_value=[sys.executable, str(script)]):
            ok, used = pdf2md.convert_pdf(pdf, out, engine="marker", progress_callback=updates.append)
        self.assertTrue(ok)
        self.assertEqual(used, "marker")
        md = out.read_text(encoding="utf-8")
        self.assertIn("Заголовок из Marker", md)
        self.assertIn("![](doc_images/_page_0_Picture_1.jpeg)", md)
        self.assertTrue((self.tmp / "doc_images" / "_page_0_Picture_1.jpeg").exists())
        self.assertFalse((self.tmp / "doc_images" / "doc_meta.json").exists())
        self.assertEqual(updates[-1].percent, 100.0)

    def test_marker_failure_returns_false_when_forced(self) -> None:
        script = self.tmp / "bad_marker.py"
        script.write_text("import sys; sys.stderr.write('boom\\n'); sys.exit(1)\n", encoding="utf-8")
        pdf = make_text_pdf(self.tmp / "doc.pdf", pages=1)
        updates = []
        with mock.patch.object(pdf2md, "find_marker_command", return_value=[sys.executable, str(script)]):
            ok, used = pdf2md.convert_pdf(pdf, self.tmp / "x.md", engine="marker", progress_callback=updates.append)
        self.assertFalse(ok)
        self.assertTrue(any("boom" in u.message for u in updates))

    # --- Structured engine (pymupdf4llm) ------------------------------------

    def _manual_pdf(self, name: str = "manual.pdf") -> Path:
        """Paragraph, raster image, paragraph, then a ruled table."""
        png = self.tmp / "pic.png"
        pix = lib.Pixmap(lib.csRGB, lib.IRect(0, 0, 120, 60), False)
        pix.clear_with(120)
        png.write_bytes(pix.tobytes("png"))
        doc = lib.open()
        page = doc.new_page()
        page.insert_text((50, 60), "Paragraph BEFORE the image.", fontsize=11)
        page.insert_image(lib.Rect(50, 80, 350, 230), filename=str(png))
        page.insert_text((50, 260), "Paragraph AFTER the image.", fontsize=11)
        rows = [["CC", "Name", "Range"], ["7", "Level", "0-127"], ["16", "Freq", "0-127"]]
        xs, y0, rh = [50, 110, 270, 390], 300, 22
        for r, row in enumerate(rows):
            for c, txt in enumerate(row):
                page.insert_text((xs[c] + 4, y0 + r * rh + 15), txt, fontsize=10)
        for r in range(len(rows) + 1):
            page.draw_line((xs[0], y0 + r * rh), (xs[-1], y0 + r * rh))
        for x in xs:
            page.draw_line((x, y0), (x, y0 + len(rows) * rh))
        pdf = self.tmp / name
        doc.save(pdf)
        doc.close()
        return pdf

    @unittest.skipUnless(pdf2md.structured_available(), "pymupdf4llm not installed")
    def test_structured_tables_and_images(self) -> None:
        pdf = self._manual_pdf("my manual.pdf")
        out = pdf.with_suffix(".md")
        updates = []
        ok, used = pdf2md.convert_pdf(pdf, out, engine="structured", progress_callback=updates.append)
        self.assertEqual((ok, used), (True, "structured"))
        md = out.read_text(encoding="utf-8")
        self.assertIn("|CC|Name|Range|", md.replace(" ", ""))
        self.assertIn("|16|Freq|0-127|", md.replace(" ", ""))
        images = list((self.tmp / "my manual_images").glob("*.png"))
        self.assertEqual(len(images), 1)
        link = f"![](my%20manual_images/{images[0].name})"
        self.assertIn(link, md)
        # reading order: before -> image -> after
        self.assertLess(md.index("BEFORE"), md.index(link))
        self.assertLess(md.index(link), md.index("AFTER"))
        self.assertEqual(updates[-1].percent, 100.0)

    @unittest.skipUnless(pdf2md.structured_available(), "pymupdf4llm not installed")
    def test_structured_without_images_leaves_no_folder(self) -> None:
        pdf = self._manual_pdf()
        out = pdf.with_suffix(".md")
        pdf2md.convert_pdf(pdf, out, engine="structured", save_images=False)
        self.assertFalse((self.tmp / "manual_images").exists())
        self.assertNotIn("![](", out.read_text(encoding="utf-8"))

    # --- Tables: merged cells, fake tables, reading order ----------------------

    def _draw_grid(self, page, xs, ys):
        for y in ys:
            page.draw_line((xs[0], y), (xs[-1], y))
        for x in xs:
            page.draw_line((x, ys[0]), (x, ys[-1]))

    @unittest.skipUnless(pdf2md.structured_available(), "pymupdf4llm not installed")
    def test_merged_cells_render_as_html_without_copying_neighbours(self) -> None:
        # Header row + 3 rows; the last column is ONE cell spanning all 3 rows.
        doc = lib.open()
        page = doc.new_page()
        xs, ys = [50, 130, 210, 380], [100, 122, 144, 166, 188]
        for y in ys:
            if y in (144, 166):
                page.draw_line((xs[0], y), (xs[2], y))  # no line through the merged cell
            else:
                page.draw_line((xs[0], y), (xs[-1], y))
        for x in xs:
            page.draw_line((x, ys[0]), (x, ys[-1]))
        for c, txt in enumerate(["Name", "LSB", "Value"]):
            page.insert_text((xs[c] + 4, 115), txt, fontsize=10)
        for r, (name, lsb) in enumerate([("Patch1", "08"), ("Patch2", "09"), ("Patch3", "10")]):
            page.insert_text((xs[0] + 4, 137 + r * 22), name, fontsize=10)
            page.insert_text((xs[1] + 4, 137 + r * 22), lsb, fontsize=10)
        page.insert_text((xs[2] + 4, 140), "0...127 shared", fontsize=10)
        pdf = self.tmp / "merged.pdf"
        doc.save(pdf)
        doc.close()
        out = pdf.with_suffix(".md")
        pdf2md.convert_pdf(pdf, out, engine="structured", save_images=False)
        md = out.read_text(encoding="utf-8")
        self.assertIn("<table>", md)
        self.assertIn('rowspan="3"', md)
        self.assertIn("0...127 shared", md)
        # PyMuPDF's default fill_empty would have put "09"/"10" into Value cells.
        self.assertNotIn("<td>09</td><td>09</td>", md)
        self.assertEqual(md.count("0...127 shared"), 1)

    @unittest.skipUnless(pdf2md.structured_available(), "pymupdf4llm not installed")
    def test_mostly_empty_grid_is_not_a_table(self) -> None:
        doc = lib.open()
        page = doc.new_page()
        self._draw_grid(page, [50, 100, 150, 200, 250, 300], [100, 130, 160, 190, 220])
        page.insert_text((55, 120), "1", fontsize=10)
        page.insert_text((50, 300), "Body text after the drawing.", fontsize=11)
        pdf = self.tmp / "drawing.pdf"
        doc.save(pdf)
        doc.close()
        md = pdf2md.convert_with_structured(pdf, self.tmp / "drawing.md", save_images=False)
        self.assertNotIn("|---", md)
        self.assertNotIn("<table>", md)
        self.assertIn("Body text after the drawing.", md)

    @unittest.skipUnless(pdf2md.structured_available(), "pymupdf4llm not installed")
    def test_left_column_table_comes_before_right_column_text(self) -> None:
        doc = lib.open()
        page = doc.new_page()
        page.insert_text((40, 60), "LEFT HEADING", fontsize=14)
        xs, ys = [40, 130, 270], [80 + 22 * i for i in range(20)]
        self._draw_grid(page, xs, ys)
        for r in range(19):
            page.insert_text((44, ys[r] + 15), f"K{r}", fontsize=9)
            page.insert_text((134, ys[r] + 15), f"V{r}", fontsize=9)
        page.insert_text((320, 60), "RIGHT HEADING", fontsize=14)
        for i in range(8):
            page.insert_text((320, 90 + i * 16), f"Right column line {i}.", fontsize=10)
        pdf = self.tmp / "twocol.pdf"
        doc.save(pdf)
        doc.close()
        md = pdf2md.convert_with_structured(pdf, self.tmp / "twocol.md", save_images=False)
        self.assertIn("| K0 | V0 |", md)
        self.assertLess(md.index("| K0 | V0 |"), md.index("RIGHT HEADING"))

    def test_render_table_regular_and_merged(self) -> None:
        import pdf2md_tables as t

        regular = [[t.Cell("A"), t.Cell("B|x")], [t.Cell("1"), t.Cell("line1\nline2")]]
        self.assertEqual(t.render_table(regular), "| A | B\\|x |\n|---|---|\n| 1 | line1<br>line2 |\n")
        merged = [[t.Cell("H1"), t.Cell("H2")], [t.Cell("a"), t.Cell("<v>", rowspan=2)], [t.Cell("b"), None]]
        html = t.render_table(merged)
        self.assertIn('<td rowspan="2">&lt;v&gt;</td>', html)
        self.assertIn("<tr><td>b</td></tr>", html)

    def test_strip_running_headers_and_page_numbers(self) -> None:
        chunks = []
        for i in range(1, 7):
            chunks.append(
                f"**microKORG XL**\n\n# Section {i}\n\nBody text {i} with number 12.\n\n**{i + 10}**"
            )
        chunks[2] = "**microKORG XL**\n\nIntro\n\n13\n\nTwo-column body 13 stays.\n\n13 | Assembly Instructions"
        out = pdf2md.strip_running_headers(chunks)
        joined = "\n".join(out)
        self.assertNotIn("microKORG XL", joined)
        self.assertNotIn("**11**", joined)
        self.assertNotIn("Assembly Instructions", joined)
        self.assertIn("Body text 1 with number 12.", joined)
        self.assertIn("Two-column body 13 stays.", joined)
        self.assertNotIn("\n13\n", "\n" + out[2] + "\n")
        # Too few pages: untouched.
        self.assertEqual(pdf2md.strip_running_headers(chunks[:3]), chunks[:3])

    def test_tidy_markdown(self) -> None:
        md = pdf2md._tidy_markdown("# **Title**\ntext\n## **Sub**\n\n\n\nline\n|a|b|\n|---|---|\n")
        self.assertEqual(md, "# Title\n\ntext\n\n## Sub\n\nline\n\n|a|b|\n|---|---|\n")

    # --- Helpers -------------------------------------------------------------

    def test_format_duration(self) -> None:
        self.assertEqual(pdf2md.format_duration(5), "~5 сек")
        self.assertEqual(pdf2md.format_duration(120), "~2 мин")
        self.assertEqual(pdf2md.format_duration(3 * 3600 + 5 * 60), "~3 ч 05 мин")


@unittest.skipUnless(os.environ.get("DISPLAY") or sys.platform == "win32", "needs a display")
class UiTests(unittest.TestCase):
    def test_tooltip_uses_toplevel_not_second_tk(self) -> None:
        from tkinter import Tk, Toplevel

        import pdf2md_ui

        root = Tk()
        try:
            app = pdf2md_ui.App(root)
            tip = pdf2md_ui.ToolTip(app.btn_convert, "hint")
            tip._show()
            root.update()
            self.assertIsInstance(tip._tip, Toplevel)
            tip._hide()
        finally:
            root.destroy()

    def test_luna_window_portrait_and_collapsible_log(self) -> None:
        from tkinter import Tk

        import pdf2md_ui

        root = Tk()
        try:
            app = pdf2md_ui.App(root)
            root.update()
            w, h = root.winfo_width(), root.winfo_height()
            self.assertGreater(h, w)  # a little taller than wide
            self.assertFalse(app.details.winfo_ismapped())
            app._log("WARN", "проверка")  # problems open the log by themselves
            root.update()
            self.assertTrue(app.details.winfo_ismapped())
            self.assertGreater(root.winfo_height(), h)
            app._toggle_details(False)
            root.update()
            self.assertEqual((root.winfo_width(), root.winfo_height()), (w, h))
            self.assertEqual(str(app.btn_convert.cget("style")), "Luna.Primary.TButton")
        finally:
            root.destroy()

    def test_engine_hints_cover_every_engine(self) -> None:
        import pdf2md_ui

        self.assertEqual(set(pdf2md_ui.ENGINE_HINTS), set(pdf2md.ENGINES))
        for name, text in pdf2md_ui.ENGINE_HINTS.items():
            self.assertIn("+ ", text, name)
            self.assertIn("− ", text, name)
        self.assertGreaterEqual(pdf2md_ui.ENGINE_HINT_DELAY_MS, 1000)

    def test_engine_list_item_shows_its_own_hint(self) -> None:
        from tkinter import Tk

        import pdf2md_ui

        root = Tk()
        try:
            app = pdf2md_ui.App(root)
            root.update()
            tip = app.engine_item_tip
            tip.delay_ms = 0
            root.tk.call("ttk::combobox::Post", app.engine_box)
            root.update()
            for idx, name in enumerate(pdf2md.ENGINES):
                _bx, by, _bw, bh = (int(v) for v in root.tk.call(tip._listbox, "bbox", idx))
                tip._on_motion(str(by + bh // 2))
                root.after(20)
                root.update()
                self.assertIsNotNone(tip._tip, name)
                label = tip._tip.winfo_children()[0].winfo_children()[0]
                self.assertEqual(label.cget("text"), pdf2md_ui.ENGINE_HINTS[name])
            root.tk.call("ttk::combobox::Unpost", app.engine_box)
            root.update()
            self.assertIsNone(tip._tip)

            # Closed box: hint follows the selected engine.
            app.engine.set("pymupdf")
            box_tip = pdf2md_ui.ToolTip(app.engine_box, lambda: pdf2md_ui.ENGINE_HINTS[app.engine.get()])
            box_tip._show()
            label = box_tip._tip.winfo_children()[0].winfo_children()[0]
            self.assertEqual(label.cget("text"), pdf2md_ui.ENGINE_HINTS["pymupdf"])
            box_tip._hide()
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
