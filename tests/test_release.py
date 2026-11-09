"""Regression coverage for safety, automatic engine choice and RU/EN preferences."""
from __future__ import annotations

import ast
import contextlib
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from string import Formatter
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app_settings
import i18n
import pdf2md
from test_pdf2md import make_text_pdf


class ReleaseTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.env = patch.dict(os.environ, {"PDF2MD_SETTINGS_FILE": str(self.root / "prefs.json")})
        self.env.start()
        i18n.set_language("ru")

    def tearDown(self):
        i18n.set_language("ru")
        self.env.stop()
        self.temp.cleanup()

    def test_cli_rejects_source_as_output_and_hardlink(self):
        src = self.root / "source.csv"
        original = b"Item,Count\nAlpha,12\n"
        src.write_bytes(original)
        alias = self.root / "alias.md"
        os.link(src, alias)
        for target in (src, alias, src.parent / "." / src.name):
            with contextlib.redirect_stderr(io.StringIO()) as error:
                result = pdf2md.main([str(src), "-o", str(target), "--language", "en"])
            self.assertEqual(result, 5)
            self.assertIn("must be different", error.getvalue())
            self.assertEqual(src.read_bytes(), original)

    def test_pdf_api_rejects_source_as_output(self):
        src = make_text_pdf(self.root / "source.pdf")
        original = src.read_bytes()
        with self.assertRaises(OSError):
            pdf2md.convert_pdf(src, src, engine="pymupdf")
        self.assertEqual(src.read_bytes(), original)

    def test_failed_publish_preserves_previous_output(self):
        out = self.root / "result.md"
        out.write_text("previous", encoding="utf-8")
        with patch.object(pdf2md.os, "replace", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError):
                pdf2md.write_markdown(out, "replacement")
        self.assertEqual(out.read_text(), "previous")
        self.assertEqual(list(self.root.glob(".pdf2md-*.tmp")), [])
        pdf2md.write_markdown(out, "new result")
        self.assertEqual(out.read_text(), "new result")

    def test_auto_detects_scan_after_text_pages_and_ignores_blank_page(self):
        src = make_text_pdf(self.root / "mixed.pdf", pages=4)
        lib = pdf2md._pymupdf()
        with lib.open(src) as doc:
            doc.new_page()
            doc.save(self.root / "blank.pdf")
            page = doc.new_page()
            pix = lib.Pixmap(lib.csRGB, lib.IRect(0, 0, 60, 80), False)
            pix.clear_with(180)
            page.insert_image(page.rect, pixmap=pix)
            doc.save(self.root / "scan.pdf")
        text = pdf2md.get_pdf_info(self.root / "blank.pdf")
        scan = pdf2md.get_pdf_info(self.root / "scan.pdf")
        self.assertFalse(text.probably_scanned)
        self.assertTrue(scan.probably_scanned)
        with patch.object(pdf2md, "marker_available", return_value=True), patch.object(pdf2md, "structured_available", return_value=True):
            self.assertEqual(pdf2md.resolve_engine("auto", text), "structured")
            self.assertEqual(pdf2md.resolve_engine("auto", scan), "marker")
            self.assertEqual(pdf2md.resolve_engine("pymupdf", scan), "pymupdf")
            self.assertIn("structured", pdf2md.format_estimate_range("auto", text))
            self.assertIn("marker", pdf2md.format_estimate_range("auto", scan))
        with patch.object(pdf2md, "marker_available", return_value=False), patch.object(pdf2md, "structured_available", return_value=False):
            self.assertEqual(pdf2md.resolve_engine("auto", scan), "pymupdf")

    def test_marker_progress_never_reverses_or_finishes_early(self):
        child = "import sys,time; " + "; ".join(
            f"sys.stderr.write('Stage: {value}%\\n'); sys.stderr.flush(); time.sleep(.3)"
            for value in (25, 100, 0, 50, 100))
        updates = []
        with subprocess.Popen([sys.executable, "-c", child], stderr=subprocess.PIPE, text=True) as proc:
            result = pdf2md._wait_marker_with_progress(proc, callback=updates.append,
                start_time=time.monotonic(), estimated_max=10, total_pages=1)
            self.assertTrue(proc.stderr.closed)
        values = [u.percent for u in updates]
        self.assertEqual(result, 0)
        self.assertGreater(max(values), 0)
        self.assertEqual(values, sorted(values))
        self.assertLess(max(values), 100)

    def test_repeated_transparent_chapter_art_is_not_a_scan(self):
        from PIL import Image
        png = io.BytesIO()
        Image.new("RGBA", (80, 100), (80, 120, 200, 100)).save(png, format="PNG")
        lib = pdf2md._pymupdf()
        with lib.open() as doc:
            for _ in range(4):
                page = doc.new_page()
                page.insert_image(page.rect, stream=png.getvalue())
                page.insert_text((72, 72), "Chapter title")
            src = self.root / "chapters.pdf"
            doc.save(src)
        self.assertFalse(pdf2md.get_pdf_info(src).probably_scanned)

    def test_preferences_validate_corruption_and_roundtrip(self):
        path = app_settings.settings_path()
        for value in ("{broken", "[]", '{"language":"xx","engine":12,"save_images":"no"}'):
            path.write_text(value, encoding="utf-8")
            self.assertEqual(app_settings.load_settings(), app_settings.DEFAULTS)
        settings = dict(app_settings.DEFAULTS, language="en", engine="structured", save_images=False, last_dir=str(self.root))
        app_settings.save_settings(settings)
        self.assertEqual(app_settings.load_settings(), settings)
        with patch.object(app_settings.os, "replace", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError):
                app_settings.save_settings(app_settings.DEFAULTS)
        self.assertEqual(app_settings.load_settings(), settings)

    def test_catalog_covers_messages_and_preserves_placeholders(self):
        def fields(text):
            return sorted((field, spec, conv) for _, field, spec, conv in Formatter().parse(text) if field is not None)
        for source, english in i18n.ENGLISH.items():
            self.assertEqual(fields(source), fields(english), source)
            self.assertFalse(re.search("[А-Яа-яЁё]", english), source)
        for name in ("pdf2md_ui.py", "pdf2md.py", "office2md.py", "marker_install.py"):
            tree = ast.parse(Path(pdf2md.__file__).with_name(name).read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "tr" and isinstance(node.args[0], ast.Constant):
                    self.assertIn(node.args[0].value, i18n.ENGLISH)
        i18n.set_language("en")
        self.assertEqual(pdf2md.format_duration(90), "~1 min 30 sec")
        import office2md
        self.assertEqual(office2md.format_cell_value(1234.5, "#,##0.00"), "1 234,50")

    def test_dnd_failure_still_creates_working_window(self):
        import tkinter
        from tkinterdnd2 import TkinterDnD
        import pdf2md_ui
        with patch.object(TkinterDnD, "_require", side_effect=RuntimeError("unavailable")):
            root = pdf2md_ui.make_root()
        try:
            app = pdf2md_ui.App(root)
            root.update()
            self.assertIs(root, tkinter._default_root)
            self.assertEqual(app.btn_convert.cget("text"), "Конвертировать")
        finally:
            root.destroy()

    def test_normal_window_registers_drop_on_root_and_entry(self):
        import pdf2md_ui
        root = pdf2md_ui.make_root()
        try:
            app = pdf2md_ui.App(root)
            self.assertTrue(root.TkdndVersion)
            self.assertTrue(hasattr(root, "drop_target_register"))
            for widget in (root, app.entry):
                self.assertTrue(root.tk.call("bind", str(widget), "<<Drop>>"))
            app.language.set("en")
            app._change_language()
            binding = root.tk.call("bind", str(app.entry), "<<Drop>>")
            self.assertTrue(binding)
            src = self.root / "drop sample.txt"
            src.write_text("Dropped text", encoding="utf-8")
            args = [""] * 18
            args[0], args[7], args[15] = "copy", root.tk.call("list", str(src)), str(app.entry)
            root.tk.call(str(binding).split()[0], *args)
            self.assertEqual(Path(app.pdf_path.get()), src)
            self.assertIn("Text", app.estimate_label.cget("text"))
        finally:
            root.destroy()

    def test_ui_language_switch_conversion_and_relaunch(self):
        import pdf2md_ui
        src = self.root / "sample.txt"
        src.write_text("Unchanged text", encoding="utf-8")
        root = pdf2md_ui.make_root()
        try:
            app = pdf2md_ui.App(root)
            app.engine.set("structured")
            app.save_images.set(False)
            app.pdf_path.set(str(src))
            app.language.set("en")
            app._change_language()
            root.update()
            self.assertEqual(app.btn_convert.cget("text"), "Convert")
            self.assertIn("Text", app.estimate_label.cget("text"))
            self.assertFalse(app.save_images.get())
            app._start_convert()
            deadline = time.monotonic() + 10
            while app.btn_convert.instate(["disabled"]) and time.monotonic() < deadline:
                root.update()
                time.sleep(.02)
            self.assertIn("Done:", app.progress_status.get())
            self.assertEqual(src.read_text(), "Unchanged text")
            self.assertTrue(src.with_suffix(".md").is_file())
            self.assertNotRegex(app.log.get("1.0", "end"), "[А-Яа-яЁё]")
            app._save_settings()
        finally:
            root.destroy()

        root = pdf2md_ui.make_root()
        try:
            app = pdf2md_ui.App(root)
            self.assertEqual(app.language.get(), "en")
            self.assertEqual(app.engine.get(), "structured")
            self.assertFalse(app.save_images.get())
            self.assertEqual(Path(app.settings["last_dir"]).resolve(), self.root.resolve())
            app.language.set("ru")
            app._change_language()
            self.assertEqual(app.btn_convert.cget("text"), "Конвертировать")
        finally:
            root.destroy()

    def test_english_overwrite_dialog_and_cancel(self):
        import tkinter as tk
        from tkinter import ttk
        import pdf2md_ui
        app_settings.save_settings(dict(app_settings.DEFAULTS, language="en"))
        root = pdf2md_ui.make_root()
        observed = []
        try:
            app = pdf2md_ui.App(root)
            src = self.root / "sample.txt"
            src.write_text("source", encoding="utf-8")
            src.with_suffix(".md").write_text("old output", encoding="utf-8")
            def answer(label):
                def visit(widget):
                    for child in widget.winfo_children():
                        yield child
                        yield from visit(child)
                dialog = next(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))
                buttons = [w for w in visit(dialog) if isinstance(w, ttk.Button)]
                observed.append([b.cget("text") for b in buttons])
                next(b for b in buttons if b.cget("text") == label).invoke()
            root.after(100, lambda: answer("No"))
            self.assertEqual(app._choose_output(src), src.with_name("sample_1.md"))
            root.after(100, lambda: answer("Cancel"))
            self.assertIsNone(app._choose_output(src))
            self.assertEqual(observed, [["Yes", "No", "Cancel"]] * 2)
            self.assertEqual(src.with_suffix(".md").read_text(), "old output")
        finally:
            root.destroy()

    def test_completion_replaces_progress_with_result_actions(self):
        import pdf2md_ui
        src = self.root / "report with spaces.txt"
        src.write_text("Completed document", encoding="utf-8")
        root = pdf2md_ui.make_root()
        try:
            app = pdf2md_ui.App(root)
            app.pdf_path.set(str(src))
            app._start_convert()
            deadline = time.monotonic() + 10
            while app.btn_convert.instate(["disabled"]) and time.monotonic() < deadline:
                root.update()
                time.sleep(.02)
            root.update()
            result = src.with_suffix(".md").resolve()
            self.assertEqual(app._last_result, result)
            self.assertFalse(app.progress.winfo_ismapped())
            self.assertTrue(app.result_actions.winfo_ismapped())
            self.assertEqual(app.progress_status.get(), "Готово: report with spaces.md")
            with patch.object(pdf2md_ui.os, "startfile", create=True) as launch:
                app.btn_open_result.invoke()
                app.btn_open_folder.invoke()
                self.assertEqual(launch.call_args_list, [unittest.mock.call(str(result)), unittest.mock.call(str(result.parent))])
            app.language.set("en")
            app._change_language()
            root.update()
            self.assertEqual(app.btn_open_result.cget("text"), "Open result")
            self.assertFalse(app.progress.winfo_ismapped())
            result.unlink()
            with patch.object(pdf2md_ui.messagebox, "showerror") as error, patch.object(pdf2md_ui.os, "startfile", create=True) as launch:
                app.btn_open_result.invoke()
                error.assert_called_once()
                launch.assert_not_called()
            # A second real conversion must restore progress, then complete again.
            app._start_convert()
            self.assertIsNone(app._last_result)
            self.assertEqual(app.result_actions.winfo_manager(), "")
            deadline = time.monotonic() + 10
            while app.btn_convert.instate(["disabled"]) and time.monotonic() < deadline:
                root.update()
                time.sleep(.02)
            self.assertTrue(result.is_file())
            self.assertEqual(app._last_result, result)
        finally:
            root.destroy()

    def test_header_and_github_link(self):
        import pdf2md_ui
        from pdf2md_theme import HeaderBand
        root = pdf2md_ui.make_root()
        try:
            app = pdf2md_ui.App(root)
            root.update()
            header = next(w for w in app.winfo_children() if isinstance(w, HeaderBand))
            texts = [header.itemcget(item, "text") for item in header.find_all() if header.type(item) == "text"]
            self.assertNotIn("PDF2MD", texts)
            self.assertTrue(any("Markdown" in text for text in texts))
            with patch.object(pdf2md_ui.webbrowser, "open") as browse:
                app.github_link.event_generate("<Button-1>")
                browse.assert_called_once_with("https://github.com/RootforgeIO")
        finally:
            root.destroy()
