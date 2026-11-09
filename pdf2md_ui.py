from __future__ import annotations

from i18n import tr

import os
import queue
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from tkinter import BOTH, END, LEFT, RIGHT, X, Y, BooleanVar, StringVar, Text, Tk, Toplevel, Menu, ttk
from tkinter import filedialog
import ui_dialogs as messagebox

import marker_install
import app_settings
import i18n
import office2md
import pdf2md
from pdf2md import ProgressUpdate, format_duration, format_estimate_range, get_pdf_info
from pdf2md_theme import (
    FIELD_BORDER,
    HeaderBand,
    LunaTheme,
    XPProgress,
    apply_luna,
    enable_dpi_awareness,
    make_app_icon,
)
from pdf2md_theme import TEXT as TEXT_COLOR


ENGINE_HINTS = {
    "auto": (
        "auto — выбор по содержимому PDF.\n\n"
        "+ Текстовые PDF: Structured, иначе PyMuPDF. Если есть страницы-сканы — Marker, когда он установлен.\n"
        "− Это эвристика: сложные таблицы требуют проверки результата. Без Marker сканы не распознаются."
    ),
    "marker": (
        "marker — нейросети: анализ вёрстки, OCR, распознавание таблиц.\n\n"
        "+ Лучшее качество: таблицы без рамок, формулы, сложные колонки, сканы (OCR), картинки.\n"
        "− Медленно: на CPU около 1 мин/стр., на GPU быстрее, но всё равно минуты; первый запуск "
        "загружает модели; шкала прогресса приблизительная.\n\n"
        "Когда: сканы, таблицы без рамок, формулы — и есть время подождать."
    ),
    "structured": (
        "structured — извлечение текстового слоя PDF с разметкой (pymupdf4llm).\n\n"
        "+ Очень быстро: секунды на сотню страниц. Заголовки, списки, жирный/курсив, "
        "таблицы в рамках (объединённые ячейки — HTML), картинки на своих местах, "
        "без колонтитулов. Текст не додумывается.\n"
        "− Таблицы без рамок выходят строками текста; сканы без текстового слоя "
        "не читаются (нет OCR); формулы — обычным текстом.\n\n"
        "Когда: мануалы и книги с текстовым слоем — основной режим."
    ),
    "pymupdf": (
        "pymupdf — сырой текст постранично.\n\n"
        "+ Самый быстрый и предсказуемый: весь текстовый слой как есть, минимум зависимостей.\n"
        "− Нет заголовков, таблиц и картинок; колонки могут перемешаться; сканы — пусто.\n\n"
        "Когда: нужен только текст (поиск, сверка) или другие движки дают сбой."
    ),
}

ENGINE_HINT_DELAY_MS = 1000
ESTIMATE_PLACEHOLDER = "Выберите файл кнопкой «Обзор…» или перетащите его в окно"

FILE_KINDS = {
    ".pdf": "PDF",
    ".docx": "Word (.docx)",
    ".doc": "Word 97-2003 (.doc, через Microsoft Word)",
    ".xlsx": "Excel (.xlsx)",
    ".xlsm": "Excel (.xlsm)",
    ".xls": "Excel 97-2003 (.xls, через Microsoft Excel)",
    ".csv": "CSV",
    ".tsv": "TSV",
    ".txt": "Текст (.txt)",
}
EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".xls")
OPEN_FILETYPES = [
    ("Все поддерживаемые", " ".join(f"*{s}" for s in pdf2md.SUPPORTED_SUFFIXES)),
    ("PDF", "*.pdf"),
    ("Word", "*.docx *.doc"),
    ("Excel", "*.xlsx *.xlsm *.xls"),
    ("CSV / TSV", "*.csv *.tsv"),
    ("Текст", "*.txt"),
    ("Все файлы", "*.*"),
]


def _show_tip_window(parent, text: str, x: int, y: int) -> Toplevel:
    # Toplevel, not a second Tk(): a second interpreter ignores the app's ttk styles
    # and multiple Tk roots are unsafe on Windows.
    tip = Toplevel(parent)
    tip.withdraw()
    tip.overrideredirect(True)
    tip.attributes("-topmost", True)
    frame = ttk.Frame(tip, padding=6, style="Tooltip.TFrame")
    frame.pack(fill=BOTH, expand=True)
    wrap = round(420 * parent.winfo_fpixels("1i") / 96)
    label = ttk.Label(frame, text=text, style="Tooltip.TLabel", justify=LEFT, wraplength=wrap)
    label.pack(fill=BOTH, expand=True)
    tip.update_idletasks()
    tip.geometry(f"+{x}+{y}")
    tip.deiconify()
    return tip


class ToolTip:
    """Hover hint. `text` may be a callable — then it is evaluated when the hint is shown."""

    def __init__(self, widget, text: str | Callable[[], str], delay_ms: int = 500):
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self._after_id = None
        self._tip = None
        widget._pdf2md_tooltip = self

        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _evt=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self._tip is not None:
            return
        text = self.text() if callable(self.text) else self.text
        if not text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self._tip = _show_tip_window(self.widget, text, x, y)

    def _hide(self, _evt=None):
        self._cancel()
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


class ComboboxItemToolTip:
    """Hint for the item under the mouse in an open ttk.Combobox drop-down list.

    The drop-down is a plain Tk listbox inside the combobox's popdown window; it has no
    Python wrapper, so the bindings are installed through Tcl.
    """

    def __init__(self, combobox: ttk.Combobox, hints: dict[str, str], delay_ms: int = 1000):
        self.combobox = combobox
        self.hints = hints
        self.delay_ms = delay_ms
        self._after_id = None
        self._tip = None
        self._index: int | None = None

        tk = combobox.tk
        popdown = str(tk.call("ttk::combobox::PopdownWindow", combobox))
        self._listbox = f"{popdown}.f.l"
        motion = combobox.register(self._on_motion)
        hide = combobox.register(self._hide)
        tk.call("bind", self._listbox, "<Motion>", f"+{motion} %y")
        tk.call("bind", self._listbox, "<Leave>", f"+{hide}")
        tk.call("bind", self._listbox, "<ButtonPress>", f"+{hide}")
        tk.call("bind", popdown, "<Unmap>", f"+{hide}")

    def _on_motion(self, y: str) -> None:
        index = int(self.combobox.tk.call(self._listbox, "nearest", int(y)))
        if index == self._index:
            return
        self._hide()
        self._index = index
        self._after_id = self.combobox.after(self.delay_ms, self._show)

    def _show(self) -> None:
        self._after_id = None
        if self._index is None or self._tip is not None:
            return
        tk = self.combobox.tk
        item = str(tk.call(self._listbox, "get", self._index))
        text = self.hints.get(item)
        bbox = tk.call(self._listbox, "bbox", self._index)
        if not text or not bbox:
            return
        _bx, by, _bw, _bh = (int(v) for v in bbox)
        x = int(tk.call("winfo", "rootx", self._listbox)) + int(tk.call("winfo", "width", self._listbox)) + 8
        y = int(tk.call("winfo", "rooty", self._listbox)) + by
        self._tip = _show_tip_window(self.combobox, text, x, y)

    def _hide(self) -> None:
        self._index = None
        if self._after_id is not None:
            try:
                self.combobox.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


@dataclass(frozen=True)
class UiOptions:
    engine: str  # auto|marker|structured|pymupdf
    sheet: str  # Excel sheet name ("" = active)
    page_markers: bool
    page_separators: bool
    save_images: bool
    strip_headers: bool


class App(ttk.Frame):
    def __init__(self, master: Tk, theme: LunaTheme | None = None):
        self.settings = app_settings.load_settings()
        i18n.set_language(self.settings["language"])
        self.theme = theme or apply_luna(master)
        super().__init__(master, padding=0)
        self.master = master
        self.pack(fill=BOTH, expand=True)

        self.pdf_path = StringVar(value="")
        self.engine = StringVar(value=self.settings["engine"])
        self.language = StringVar(value=self.settings["language"])
        self.sheet = StringVar(value="")
        self._sheets_for: tuple[str, float] | None = None  # (path, mtime) the sheet list was read for
        self.page_markers = BooleanVar(value=self.settings["page_markers"])
        self.page_separators = BooleanVar(value=self.settings["page_separators"])
        self.save_images = BooleanVar(value=self.settings["save_images"])
        self.strip_headers = BooleanVar(value=self.settings["strip_headers"])
        self.progress_status = StringVar(value=tr('Готово к конвертации'))

        self._worker_thread: threading.Thread | None = None
        self._last_result: Path | None = None
        self._messages: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self._poll_after_id: str | None = None
        self._cancel_event = threading.Event()
        self.log_file = None  # set by install_error_logging(): the log also goes there
        self.marker_install_index = None
        self._settings_after_id = None

        self._details_open = False
        self._build_ui()
        self._poll_messages()
        master.protocol("WM_DELETE_WINDOW", self._on_close)

        self.engine.trace_add("write", lambda *_: self._refresh_estimate())
        for name in ("engine", "page_markers", "page_separators", "save_images", "strip_headers"):
            getattr(self, name).trace_add("write", self._schedule_settings)

    def _schedule_settings(self, *_):
        if self._settings_after_id is not None:
            self.after_cancel(self._settings_after_id)
        self._settings_after_id = self.after(300, self._save_settings)

    def _save_settings(self):
        if self._settings_after_id is not None:
            self.after_cancel(self._settings_after_id)
            self._settings_after_id = None
        for name in ("engine", "page_markers", "page_separators", "save_images", "strip_headers", "language"):
            self.settings[name] = getattr(self, name).get()
        try:
            app_settings.save_settings(self.settings)
        except OSError as exc:
            self._log("WARN", tr("Не удалось сохранить настройки: {error}").format(error=exc))

    def _build_menu(self):
        previous = getattr(self, "menu", None)
        self.menu = Menu(self.master, tearoff=False)
        self.settings_menu = Menu(self.menu, tearoff=False)
        languages = Menu(self.settings_menu, tearoff=False)
        for code, label in (("ru", "Русский"), ("en", "English")):
            languages.add_radiobutton(label=label, variable=self.language, value=code,
                                     command=self._change_language)
        self.settings_menu.add_cascade(label=tr("Язык"), menu=languages)
        self.marker_install_index = None
        if marker_install.installed_build() and marker_install.installed_marker_command() is None:
            self.settings_menu.add_separator()
            self.settings_menu.add_command(label=tr('Установить Marker…'), command=self._start_marker_install)
            self.marker_install_index = self.settings_menu.index("end")
        self.menu.add_cascade(label=tr("Настройки"), menu=self.settings_menu)
        self.master.config(menu=self.menu)
        if previous is not None:
            previous.destroy()

    def _change_language(self):
        if self._worker_thread and self._worker_thread.is_alive():
            self.language.set(i18n.get_language())
            return
        self._save_settings()
        i18n.set_language(self.language.get())
        self.engine_item_tip._hide()
        previous_log = self.log.get("1.0", "end-1c")
        def clear(widget):
            tip = getattr(widget, "_pdf2md_tooltip", None)
            if tip is not None:
                tip._hide()
            for child in widget.winfo_children():
                clear(child)
        clear(self)
        for child in self.winfo_children():
            child.destroy()
        self._details_open = False
        self._sheets_for = None
        self._build_ui()
        self.log.insert(END, previous_log)
        self.progress_status.set(tr("Готово к конвертации"))
        self._show_result()
        self._refresh_estimate()

    def _build_ui(self):
        t = self.theme
        px = t.px
        self.master.title("PDF2MD")
        self._build_menu()
        self._set_icon()

        HeaderBand(self, t, "", "PDF, Word, Excel, CSV, TXT  →  Markdown").pack(fill=X)
        body = ttk.Frame(self, padding=(px(14), px(12), px(14), px(10)))
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=1)
        self._body = body

        # --- file --------------------------------------------------------------------------
        ttk.Label(body, text=tr('Файл'), style="Bold.TLabel").grid(row=0, column=0, sticky="w")
        row = ttk.Frame(body)
        row.grid(row=1, column=0, sticky="ew", pady=(px(3), 0))
        row.columnconfigure(0, weight=1)
        entry = ttk.Entry(row, textvariable=self.pdf_path, style="Luna.TEntry")
        entry.grid(row=0, column=0, sticky="ew")
        entry.bind("<FocusOut>", lambda _e: self._refresh_estimate())
        entry.bind("<Return>", lambda _e: self._refresh_estimate())
        self.entry = entry
        btn_browse = ttk.Button(row, text=tr('Обзор…'), command=self._pick_pdf, style="Luna.TButton")
        btn_browse.grid(row=0, column=1, padx=(px(6), 0))
        ToolTip(
            btn_browse,
            tr('Выбор файла: PDF, Word (.docx, .doc), Excel (.xlsx, .xlsm, .xls), CSV/TSV, TXT.\nФайл можно и просто перетащить в окно.\nРезультат сохраняется рядом с ним: <имя>.md (картинки — в <имя>_images).'),
        )
        self.estimate_label = ttk.Label(
            body, text=tr(ESTIMATE_PLACEHOLDER), style="Muted.TLabel", wraplength=px(360), justify=LEFT
        )
        self.estimate_label.grid(row=2, column=0, sticky="w", pady=(px(4), 0))

        ttk.Separator(body, style="Luna.TSeparator").grid(row=3, column=0, sticky="ew", pady=px(10))

        # --- settings ------------------------------------------------------------------------
        opts = ttk.Frame(body)
        opts.grid(row=4, column=0, sticky="ew")
        opts.columnconfigure(1, weight=1)
        ttk.Label(opts, text=tr('Движок:')).grid(row=0, column=0, sticky="w")
        engine_box = ttk.Combobox(
            opts, textvariable=self.engine, values=list(pdf2md.ENGINES), state="readonly", width=16,
            style="Luna.TCombobox",
        )
        engine_box.grid(row=0, column=1, sticky="w", padx=(px(8), 0))
        self.engine_box = engine_box
        # Closed box: hint for the selected engine; open list: hint for the item under the mouse.
        ToolTip(engine_box, lambda: tr(ENGINE_HINTS.get(self.engine.get(), "")), delay_ms=ENGINE_HINT_DELAY_MS)
        self.engine_item_tip = ComboboxItemToolTip(engine_box, {key: tr(value) for key, value in ENGINE_HINTS.items()}, delay_ms=ENGINE_HINT_DELAY_MS)

        # Shown only for Excel files (see _update_for_file_type).
        self.sheet_row = ttk.Frame(opts)
        self.sheet_row.columnconfigure(1, weight=1)
        ttk.Label(self.sheet_row, text=tr('Лист:')).grid(row=0, column=0, sticky="w")
        self.sheet_box = ttk.Combobox(
            self.sheet_row, textvariable=self.sheet, state="readonly", width=30, style="Luna.TCombobox"
        )
        self.sheet_box.grid(row=0, column=1, sticky="ew", padx=(px(22), 0))
        ToolTip(
            self.sheet_box,
            tr('Excel: в .md попадает один лист — выберите какой. По умолчанию — активный лист книги\n(тот, что был открыт при сохранении). Объединённые ячейки сохраняются (HTML-таблица),\nформулы — их значения, сохранённые Excel. В CLI: --sheet "Имя" или номер с 1.'),
        )
        self._opts = opts

        checks = ttk.Frame(opts)
        checks.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(px(8), 0))
        checks.columnconfigure(0, weight=1)
        checks.columnconfigure(1, weight=1)

        def check(text: str, var, r: int, c: int, tip: str):
            cb = ttk.Checkbutton(checks, text=text, variable=var, style="Luna.TCheckbutton")
            cb.grid(row=r, column=c, sticky="w", pady=px(1))
            ToolTip(cb, tip)
            return cb

        check(
            tr('Сохранять картинки'), self.save_images, 0, 0,
            tr('Для PDF (structured, marker), Word и Excel: картинки сохраняются рядом с .md в папку\n<имя>_images, в тексте — относительные ссылки ![](<имя>_images/...). Для pymupdf не действует.'),
        )
        chk_hdr = check(
            tr('Убирать колонтитулы'), self.strip_headers, 0, 1,
            tr('Для structured: убирает строки, повторяющиеся вверху/внизу страниц\n(название документа, «12 | Раздел») и номера страниц. Основной текст не трогается.'),
        )
        chk1 = check(
            tr('Маркеры страниц'), self.page_markers, 1, 0,
            tr('Вставляет строки вида <!-- page: N --> перед текстом каждой страницы (HTML-комментарии,\nв просмотре не видны). Полезно для поиска/ссылок и контроля соответствия. Для structured и pymupdf.'),
        )
        chk2 = check(
            tr('Линии между страницами'), self.page_separators, 1, 1,
            tr('Между страницами вставляется разделитель `---` (горизонтальная линия).\nНе влияет на смысл текста, только на удобство чтения.'),
        )
        # Widgets that only make sense for PDF; disabled for Word/Excel/CSV/TXT.
        self._pdf_only_widgets = [engine_box, chk_hdr, chk1, chk2]

        ttk.Separator(body, style="Luna.TSeparator").grid(row=5, column=0, sticky="ew", pady=px(10))

        # --- progress ------------------------------------------------------------------------
        self.progress = XPProgress(body, t)
        self.progress.grid(row=6, column=0, sticky="ew")
        self.result_actions = ttk.Frame(body)
        self.result_actions.grid(row=6, column=0, sticky="w")
        self.btn_open_result = ttk.Button(
            self.result_actions, text=tr('Открыть результат'), command=self._open_result,
            style="Luna.Primary.TButton")
        self.btn_open_result.pack(side=LEFT)
        self.btn_open_folder = ttk.Button(
            self.result_actions, text=tr('Открыть папку'), command=lambda: self._open_result(folder=True),
            style="Luna.TButton")
        self.btn_open_folder.pack(side=LEFT, padx=(px(6), 0))
        self.btn_clear_result = ttk.Button(
            self.result_actions, text="×", width=3, command=self._clear_result,
            style="Luna.TButton")
        self.btn_clear_result.pack(side=LEFT, padx=(px(6), 0))
        ToolTip(self.btn_clear_result, tr('Очистить результат'))
        self.result_actions.grid_remove()
        self.progress_label = ttk.Label(
            body, textvariable=self.progress_status, wraplength=px(360), justify=LEFT
        )
        self.progress_label.grid(row=7, column=0, sticky="w", pady=(px(5), 0))
        ToolTip(self.progress_label, lambda: str(self._last_result or ""))

        actions = ttk.Frame(body)
        actions.grid(row=8, column=0, sticky="ew", pady=(px(12), 0))
        self.details_link = ttk.Label(actions, text=tr('▸ Подробности'), style="Link.TLabel", cursor="hand2")
        self.details_link.pack(side=LEFT)
        self.details_link.bind("<Button-1>", lambda _e: self._toggle_details())
        ToolTip(self.details_link, tr('Журнал конвертации: что сделано, предупреждения, ошибки.'))
        self.btn_cancel = ttk.Button(
            actions, text=tr('Отмена'), command=self._cancel, state="disabled", style="Luna.TButton"
        )
        self.btn_cancel.pack(side=RIGHT)
        ToolTip(self.btn_cancel, tr('Остановить конвертацию. Файл .md не записывается; процесс Marker завершается.'))
        self.btn_convert = ttk.Button(
            actions, text=tr('Конвертировать'), command=self._start_convert, style="Luna.Primary.TButton"
        )
        self.btn_convert.pack(side=RIGHT, padx=(0, px(6)))
        ToolTip(
            self.btn_convert,
            tr('Запускает конвертацию. Файл .md будет создан рядом с исходным файлом\n(если такой уже есть — программа спросит, перезаписать ли его).'),
        )

        # --- details (collapsible log) ---------------------------------------------------------
        self.details = ttk.Frame(body)
        self.log = Text(
            self.details, height=9, wrap="word", relief="flat", borderwidth=0, font=t.font(8),
            background="#FFFFFF", foreground=TEXT_COLOR, highlightthickness=1,
            highlightbackground=FIELD_BORDER, highlightcolor=FIELD_BORDER, padx=px(4), pady=px(3),
        )
        self.log.pack(fill=BOTH, expand=True)
        for tag, color in (("WARN", "#9C5700"), ("ERROR", "#C00000"), ("DONE", "#1E6E12")):
            self.log.tag_configure(tag, foreground=color)
        body.rowconfigure(9, weight=1)

        ttk.Label(
            body, text=tr('Результат сохраняется рядом с исходным файлом: <имя>.md'), style="Small.TLabel"
        ).grid(row=10, column=0, sticky="w", pady=(px(8), 0))

        self.github_link = ttk.Label(body, text="GitHub · RootforgeIO", style="Link.TLabel",
                                     cursor="hand2", takefocus=True)
        self.github_link.grid(row=11, column=0, sticky="w", pady=(px(5), 0))
        for event in ("<Button-1>", "<Return>", "<space>"):
            self.github_link.bind(event, lambda _event: webbrowser.open("https://github.com/RootforgeIO"))

        self._enable_drop_target()
        self._fit_window()

    def _clear_result(self):
        if self._last_result is None or (self._worker_thread and self._worker_thread.is_alive()):
            return
        self._last_result = None
        self.result_actions.grid_remove()
        self.progress.config(value=0)
        self.progress.grid()
        self.progress_status.set(tr('Готово к конвертации'))
        self._toggle_details(False)
        self._fit_window()

    def _show_result(self):
        if self._last_result is not None:
            self.progress.grid_remove()
            self.result_actions.grid()
            self.progress_status.set(tr('Готово: {name}').format(name=self._last_result.name))
            self._fit_window()

    def _open_result(self, folder: bool = False):
        if self._last_result is None:
            return
        target = self._last_result.parent if folder else self._last_result
        try:
            if not target.exists():
                raise FileNotFoundError(tr('Файл или папка больше не существует.'))
            if sys.platform == "win32":
                os.startfile(str(target))
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(target)])
        except OSError as exc:
            messagebox.showerror(tr('Не удалось открыть'),
                tr('Не удалось открыть файл или папку: {error}\nДля Markdown выберите подходящее приложение в Windows.').format(error=exc),
                parent=self.master)

    # --- window helpers ----------------------------------------------------------------------

    def _set_icon(self) -> None:
        try:
            from PIL import ImageTk

            self._icons = [ImageTk.PhotoImage(make_app_icon(s)) for s in (16, 24, 32, 48, 64)]
            self.master.iconphoto(True, *self._icons)
        except Exception:
            pass  # cosmetic

    def _fit_window(self) -> None:
        """Narrow, a little taller than wide (the user's wish); the log adds height when opened."""
        self.master.update_idletasks()
        width = self.theme.px(400)
        height = max(self.master.winfo_reqheight(), round(width * 1.06))
        self.master.geometry(f"{width}x{height}")
        self.master.minsize(width, height)

    def _toggle_details(self, open_: bool | None = None) -> None:
        open_ = (not self._details_open) if open_ is None else open_
        if open_ == self._details_open:
            return
        self._details_open = open_
        self.details_link.config(text=tr('▾ Подробности') if open_ else tr('▸ Подробности'))
        self.master.update_idletasks()
        before = self.master.winfo_height()
        min_w, min_h = self.master.minsize()
        if open_:
            self.details.grid(row=9, column=0, sticky="nsew", pady=(self.theme.px(8), 0))
            self.master.update_idletasks()
            # remember exactly what was added: the log stretches with the window afterwards
            self._details_extra = self.details.winfo_reqheight() + self.theme.px(8)
            self.master.minsize(min_w, min_h + self._details_extra)
            self.master.geometry(f"{self.master.winfo_width()}x{before + self._details_extra}")
        else:
            extra = getattr(self, "_details_extra", 0)
            self.details.grid_remove()
            self.master.minsize(min_w, max(1, min_h - extra))
            self.master.geometry(f"{self.master.winfo_width()}x{max(self.master.minsize()[1], before - extra)}")

    def _enable_drop_target(self) -> None:
        """Drop a file from Explorer anywhere on the window (tkinterdnd2, if the root supports it)."""
        try:
            from tkinterdnd2 import DND_FILES
        except Exception:
            return
        if not hasattr(self.master, "drop_target_register"):
            return  # root is a plain Tk(), not TkinterDnD.Tk()

        def on_drop(event):
            paths = self.master.tk.splitlist(event.data)
            if paths:
                self.pdf_path.set(paths[0])
                self._refresh_estimate()
                if len(paths) > 1:
                    self._log("WARN", tr('Перетащено файлов: {v0} — взят первый (пакетная обработка — позже)').format(v0=len(paths)))
            return event.action

        def register(widget):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", on_drop)
            except Exception:
                pass
            for child in widget.winfo_children():
                register(child)

        register(self.master)

    def _pick_pdf(self):
        path = filedialog.askopenfilename(title=tr('Выберите файл'),
            filetypes=[(tr(label), pattern) for label, pattern in OPEN_FILETYPES],
            initialdir=self.settings["last_dir"] or None, parent=self.master)
        if path:
            self.pdf_path.set(path)
            self._refresh_estimate()

    def _update_for_file_type(self, src: Path | None) -> None:
        """Engine and page options apply to PDF only; the sheet selector to Excel only."""
        suffix = src.suffix.lower() if src else ".pdf"
        is_pdf = suffix == ".pdf"
        for w in self._pdf_only_widgets:
            enabled = "readonly" if isinstance(w, ttk.Combobox) else "normal"
            w.config(state=enabled if is_pdf else "disabled")

        if src is None or suffix not in EXCEL_SUFFIXES:
            self.sheet_row.grid_remove()
            self._sheets_for = None
            return
        self.sheet_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(self.theme.px(6), 0))
        key = (str(src), src.stat().st_mtime)
        if self._sheets_for == key:
            return
        names, active = office2md.list_sheets(src)
        self.sheet_box.config(values=names)
        self.sheet.set(names[active] if names else "")
        self._sheets_for = key

    def _refresh_estimate(self):
        pdf = self.pdf_path.get().strip().strip('"')
        if not pdf:
            self._update_for_file_type(None)
            self.estimate_label.config(text=tr(ESTIMATE_PLACEHOLDER))
            return

        pdf_path = Path(pdf)
        if not pdf_path.is_file() or not pdf2md.is_supported_file(pdf_path):
            self._update_for_file_type(None)
            self.estimate_label.config(
                text=tr('Укажите существующий файл: ') + ", ".join(pdf2md.SUPPORTED_SUFFIXES)
            )
            return

        directory = str(pdf_path.resolve().parent)
        if directory != self.settings["last_dir"]:
            self.settings["last_dir"] = directory
            self._schedule_settings()
        if pdf_path.suffix.lower() != ".pdf":
            kind = tr(FILE_KINDS.get(pdf_path.suffix.lower(), pdf_path.suffix))
            size_kb = max(1, round(pdf_path.stat().st_size / 1024))
            if pdf_path.suffix.lower() == ".xls":
                # the sheet list needs Excel to re-save the file first: a few seconds
                self.estimate_label.config(text=tr('{v0}, {v1} КБ · читаю листы через Microsoft Excel…').format(v0=kind, v1=size_kb))
                self.update_idletasks()
            try:
                self._update_for_file_type(pdf_path)
            except Exception as e:
                self.estimate_label.config(text=tr('{v0}, {v1} КБ · не удалось прочитать: {v2}').format(v0=kind, v1=size_kb, v2=e))
                return
            extra = ""
            if pdf_path.suffix.lower() in EXCEL_SUFFIXES:
                extra = tr(' · листов: {v0}').format(v0=len(self.sheet_box.cget('values')))
            self.estimate_label.config(text=tr('{v0}, {v1} КБ{v2} · обычно секунды').format(v0=kind, v1=size_kb, v2=extra))
            return

        self._update_for_file_type(pdf_path)
        try:
            info = get_pdf_info(pdf_path)
            text = format_estimate_range(self.engine.get(), info)
            if info.probably_scanned:
                text += tr(' · есть страницы без текстового слоя (OCR дольше)')
            self.estimate_label.config(text=text)
        except Exception as e:
            self.estimate_label.config(text=tr('Не удалось оценить: {v0}').format(v0=e))

    def _log(self, level: str, msg: str):
        tag = level if level in ("WARN", "ERROR", "DONE") else ""
        self.log.insert(END, f"[{level}] {msg}\n", tag)
        self.log.see(END)
        if level in ("WARN", "ERROR"):
            self._toggle_details(True)  # problems are never hidden in a collapsed log
        if self.log_file is not None:
            try:
                self.log_file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} [{level}] {msg}\n")
            except Exception:
                pass

    def _format_progress_status(self, update: ProgressUpdate) -> str:
        parts = [f"{update.percent:.0f}%"]
        if update.current_page and update.total_pages:
            parts.append(tr('стр. {v0}/{v1}').format(v0=update.current_page, v1=update.total_pages))
        if update.eta_seconds is not None and update.percent < 100:
            parts.append(tr('осталось {v0}').format(v0=format_duration(update.eta_seconds)))
        elif update.elapsed_seconds is not None and update.percent >= 100:
            parts.append(tr('за {v0}').format(v0=format_duration(update.elapsed_seconds)))
        parts.append(update.message)
        return " · ".join(parts)

    def _set_busy(self, busy: bool):
        self.menu.entryconfig(0, state="disabled" if busy else "normal")
        if busy:
            self._last_result = None
            self.result_actions.grid_remove()
            self.progress.grid()
            self._cancel_event = threading.Event()
            self.btn_convert.config(state="disabled")
            self.btn_cancel.config(state="normal")
            self.progress.config(value=0)
        else:
            self.btn_convert.config(state="normal")
            self.btn_cancel.config(state="disabled")

    def _cancel(self):
        if self._worker_thread and self._worker_thread.is_alive():
            self._cancel_event.set()
            self.btn_cancel.config(state="disabled")
            self.progress_status.set(tr('Отмена… (дожидаюсь остановки)'))

    def _on_close(self):
        """Closing while converting: ask, then stop the conversion (and Marker) before exiting."""
        worker = self._worker_thread
        if worker and worker.is_alive():
            if not messagebox.askyesno(
                tr('Идёт обработка'), tr('Прервать текущую операцию и закрыть программу?'), parent=self.master
            ):
                return
            self._cancel_event.set()
            deadline = time.monotonic() + 15
            while worker.is_alive() and time.monotonic() < deadline:
                self.master.update()
                worker.join(0.05)
        self._save_settings()
        self.master.destroy()

    def _start_marker_install(self):
        if self._worker_thread and self._worker_thread.is_alive():
            return
        options = self._choose_marker_install()
        if options is None:
            return
        self._set_busy(True)
        self.progress_status.set(tr('Подготовка установки Marker…'))
        self._worker_thread = threading.Thread(
            target=self._install_marker_worker, args=options
        )
        self._worker_thread.start()

    def _choose_marker_install(self) -> tuple[Path, str] | None:
        dialog = Toplevel(self.master)
        dialog.title(tr('Установить Marker'))
        dialog.transient(self.master)
        dialog.resizable(False, False)
        px = self.theme.px
        body = ttk.Frame(dialog, padding=px(12))
        body.pack(fill=BOTH, expand=True)
        default = marker_install.default_download_dir()
        folder = StringVar(master=dialog, value=str(default))
        choice = None

        def browse():
            selected = Path(folder.get())
            path = filedialog.askdirectory(
                title=tr('Выберите папку для Marker'), parent=dialog,
                initialdir=selected if selected.is_dir() else default.parent,
            )
            if path:
                folder.set(path)

        def accept():
            nonlocal choice
            path = Path(folder.get().strip().strip('"'))
            if not path.is_absolute():
                messagebox.showerror(tr('Папка для Marker'), tr('Укажите полный путь к папке.'), parent=dialog)
                return
            choice = path, ("cpu", "cuda")[device.current()]
            dialog.destroy()

        ttk.Label(body, text=tr('Папка для Marker и моделей')).grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(body, textvariable=folder, width=38, style="Luna.TEntry")
        entry.grid(row=1, column=0, sticky="ew", pady=(px(4), px(10)))
        ttk.Button(body, text=tr('Обзор…'), command=browse, style="Luna.TButton").grid(
            row=1, column=1, sticky="n", padx=(px(6), 0), pady=px(2)
        )
        ttk.Label(body, text=tr('Вариант')).grid(row=2, column=0, sticky="w")
        device = ttk.Combobox(body, values=("CPU", "NVIDIA CUDA 12.6"), state="readonly", style="Luna.TCombobox")
        device.current(0)
        device.grid(row=3, column=0, sticky="w", pady=(px(4), px(8)))
        ttk.Label(body, text=tr('Для загрузки понадобится несколько ГБ.\nДля CUDA нужен совместимый драйвер NVIDIA.')).grid(
            row=4, column=0, columnspan=2, sticky="w"
        )
        actions = ttk.Frame(body)
        actions.grid(row=5, column=0, columnspan=2, sticky="e", pady=(px(12), 0))
        ttk.Button(actions, text=tr('Отмена'), command=dialog.destroy, style="Luna.TButton").pack(side=RIGHT)
        ttk.Button(actions, text=tr('Установить'), command=accept, style="Luna.Primary.TButton").pack(
            side=RIGHT, padx=(0, px(6))
        )
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.wait_visibility()
        dialog.grab_set()
        entry.focus_set()
        self.master.wait_window(dialog)
        return choice

    def _install_marker_worker(self, folder: Path, device: str):
        def report(percent, message):
            self._messages.put(("PROGRESS", ProgressUpdate(percent=percent, message=message)))

        try:
            marker_install.install_marker(folder, device, self._cancel_event, report)
        except pdf2md.ConversionCancelled:
            self._messages.put(("CANCELLED", tr('Установка Marker отменена')))
        except Exception as exc:
            self._messages.put(("INFO", str(exc)))
            self._messages.put(("ERROR", tr('Не удалось установить Marker. Причина в журнале.')))
        else:
            self._messages.put(("MARKER_INSTALLED", tr('Marker установлен')))

    def _choose_output(self, src: Path) -> Path | None:
        """`<name>.md`; if it already exists — overwrite, a free `<name>_N.md`, or cancel."""
        out = src.with_suffix(".md")
        if not out.exists():
            return out
        n = 1
        while (alt := src.with_name(f"{src.stem}_{n}.md")).exists():
            n += 1
        answer = messagebox.askyesnocancel(
            tr('Файл уже существует'),
            tr('{v0} уже есть в этой папке.\n\nДа — перезаписать его\nНет — сохранить как {v1}\nОтмена — не конвертировать').format(v0=out.name, v1=alt.name),
            parent=self.master,
        )
        if answer is None:
            return None
        return out if answer else alt

    def _start_convert(self):
        if self._worker_thread and self._worker_thread.is_alive():
            return

        pdf = self.pdf_path.get().strip().strip('"')
        if not pdf:
            messagebox.showerror(tr('Ошибка'), tr('Выберите файл.'))
            return

        pdf_path = Path(pdf)
        if not pdf_path.is_file() or not pdf2md.is_supported_file(pdf_path):
            messagebox.showerror(
                tr('Ошибка'), tr('Нужен существующий файл: ') + ", ".join(pdf2md.SUPPORTED_SUFFIXES)
            )
            return
        is_pdf = pdf_path.suffix.lower() == ".pdf"
        if not is_pdf:
            self._refresh_estimate()  # make sure the sheet list belongs to this file

        options = UiOptions(
            engine=self.engine.get(),
            sheet=self.sheet.get() if pdf_path.suffix.lower() in EXCEL_SUFFIXES else "",
            page_markers=bool(self.page_markers.get()),
            page_separators=bool(self.page_separators.get()),
            save_images=bool(self.save_images.get()),
            strip_headers=bool(self.strip_headers.get()),
        )
        if is_pdf and options.engine == "marker" and self.marker_install_index is not None:
            if marker_install.installed_marker_command() is None:
                self._start_marker_install()
                return
        out_md = self._choose_output(pdf_path)
        if out_md is None:
            return

        if not is_pdf:
            self._set_busy(True)
            self.progress_status.set(tr('0% · запуск…'))
            self._log("INFO", tr("Файл: {path}").format(path=pdf_path))
            if options.sheet:
                self._log("INFO", tr("Лист: {sheet}").format(sheet=options.sheet))
            self._log("INFO", tr("Конвертация…"))
            self._worker_thread = threading.Thread(
                target=self._convert_worker, args=(pdf_path, out_md, options), daemon=True
            )
            self._worker_thread.start()
            return

        if self._should_warn_about_marker_download(options):
            ok = messagebox.askokcancel(
                tr('Первый запуск Marker'),
                tr('Первый запуск Marker может скачать несколько ГБ моделей. Они будут сохранены в portable-папке models и потом будут переноситься вместе с программой.\n\nПродолжить?'),
            )
            if not ok:
                return
            self._remember_marker_download_notice()

        try:
            info = get_pdf_info(pdf_path)
            estimate_text = format_estimate_range(options.engine, info)
        except Exception as e:
            messagebox.showerror(tr('Ошибка'), tr('Не удалось прочитать PDF: {v0}').format(v0=e))
            return

        self._set_busy(True)
        self.progress_status.set(tr('0% · запуск…'))
        self._log("INFO", tr("Файл: {path}").format(path=pdf_path))
        self._log("INFO", tr("Движок: {engine}").format(engine=options.engine))
        self._log("INFO", tr("Оценка: {estimate}").format(estimate=estimate_text))
        self._log("INFO", tr("Конвертация…"))

        self._worker_thread = threading.Thread(
            target=self._convert_worker,
            args=(pdf_path, out_md, options),
            daemon=True,
        )
        self._worker_thread.start()

    def _should_warn_about_marker_download(self, options: UiOptions) -> bool:
        root = os.environ.get("PDF2MD_PORTABLE_ROOT")
        if not root:
            return False
        root_path = Path(root)
        notice_seen = root_path / "models" / ".marker_download_notice_seen"
        if notice_seen.exists():
            return False
        resolved = pdf2md.resolve_engine(options.engine, get_pdf_info(Path(self.pdf_path.get())))
        if resolved != "marker":
            return False
        datalab_models = root_path / "models" / "datalab" / "models"
        return not self._has_marker_models(datalab_models)

    def _has_marker_models(self, models_dir: Path) -> bool:
        if not models_dir.exists():
            return False
        return any(models_dir.rglob("manifest.json"))

    def _remember_marker_download_notice(self) -> None:
        root = os.environ.get("PDF2MD_PORTABLE_ROOT")
        if not root:
            return
        notice = Path(root) / "models" / ".marker_download_notice_seen"
        try:
            notice.parent.mkdir(parents=True, exist_ok=True)
            notice.write_text("shown\n", encoding="utf-8")
        except Exception:
            pass

    def _convert_worker(self, pdf_path: Path, out_md: Path, options: UiOptions):
        def on_progress(update: ProgressUpdate) -> None:
            self._messages.put(("PROGRESS", update))

        try:
            ok, used = pdf2md.convert_file(
                pdf_path,
                out_md,
                engine=options.engine,
                sheet=options.sheet or None,
                page_markers=options.page_markers,
                page_separators=options.page_separators,
                save_images=options.save_images,
                strip_headers=options.strip_headers,
                progress_callback=on_progress,
                cancel_event=self._cancel_event,
            )
            if not ok:
                self._messages.put(
                    (
                        "ERROR",
                        tr('Marker не смог завершить конвертацию. Подробности выше в логе.'),
                    )
                )
                return
            self._messages.put(("DONE", (out_md.resolve(), used)))
        except pdf2md.ConversionCancelled:
            self._messages.put(("CANCELLED", tr('Конвертация отменена — файл не записан')))
        except (pdf2md.PdfOpenError, office2md.OfficeError) as e:  # already readable
            self._messages.put(("ERROR", str(e)))
        except OSError as e:
            self._messages.put(("ERROR", pdf2md.describe_write_error(e)))
        except Exception as e:
            self._messages.put(("ERROR", f"{type(e).__name__}: {e}"))
            traceback.print_exc()  # goes to logs\pdf2md.log under PDF2MD.exe

    def _poll_messages(self):
        try:
            while True:
                level, payload = self._messages.get_nowait()
                if level == "MARKER_INSTALLED":
                    self._build_menu()
                    self._refresh_estimate()
                    level = "DONE"
                if level == "PROGRESS" and isinstance(payload, ProgressUpdate):
                    self.progress.config(value=payload.percent)
                    self.progress_status.set(self._format_progress_status(payload))
                    if payload.message.startswith(("ERROR:", "WARN:", "INFO:")):
                        level, _, text = payload.message.partition(":")
                        self._log(level, text.strip())
                    continue

                if level == "DONE" and isinstance(payload, tuple):
                    self._last_result, used = payload
                    msg = tr('Готово: {v0} ({v1})').format(v0=self._last_result, v1=used)
                else:
                    msg = str(payload)
                self._log("WARN" if level == "CANCELLED" else level, msg)
                if level in ("DONE", "ERROR", "CANCELLED"):
                    if level == "DONE":
                        self.progress.config(value=100)
                        self.progress_status.set(f"100% · {msg}")
                        self._show_result()
                    elif level == "CANCELLED":
                        self.progress.config(value=0)
                        self.progress_status.set(msg)
                    else:
                        self.progress_status.set(tr('Ошибка · {v0}').format(v0=msg))
                    self._set_busy(False)
        except queue.Empty:
            pass
        self._poll_after_id = self.after(100, self._poll_messages)

    def destroy(self):
        if self._settings_after_id is not None:
            self.after_cancel(self._settings_after_id)
            self._settings_after_id = None
        # A pending poll would fire into a deleted Tcl command after the window closes.
        if self._poll_after_id is not None:
            try:
                self.after_cancel(self._poll_after_id)
            except Exception:
                pass
            self._poll_after_id = None
        super().destroy()


LOG_MAX_BYTES = 2 * 1024 * 1024


def log_dir() -> Path:
    """`<portable>\\logs` for PDF2MD.exe, else `%LOCALAPPDATA%\\PDF2MD\\logs`."""
    portable = os.environ.get("PDF2MD_PORTABLE_ROOT")
    if portable and not marker_install.installed_build():
        return Path(portable) / "logs"
    return marker_install.user_data_dir() / "logs"


def install_error_logging(root: Tk, app: "App | None" = None):
    """Write errors and the conversion log to logs\\pdf2md.log.

    Under pythonw (PDF2MD.exe) there is no console: without this an exception in a Tk callback
    or a worker thread disappears silently. Returns the open log file, or None.
    """
    try:
        directory = log_dir()
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / "pdf2md.log"
        if path.exists() and path.stat().st_size > LOG_MAX_BYTES:
            path.replace(directory / "pdf2md.log.1")
        log = path.open("a", encoding="utf-8", buffering=1)
    except OSError:
        return None
    log.write(f"\n{time.strftime('%Y-%m-%d %H:%M:%S')} === PDF2MD started (Python {sys.version.split()[0]}) ===\n")
    if sys.stderr is None:  # pythonw
        sys.stderr = log
    if sys.stdout is None:
        sys.stdout = log

    def report_callback_exception(exc, value, tb):
        log.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} [CRASH] UI callback\n")
        traceback.print_exception(exc, value, tb, file=log)
        messagebox.showerror(
            tr('PDF2MD — внутренняя ошибка'),
            tr('{v0}: {v1}\n\nПодробности записаны в журнал:\n{v2}').format(v0=exc.__name__, v1=value, v2=path),
            parent=root,
        )

    def thread_excepthook(args):
        log.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} [CRASH] thread {args.thread.name if args.thread else '?'}\n")
        traceback.print_exception(args.exc_type, args.exc_value, args.exc_traceback, file=log)

    root.report_callback_exception = report_callback_exception
    threading.excepthook = thread_excepthook
    if app is not None:
        app.log_file = log
    return log


def make_root() -> Tk:
    """Tk root with drag & drop when tkinterdnd2 is available (plain Tk otherwise)."""
    try:
        from tkinterdnd2 import TkinterDnD
    except ImportError:
        return Tk()

    class DropRoot(Tk, TkinterDnD.DnDWrapper):
        pass

    root = DropRoot()
    try:
        # Load into this interpreter. A failed TkinterDnD.Tk() constructor leaves
        # a hidden default root, causing PhotoImage to attach to the wrong Tk.
        root.TkdndVersion = TkinterDnD._require(root)
    except Exception:
        pass  # the same root remains usable without drag and drop
    return root


def main():
    enable_dpi_awareness()  # before the first window: crisp text at 125 % / 150 %
    root = make_root()
    app = App(root)
    install_error_logging(root, app)
    root.mainloop()


if __name__ == "__main__":
    main()
