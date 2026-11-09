"""GUI smoke test without mouse clicks (no computer use needed).

Opens the real Tkinter window, fills in the file path (PDF/DOCX/XLSX/CSV/TXT), runs the conversion the
same way the "Конвертировать" button does, waits for DONE/ERROR and prints the
status line and log. Optionally saves a screenshot of the window (Windows /
macOS, needs Pillow — bundled in the portable runtime).

    python tests/ui_smoke.py samples\\file.pdf
    python tests/ui_smoke.py samples\\file.pdf --engine structured --screenshot out\\ui.png

Exit code: 0 = converted, 1 = error/timeout, 2 = bad arguments.
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def _screen_scale(root) -> float:
    """Windows display scaling (125 %, 150 %…) for a DPI-unaware Tk window.

    The app is not DPI-aware, so Tk reports logical coordinates, while ImageGrab
    captures physical pixels — without this the screenshot is shifted and cropped.
    """
    if sys.platform != "win32":
        return 1.0
    try:
        import ctypes

        hwnd = int(root.wm_frame(), 16)
        user32 = ctypes.windll.user32
        user32.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        user32.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.GetDpiForWindow.restype = ctypes.c_uint
        # Inside a DPI-unaware thread every DPI query is virtualised to 96;
        # switch this thread to per-monitor-aware just for the query.
        old = user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2
        try:
            monitor = user32.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
            dpi_x, dpi_y = ctypes.c_uint(), ctypes.c_uint()
            ctypes.windll.shcore.GetDpiForMonitor(monitor, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y))
        finally:
            if old:
                user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(old))
        return dpi_x.value / 96.0 if dpi_x.value else 1.0
    except Exception:
        return 1.0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="PDF2MD GUI smoke test")
    ap.add_argument("pdf", metavar="file", type=Path, help="PDF, DOCX, XLSX, CSV, TSV or TXT")
    ap.add_argument("--engine", default="auto")
    ap.add_argument("--sheet", default=None, help="Excel sheet name (default: active)")
    ap.add_argument("--timeout", type=float, default=600.0, help="seconds")
    ap.add_argument("--screenshot", type=Path, default=None)
    args = ap.parse_args(argv)
    if not args.pdf.exists():
        print(f"ERROR: no such file: {args.pdf}", file=sys.stderr)
        return 2

    from tkinter import Tk
    import ui_dialogs as messagebox

    import pdf2md_ui

    # Never block on a modal dialog (e.g. the first-run Marker download notice).
    messagebox.askokcancel = lambda *a, **k: True  # type: ignore[assignment]
    messagebox.askyesnocancel = lambda *a, **k: True  # type: ignore[assignment]  # overwrite .md
    messagebox.showerror = lambda title, msg, *a, **k: print(f"DIALOG {title}: {msg}")  # type: ignore[assignment]

    # Same start as the real program: DPI-aware, drag & drop root.
    dpi_aware = pdf2md_ui.enable_dpi_awareness()
    root = pdf2md_ui.make_root()
    app = pdf2md_ui.App(root)
    root.geometry("+40+20")
    app.engine.set(args.engine)
    app.pdf_path.set(str(args.pdf.resolve()))
    app._refresh_estimate()
    if args.sheet:
        app.sheet.set(args.sheet)
    print("ESTIMATE:", app.estimate_label.cget("text"))

    app._start_convert()
    started = time.monotonic()
    ok = False
    if app._worker_thread is None:  # rejected before starting (bad file, dialog)
        print("STATUS:", app.progress_status.get())
        root.destroy()
        return 1
    while time.monotonic() - started < args.timeout:
        root.update()
        time.sleep(0.05)
        worker = app._worker_thread
        if worker is not None and not worker.is_alive() and str(app.btn_convert.cget("state")) == "normal":
            for _ in range(10):  # let the last queue messages render
                root.update()
                time.sleep(0.05)
            ok = app._last_result is not None and app._last_result.is_file()
            break
    else:
        print("ERROR: timeout", file=sys.stderr)

    print("STATUS:", app.progress_status.get())
    print("LOG:\n" + app.log.get("1.0", "end").strip())

    if args.screenshot:
        try:
            from PIL import ImageGrab

            root.attributes("-topmost", True)
            root.lift()
            root.update()
            time.sleep(0.3)
            x, y = root.winfo_rootx(), root.winfo_rooty()
            w, h = root.winfo_width(), root.winfo_height()
            # DPI-aware window: Tk already reports physical pixels
            scale = 1.0 if dpi_aware else _screen_scale(root)
            bbox = tuple(round(v * scale) for v in (x, y, x + w, y + h))
            args.screenshot.parent.mkdir(parents=True, exist_ok=True)
            ImageGrab.grab(bbox=bbox).save(args.screenshot)
            print("SCREENSHOT:", args.screenshot)
        except Exception as e:  # screenshots are best effort
            print(f"WARN: screenshot failed: {e}", file=sys.stderr)

    root.destroy()
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
