"""Windows XP "Luna" look for the PDF2MD window (Tkinter / ttk).

ttk elements (buttons, entry, combobox, checkbox) get images drawn here with Pillow at the
screen's DPI scale (antialiased by 4x supersampling), so they stay crisp at 125 % / 150 %.
Things ttk cannot draw — the blue header band and the XP progress bar made of green blocks —
are small Canvas widgets (`HeaderBand`, `XPProgress`).

Palette taken from Windows XP Luna (blue): window #ECE9D8, button border #003C74, field
border #7F9DB9, tooltip #FFFFE1, progress chunks green.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from tkinter import Canvas, Misc, ttk

from PIL import Image, ImageDraw, ImageTk

WINDOW_BG = "#ECE9D8"
TEXT = "#000000"
MUTED = "#6D6A5D"
LINK = "#0046D5"
FIELD_BORDER = "#7F9DB9"
TOOLTIP_BG = "#FFFFE1"
FONT_FAMILY = "Tahoma"

_SS = 4  # supersampling factor for antialiased shapes


def enable_dpi_awareness() -> bool:
    """Crisp text instead of Windows' blurry bitmap scaling. Must run before Tk() is created."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system-DPI aware
        return True
    except Exception:
        try:
            import ctypes

            return bool(ctypes.windll.user32.SetProcessDPIAware())
        except Exception:
            return False


def _hex(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    return tuple(int(color[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def _gradient(stops: list[tuple[float, str]], t: float) -> tuple[int, int, int]:
    for (p0, c0), (p1, c1) in zip(stops, stops[1:]):
        if p0 <= t <= p1:
            k = 0.0 if p1 == p0 else (t - p0) / (p1 - p0)
            a, b = _hex(c0), _hex(c1)
            return tuple(round(a[i] + (b[i] - a[i]) * k) for i in range(3))  # type: ignore[return-value]
    return _hex(stops[-1][1])


def _rounded_box(
    w: int,
    h: int,
    radius: float,
    outline: str | None,
    stops: list[tuple[float, str]],
    *,
    inner_top: str | None = None,
    inner_bottom: str | None = None,
) -> Image.Image:
    """Rounded rectangle with a vertical gradient, 1 px border and optional 2 px inner glow lines."""
    W, H, R = w * _SS, h * _SS, radius * _SS
    fill = Image.new("RGBA", (W, H))
    px = fill.load()
    for y in range(H):
        rgb = _gradient(stops, y / max(1, H - 1))
        for x in range(W):
            px[x, y] = (*rgb, 255)
    mask = Image.new("L", (W, H), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, W - 1, H - 1), R, fill=255)
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    img.paste(fill, (0, 0), mask)
    d = ImageDraw.Draw(img)
    if inner_top:
        d.rounded_rectangle((_SS, _SS, W - 1 - _SS, _SS * 3), R / 2, fill=inner_top)
    if inner_bottom:
        d.rounded_rectangle((_SS, H - 1 - _SS * 3, W - 1 - _SS, H - 1 - _SS), R / 2, fill=inner_bottom)
    if outline:
        d.rounded_rectangle((0, 0, W - 1, H - 1), R, outline=outline, width=_SS)
    return img.resize((w, h), Image.LANCZOS)


def _check_mark(size: int, color: str) -> Image.Image:
    S = size * _SS
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(img).line(
        [(S * 0.22, S * 0.52), (S * 0.42, S * 0.72), (S * 0.78, S * 0.28)], fill=color, width=max(_SS * 2, S // 7), joint="curve"
    )
    return img.resize((size, size), Image.LANCZOS)


def _chevron(w: int, h: int, color: str) -> Image.Image:
    W, H = w * _SS, h * _SS
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(img).line(
        [(W * 0.3, H * 0.4), (W * 0.5, H * 0.6), (W * 0.7, H * 0.4)], fill=color, width=max(_SS * 2, W // 8), joint="curve"
    )
    return img.resize((w, h), Image.LANCZOS)


@dataclass
class LunaTheme:
    scale: float
    images: dict[str, ImageTk.PhotoImage] = field(default_factory=dict)

    def px(self, value: float) -> int:
        return max(1, round(value * self.scale))

    def font(self, size: int = 8, *weight: str) -> tuple:
        return (FONT_FAMILY, size, *weight)


def _photo(theme: LunaTheme, name: str, img: Image.Image) -> ImageTk.PhotoImage:
    photo = ImageTk.PhotoImage(img)
    theme.images[name] = photo  # keep a reference: Tk drops unreferenced images
    return photo


# button gradients (top -> bottom)
_BTN = [(0.0, "#FFFFFF"), (0.55, "#F4F3EE"), (1.0, "#D6D0C5")]
_BTN_PRESSED = [(0.0, "#D9D6CC"), (0.4, "#E8E6DE"), (1.0, "#F3F2EE")]
_BTN_DISABLED = [(0.0, "#F7F6F1"), (1.0, "#F1EFE6")]
_GREEN = [(0.0, "#8FDB6B"), (0.45, "#4DB22E"), (1.0, "#2F8F1C")]
_GREEN_HOVER = [(0.0, "#A6E886"), (0.45, "#5FC23F"), (1.0, "#379E22")]
_GREEN_PRESSED = [(0.0, "#2C8519"), (0.5, "#3E9F27"), (1.0, "#56B937")]
_COMBO_BTN = [(0.0, "#E3ECFE"), (0.5, "#C7D8FB"), (1.0, "#AFC6F6")]
_COMBO_BTN_HOVER = [(0.0, "#F0F5FF"), (0.5, "#D9E5FD"), (1.0, "#C2D5FA")]
_CHECK_BOX = [(0.0, "#DCDBD5"), (1.0, "#FFFFFF")]


def apply_luna(root: Misc) -> LunaTheme:
    """Install the Luna style on `root` (ttk styles prefixed "Luna.")."""
    scale = root.winfo_fpixels("1i") / 96.0
    theme = LunaTheme(scale=scale)
    style = ttk.Style(root)
    style.theme_use("clam")  # fully image/colour driven, no native look leaking in
    root.configure(background=WINDOW_BG)
    root.option_add("*Font", theme.font(8))
    root.option_add("*TCombobox*Listbox.font", theme.font(8))
    root.option_add("*TCombobox*Listbox.background", "#FFFFFF")
    root.option_add("*TCombobox*Listbox.selectBackground", "#316AC5")
    root.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")

    style.configure(".", background=WINDOW_BG, foreground=TEXT, font=theme.font(8))
    style.configure("TFrame", background=WINDOW_BG)
    style.configure("TLabel", background=WINDOW_BG, foreground=TEXT)
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("Small.TLabel", foreground=MUTED, font=theme.font(7))
    style.configure("Bold.TLabel", font=theme.font(8, "bold"))
    style.configure("Link.TLabel", foreground=LINK)
    style.configure("Tooltip.TFrame", background=TOOLTIP_BG, borderwidth=1, relief="solid")
    style.configure("Tooltip.TLabel", background=TOOLTIP_BG, foreground=TEXT, font=theme.font(8))

    # --- buttons -------------------------------------------------------------------------
    bw, bh, r = theme.px(24), theme.px(23), theme.px(3)
    border = theme.px(5)

    def button_images(prefix: str, normal, hover, pressed, outline: str, hover_glow=True):
        imgs = {
            "normal": _rounded_box(bw, bh, r, outline, normal),
            "hover": _rounded_box(
                bw, bh, r, outline, hover,
                inner_top="#FFF0CF" if hover_glow else None, inner_bottom="#F8B330" if hover_glow else None,
            ),
            "pressed": _rounded_box(bw, bh, r, outline, pressed),
            "disabled": _rounded_box(bw, bh, r, "#C9C7BA", _BTN_DISABLED),
        }
        return {k: _photo(theme, f"{prefix}.{k}", v) for k, v in imgs.items()}

    btn = button_images("btn", _BTN, _BTN, _BTN_PRESSED, "#003C74")
    style.element_create(
        "Luna.Button.button", "image", btn["normal"],
        ("disabled", btn["disabled"]), ("pressed", btn["pressed"]), ("active", btn["hover"]),
        border=border, sticky="nsew",
    )
    style.layout("Luna.TButton", [("Luna.Button.button", {"sticky": "nsew", "children": [
        ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})])
    style.configure("Luna.TButton", padding=(theme.px(10), theme.px(3)), anchor="center", foreground=TEXT)
    style.map("Luna.TButton", foreground=[("disabled", "#ACA899")])

    green = button_images("green", _GREEN, _GREEN_HOVER, _GREEN_PRESSED, "#1E6E12", hover_glow=False)
    style.element_create(
        "Luna.Primary.button", "image", green["normal"],
        ("disabled", green["disabled"]), ("pressed", green["pressed"]), ("active", green["hover"]),
        border=border, sticky="nsew",
    )
    style.layout("Luna.Primary.TButton", [("Luna.Primary.button", {"sticky": "nsew", "children": [
        ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})])
    style.configure(
        "Luna.Primary.TButton", padding=(theme.px(14), theme.px(4)), anchor="center",
        foreground="#FFFFFF", font=theme.font(8, "bold"),
    )
    style.map("Luna.Primary.TButton", foreground=[("disabled", "#ACA899")])

    # --- entry / combobox fields ------------------------------------------------------------
    fw, fh = theme.px(20), theme.px(21)
    field_img = Image.new("RGBA", (fw, fh), "#FFFFFF")
    ImageDraw.Draw(field_img).rectangle((0, 0, fw - 1, fh - 1), outline=FIELD_BORDER, width=max(1, round(scale)))
    field_dis = Image.new("RGBA", (fw, fh), "#F5F4EA")
    ImageDraw.Draw(field_dis).rectangle((0, 0, fw - 1, fh - 1), outline="#C9C7BA", width=max(1, round(scale)))
    field_n, field_d = _photo(theme, "field", field_img), _photo(theme, "field.dis", field_dis)
    fb = theme.px(3)
    style.element_create("Luna.Entry.field", "image", field_n, ("disabled", field_d), border=fb, sticky="nsew")
    style.layout("Luna.TEntry", [("Luna.Entry.field", {"sticky": "nsew", "children": [
        ("Entry.padding", {"sticky": "nsew", "children": [("Entry.textarea", {"sticky": "nsew"})]})]})])
    style.configure("Luna.TEntry", padding=(theme.px(3), theme.px(2)), foreground=TEXT)
    style.map("Luna.TEntry", foreground=[("disabled", "#ACA899")])

    aw, ah = theme.px(17), theme.px(19)
    arrow = _chevron(aw, ah, "#4D6185")

    def arrow_btn(stops):
        img = _rounded_box(aw, ah, theme.px(2), "#B4C8F6", stops)
        img.alpha_composite(arrow)
        return img

    arr_n = _photo(theme, "combo.arrow", arrow_btn(_COMBO_BTN))
    arr_h = _photo(theme, "combo.arrow.hover", arrow_btn(_COMBO_BTN_HOVER))
    arr_d_img = _rounded_box(aw, ah, theme.px(2), "#D7D5CB", _BTN_DISABLED)
    arr_d_img.alpha_composite(_chevron(aw, ah, "#B8B5A8"))
    arr_d = _photo(theme, "combo.arrow.dis", arr_d_img)
    style.element_create("Luna.Combobox.field", "image", field_n, ("disabled", field_d), border=fb, sticky="nsew")
    style.element_create(
        "Luna.Combobox.downarrow", "image", arr_n, ("disabled", arr_d), ("active", arr_h), ("pressed", arr_h), sticky=""
    )
    style.layout("Luna.TCombobox", [("Luna.Combobox.field", {"sticky": "nsew", "children": [
        ("Luna.Combobox.downarrow", {"side": "right", "sticky": "ns"}),
        ("Combobox.padding", {"expand": "1", "sticky": "nsew", "children": [("Combobox.textarea", {"sticky": "nsew"})]})]})])
    style.configure("Luna.TCombobox", padding=(theme.px(3), theme.px(1), theme.px(1), theme.px(1)), foreground=TEXT,
                    arrowsize=aw)
    style.map(
        "Luna.TCombobox",
        fieldbackground=[("readonly", "#FFFFFF")],
        foreground=[("disabled", "#ACA899")],
        selectbackground=[("readonly", "#FFFFFF")],
        selectforeground=[("readonly", TEXT)],
    )

    # --- checkbox ---------------------------------------------------------------------------
    cs = theme.px(13)
    box = _rounded_box(cs, cs, 0, "#1C5180", _CHECK_BOX)
    box_hover = _rounded_box(cs, cs, 0, "#1C5180", _CHECK_BOX, inner_top="#FFE4A3", inner_bottom="#F8B330")
    box_dis = _rounded_box(cs, cs, 0, "#CAC8BB", [(0.0, "#F5F4EA"), (1.0, "#F5F4EA")])
    mark = _check_mark(cs, "#21A121")
    mark_dis = _check_mark(cs, "#CAC8BB")

    gap = theme.px(5)  # space between the box and its label, baked into the image

    def with_mark(img, m):
        out = img.copy()
        if m is not None:
            out.alpha_composite(m)
        padded = Image.new("RGBA", (cs + gap, cs), (0, 0, 0, 0))
        padded.paste(out, (0, 0))
        return padded

    c_off, c_on = _photo(theme, "chk.off", with_mark(box, None)), _photo(theme, "chk.on", with_mark(box, mark))
    c_off_h = _photo(theme, "chk.off.h", with_mark(box_hover, None))
    c_on_h = _photo(theme, "chk.on.h", with_mark(box_hover, mark))
    c_off_d = _photo(theme, "chk.off.d", with_mark(box_dis, None))
    c_on_d = _photo(theme, "chk.on.d", with_mark(box_dis, mark_dis))
    style.element_create(
        "Luna.Checkbutton.indicator", "image", c_off,
        ("disabled", "selected", c_on_d), ("disabled", c_off_d),
        ("active", "selected", c_on_h), ("active", c_off_h), ("selected", c_on),
        sticky="",
    )
    style.layout("Luna.TCheckbutton", [("Checkbutton.padding", {"sticky": "nsew", "children": [
        ("Luna.Checkbutton.indicator", {"side": "left", "sticky": ""}),
        ("Checkbutton.label", {"side": "left", "sticky": "nsew"})]})])
    style.configure("Luna.TCheckbutton", background=WINDOW_BG, foreground=TEXT, padding=(0, theme.px(2)))
    style.map("Luna.TCheckbutton", foreground=[("disabled", "#ACA899")], background=[("active", WINDOW_BG)])

    style.configure("Luna.TSeparator", background="#D0CBB8")
    return theme


class HeaderBand(Canvas):
    """Blue XP gradient band with the program title (like an XP wizard / task pane header)."""

    def __init__(self, master, theme: LunaTheme, title: str, subtitle: str):
        height = theme.px(54)
        super().__init__(master, height=height, highlightthickness=0, borderwidth=0, background="#0A4DC7")
        self.theme, self.title, self.subtitle = theme, title, subtitle
        self.bind("<Configure>", self._draw)

    def _draw(self, _evt=None):
        w, h = max(1, self.winfo_width()), int(self.cget("height"))
        img = Image.new("RGB", (w, h))
        stops = [(0.0, "#3D82F0"), (0.08, "#2563DC"), (0.85, "#0C47C1"), (1.0, "#0A3DA8")]
        d = ImageDraw.Draw(img)
        for y in range(h):
            d.line([(0, y), (w, y)], fill=_gradient(stops, y / max(1, h - 1)))
        d.line([(0, h - 1), (w, h - 1)], fill="#F29A1D")  # XP orange accent line
        self._img = ImageTk.PhotoImage(img)
        self.delete("all")
        self.create_image(0, 0, image=self._img, anchor="nw")
        x = self.theme.px(14)
        self.create_text(x, self.theme.px(19), text=self.title, anchor="w", fill="#FFFFFF",
                         font=self.theme.font(13, "bold"))
        self.create_text(x, self.theme.px(38), text=self.subtitle, anchor="w", fill="#D9E6FF",
                         font=self.theme.font(8))
        icon_size = self.theme.px(40)
        if getattr(self, "_icon", None) is None or self._icon.width() != icon_size:
            self._icon = ImageTk.PhotoImage(make_app_icon(icon_size))
        self.create_image(w - self.theme.px(12), h // 2, image=self._icon, anchor="e")


class XPProgress(Canvas):
    """XP progress bar: white rounded trough, green chunks with gaps."""

    def __init__(self, master, theme: LunaTheme):
        super().__init__(master, height=theme.px(17), highlightthickness=0, borderwidth=0, background=WINDOW_BG)
        self.theme = theme
        self.value = 0.0
        self._chunk = None
        self.bind("<Configure>", self._draw)

    def set(self, percent: float) -> None:
        self.value = max(0.0, min(100.0, float(percent)))
        self._draw()

    def config(self, **kw):  # ttk.Progressbar compatibility: .config(value=…)
        if "value" in kw:
            self.set(kw.pop("value"))
        if kw:
            super().config(**kw)

    configure = config

    def cget(self, key):
        return self.value if key == "value" else super().cget(key)

    def _draw(self, _evt=None):
        t = self.theme
        w, h = max(4, self.winfo_width()), int(super().cget("height"))
        trough = _rounded_box(w, h, t.px(3), "#686868", [(0.0, "#FFFFFF"), (1.0, "#FFFFFF")])
        self._trough = ImageTk.PhotoImage(trough)
        self.delete("all")
        self.create_image(0, 0, image=self._trough, anchor="nw")
        cw, gap, pad = t.px(8), t.px(2), t.px(3)
        if self._chunk is None or self._chunk.height() != h - 2 * pad:
            chunk = Image.new("RGB", (cw, h - 2 * pad))
            cd = ImageDraw.Draw(chunk)
            stops = [(0.0, "#C9F2C0"), (0.35, "#3DD045"), (0.65, "#2BB534"), (1.0, "#8EE28F")]
            for y in range(chunk.height):
                cd.line([(0, y), (cw, y)], fill=_gradient(stops, y / max(1, chunk.height - 1)))
            self._chunk = ImageTk.PhotoImage(chunk)
        inner = w - 2 * pad
        n = int((inner + gap) * self.value / 100 // (cw + gap))
        for i in range(n):
            self.create_image(pad + i * (cw + gap), pad, image=self._chunk, anchor="nw")


def make_app_icon(size: int = 256) -> Image.Image:
    """Program icon: a page with a folded corner, blue XP band and the Markdown "M↓" mark."""
    S = size * _SS
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = S * 0.12
    fold = S * 0.22
    page = [(m, m), (S - m - fold, m), (S - m, m + fold), (S - m, S - m), (m, S - m)]
    d.polygon([(x + S * 0.025, y + S * 0.03) for x, y in page], fill=(0, 0, 0, 70))  # shadow
    d.polygon(page, fill="#FFFFFF", outline="#003C74")
    d.line(page + [page[0]], fill="#003C74", width=max(_SS, S // 64))
    d.polygon([(S - m - fold, m), (S - m - fold, m + fold), (S - m, m + fold)], fill="#C9D7F2", outline="#003C74")
    band_top, band_bottom = S * 0.50, S * 0.80
    for y in range(int(band_top), int(band_bottom)):
        rgb = _gradient([(0.0, "#3D82F0"), (1.0, "#0C47C1")], (y - band_top) / (band_bottom - band_top))
        d.line([(m + _SS * 2, y), (S - m - _SS * 2, y)], fill=rgb)
    # "M" and a down arrow in white on the band
    cx, cy, hgt = S * 0.40, (band_top + band_bottom) / 2, (band_bottom - band_top) * 0.62
    lw = max(_SS * 2, int(S * 0.045))
    mx0, mx1 = cx - hgt * 0.55, cx + hgt * 0.55
    d.line([(mx0, cy + hgt / 2), (mx0, cy - hgt / 2), (cx, cy + hgt * 0.1), (mx1, cy - hgt / 2), (mx1, cy + hgt / 2)],
           fill="#FFFFFF", width=lw, joint="curve")
    ax = S * 0.68
    d.line([(ax, cy - hgt / 2), (ax, cy + hgt / 2)], fill="#FFFFFF", width=lw)
    d.polygon([(ax - hgt * 0.32, cy + hgt * 0.08), (ax + hgt * 0.32, cy + hgt * 0.08), (ax, cy + hgt * 0.55)], fill="#FFFFFF")
    # text lines above the band
    for i, frac in enumerate((0.26, 0.34, 0.42)):
        right = S - m - (fold if i == 0 else 0) - S * 0.08
        d.line([(m + S * 0.08, S * frac), (right, S * frac)], fill="#9DB4DD", width=max(_SS, S // 48))
    return img.resize((size, size), Image.LANCZOS)
