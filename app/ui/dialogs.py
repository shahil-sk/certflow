"""
All Toplevel dialog windows — dark-themed.
"""
import tkinter as tk
from tkinter import colorchooser

from PIL import Image, ImageTk

from app.constants import C
from app.helpers import cmyk_to_hex
from app.ui.widgets import flat_button, label, hsep

_LANCZOS = Image.Resampling.LANCZOS


def _dark_win(parent, title, w=None, h=None, resizable=(True, True)):
    """Create a consistently styled Toplevel."""
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.grab_set()
    win.configure(bg=C["surface"])
    win.resizable(*resizable)
    if w and h:
        win.geometry(f"{w}x{h}")
    return win


def show_preview(parent: tk.Tk, img: Image.Image) -> None:
    """Resizable certificate preview window."""
    pw = max(parent.winfo_width() - 80, 700)
    ph = max(parent.winfo_height() - 80, 480)
    win = _dark_win(parent, "Preview", pw, ph)

    # toolbar
    bar = tk.Frame(win, bg=C["nav"], height=40)
    bar.pack(fill="x")
    bar.pack_propagate(False)
    tk.Label(bar, text="Preview — first record",
             font=("Segoe UI", 9), fg=C["subtext"],
             bg=C["nav"]).pack(side="left", padx=14)
    tk.Button(
        bar, text="✕  Close",
        command=win.destroy,
        bg=C["nav"], fg=C["subtext"],
        relief="flat", bd=0, cursor="hand2",
        font=("Segoe UI", 8),
        activebackground=C["surface"],
        activeforeground=C["danger"],
        padx=12, pady=4,
    ).pack(side="right", padx=6)

    # image area
    img_frame = tk.Frame(win, bg=C["canvas_bg"])
    img_frame.pack(fill="both", expand=True, padx=12, pady=12)

    iw, ih  = img.size
    ratio   = min((pw - 40) / iw, (ph - 80) / ih, 1.0)
    photo   = ImageTk.PhotoImage(
        img.resize((max(int(iw * ratio), 1),
                    max(int(ih * ratio), 1)), _LANCZOS))
    lbl = tk.Label(img_frame, image=photo, bg=C["canvas_bg"], bd=0)
    lbl.image = photo
    lbl.pack(expand=True)

    win.bind("<Escape>", lambda e: win.destroy())
    win.lift()
    win.focus_set()


def pick_color_rgb(parent: tk.Tk, field: str, font_settings: dict) -> None:
    cur    = font_settings[field]["color"].get()
    init   = cur if cur.startswith("#") else "#000000"
    chosen = colorchooser.askcolor(
        title=f"Color for ‘{field}’", initialcolor=init, parent=parent)
    if chosen[1]:
        font_settings[field]["color"].set(chosen[1])


def pick_color_cmyk(parent: tk.Tk, field: str, font_settings: dict) -> None:
    cur = font_settings[field]["color"].get()
    try:
        vals = list(map(float, cur[5:-1].split(",")))
    except Exception:
        vals = [0.0, 0.0, 0.0, 0.0]

    win = _dark_win(parent, f"CMYK — {field}",
                    320, 260, resizable=(False, False))

    cvars = {ch: tk.DoubleVar(value=v)
             for ch, v in zip("CMYK", vals)}

    # preview swatch
    preview = tk.Label(win, width=24, height=3,
                       bg=C["surface2"], relief="flat")
    preview.grid(row=0, column=0, columnspan=2,
                 padx=20, pady=(16, 10), sticky="ew")

    def _refresh(*_):
        cmyk_str = "cmyk({:.2f},{:.2f},{:.2f},{:.2f})".format(
            cvars["C"].get(), cvars["M"].get(),
            cvars["Y"].get(), cvars["K"].get())
        font_settings[field]["color"].set(cmyk_str)
        try:
            preview.config(bg=cmyk_to_hex(cmyk_str))
        except Exception:
            pass

    for row, (ch_name, ch) in enumerate(
            zip(("Cyan", "Magenta", "Yellow", "Black"), "CMYK"), start=1):
        tk.Label(
            win, text=ch_name,
            font=("Segoe UI", 8), fg=C["subtext"],
            bg=C["surface"], anchor="e", width=7,
        ).grid(row=row, column=0, padx=(16, 6), pady=4, sticky="e")
        tk.Scale(
            win, from_=0, to=1, resolution=0.01,
            variable=cvars[ch], command=_refresh,
            orient="horizontal", length=210,
            bg=C["surface"], fg=C["text"],
            troughcolor=C["surface3"],
            highlightthickness=0, relief="flat",
            showvalue=False,
        ).grid(row=row, column=1, padx=(0, 16))

    _refresh()
    tk.Button(
        win, text="Apply",
        command=win.destroy,
        bg=C["accent"], fg=C["white"],
        relief="flat", bd=0, cursor="hand2",
        font=("Segoe UI", 9, "bold"),
        activebackground=C["accent2"],
        activeforeground=C["white"],
        padx=28, pady=6,
    ).grid(row=5, column=0, columnspan=2, pady=(8, 16))
    parent.wait_window(win)
