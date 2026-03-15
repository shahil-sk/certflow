"""
Reusable tkinter widget factory functions.
All styling is derived from app.constants.C — never hard-code colours here.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C


# ---------------------------------------------------------------------------
# Basic factories
# ---------------------------------------------------------------------------
def flat_button(
    parent, text: str, command,
    bg: str, active_bg: str,
    font_size: int = 9, bold: bool = False,
    **kw,
) -> tk.Button:
    weight = "bold" if bold else "normal"
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=C["white"], relief="flat", cursor="hand2",
        font=("Segoe UI", font_size, weight),
        activebackground=active_bg, activeforeground=C["white"],
        bd=0, highlightthickness=0, **kw,
    )


def label(
    parent, text: str,
    font_size: int = 9, bold: bool = False,
    color: str | None = None, bg: str | None = None,
    **kw,
) -> tk.Label:
    weight = "bold" if bold else "normal"
    return tk.Label(
        parent, text=text,
        font=("Segoe UI", font_size, weight),
        fg=color or C["text"],
        bg=bg or C["surface"],
        **kw,
    )


def hsep(parent, padx: int = 0, pady: int = 0) -> tk.Frame:
    """1-px horizontal separator."""
    f = tk.Frame(parent, bg=C["border"], height=1)
    f.pack(fill="x", padx=padx, pady=pady)
    return f


def card(parent, **kw) -> tk.Frame:
    """A rounded-looking dark card frame."""
    return tk.Frame(
        parent,
        bg=kw.pop("bg", C["surface2"]),
        highlightthickness=1,
        highlightbackground=C["border"],
        **kw,
    )


# ---------------------------------------------------------------------------
# TTK styles — call once at startup
# ---------------------------------------------------------------------------
def setup_ttk_styles() -> None:
    s = ttk.Style()
    s.theme_use("clam")

    # progress bar
    s.configure(
        "Thin.Horizontal.TProgressbar",
        troughcolor=C["surface3"], background=C["accent"],
        thickness=5, borderwidth=0,
    )

    # combobox
    s.configure(
        "Flat.TCombobox",
        fieldbackground=C["surface3"], background=C["surface3"],
        foreground=C["text"], selectbackground=C["accent_dim"],
        borderwidth=0, relief="flat", arrowcolor=C["subtext"],
    )
    s.map(
        "Flat.TCombobox",
        fieldbackground=[("readonly", C["surface3"])],
        foreground=[("readonly", C["text"])],
    )

    # scrollbar
    s.configure(
        "Dark.Vertical.TScrollbar",
        background=C["surface3"], troughcolor=C["surface"],
        borderwidth=0, arrowsize=10, arrowcolor=C["subtext"],
    )
    s.configure(
        "Dark.Horizontal.TScrollbar",
        background=C["surface3"], troughcolor=C["surface"],
        borderwidth=0, arrowsize=10, arrowcolor=C["subtext"],
    )
    s.map("Dark.Vertical.TScrollbar",
          background=[("active", C["accent_dim"])])
    s.map("Dark.Horizontal.TScrollbar",
          background=[("active", C["accent_dim"])])
