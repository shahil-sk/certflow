"""
Single scrollable field list inside the control panel.
Builds and rebuilds the per-field rows when Excel data changes.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C
from app.ui.widgets import label

_ALIGN_OPTIONS = ["left", "center", "right"]
_ALIGN_ICONS   = {"left": "\u2b05", "center": "\u2b0c", "right": "\u27a1"}  # ← ⬌ ➡


class FieldList(tk.Frame):
    """
    Drop-in replacement for the old fields_frame_outer.
    Call .rebuild(...) to refresh after Excel load.
    """

    def __init__(self, parent):
        super().__init__(parent, bg=C["surface"])
        self.pack(fill="x")

    def rebuild(
        self,
        fields: list,
        field_vars: dict,
        font_settings: dict,
        available_fonts: list,
        update_cb,
        color_cb,
    ) -> None:
        for w in self.winfo_children():
            w.destroy()

        scroll_cv = tk.Canvas(
            self, height=310, bg=C["surface"], highlightthickness=0)
        vsb = ttk.Scrollbar(
            self, orient="vertical", command=scroll_cv.yview,
            style="Flat.Vertical.TScrollbar")
        inner = tk.Frame(scroll_cv, bg=C["surface"])
        inner.bind(
            "<Configure>",
            lambda e: scroll_cv.configure(
                scrollregion=scroll_cv.bbox("all")),
        )
        scroll_cv.create_window((0, 0), window=inner, anchor="nw")
        scroll_cv.configure(yscrollcommand=vsb.set)

        # bind mousewheel on the inner list
        scroll_cv.bind("<Enter>",
            lambda e, cv=scroll_cv: cv.bind_all(
                "<MouseWheel>",
                lambda ev: cv.yview_scroll(int(-1*(ev.delta/120)), "units")))
        scroll_cv.bind("<Leave>",
            lambda e, cv=scroll_cv: cv.unbind_all("<MouseWheel>"))

        for i, field in enumerate(fields):
            _FieldRow(inner, field, field_vars, font_settings,
                      available_fonts, update_cb, color_cb,
                      alt=(i % 2 == 1))

        vsb.pack(side="right", fill="y")
        scroll_cv.pack(side="left", fill="both", expand=True)


class _FieldRow(tk.Frame):
    """One row in the field list."""

    def __init__(
        self, parent, field, field_vars, font_settings,
        available_fonts, update_cb, color_cb, alt=False,
    ):
        bg = C["row_alt"] if alt else C["surface"]
        super().__init__(parent, bg=bg, pady=8, padx=12)
        self.pack(fill="x")
        self._build(field, field_vars, font_settings,
                    available_fonts, update_cb, color_cb, bg)

    def _build(self, field, field_vars, font_settings,
               available_fonts, update_cb, color_cb, bg):

        # ── Row 1: field name + visibility toggle ───────────────────────
        top = tk.Frame(self, bg=bg)
        top.pack(fill="x", pady=(0, 4))
        label(top, field.title(), font_size=9, bold=True, bg=bg).pack(side="left")
        label(top, "Visible", font_size=8, color=C["subtext"],
              bg=bg).pack(side="right", padx=(0, 2))
        tk.Checkbutton(
            top, variable=field_vars[field],
            bg=bg, activebackground=bg, relief="flat", bd=0,
            command=lambda f=field: update_cb(f),
        ).pack(side="right")

        # ── Row 2: size + font + colour ──────────────────────────────────
        ctrl = tk.Frame(self, bg=bg)
        ctrl.pack(fill="x", pady=(0, 4))

        label(ctrl, "Size", font_size=8, color=C["subtext"], bg=bg).pack(side="left")
        spin = tk.Spinbox(
            ctrl, from_=8, to=300, width=4,
            textvariable=font_settings[field]["size"],
            font=("Segoe UI", 9), relief="flat", bd=1,
            command=lambda f=field: update_cb(f),
        )
        spin.bind("<Return>", lambda e, f=field: update_cb(f))
        spin.pack(side="left", padx=(3, 10))

        label(ctrl, "Font", font_size=8, color=C["subtext"], bg=bg).pack(side="left")
        cb = ttk.Combobox(
            ctrl, values=available_fonts,
            textvariable=font_settings[field]["font_name"],
            width=12, state="readonly", style="Flat.TCombobox",
        )
        cb.bind("<<ComboboxSelected>>", lambda e, f=field: update_cb(f))
        cb.pack(side="left", padx=(3, 8))

        swatch = tk.Label(
            ctrl, width=2, height=1,
            bg=font_settings[field]["color"].get(),
            relief="flat", bd=1,
            highlightthickness=1, highlightbackground=C["border"],
        )
        swatch.pack(side="left", padx=(0, 3))
        font_settings[field]["_swatch"] = swatch

        tk.Button(
            ctrl, text="Color",
            command=lambda f=field: color_cb(f),
            bg="#e8eaf0", fg=C["text"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 8),
            activebackground="#d0d4df",
            padx=7, pady=2,
        ).pack(side="left")

        # ── Row 3: alignment toggle buttons ─────────────────────────────
        align_row = tk.Frame(self, bg=bg)
        align_row.pack(fill="x", pady=(2, 0))
        label(align_row, "Align", font_size=8,
              color=C["subtext"], bg=bg).pack(side="left")

        align_var = font_settings[field]["align"]

        btn_refs = {}   # keep refs so we can highlight the active one

        def _set_align(val, f=field):
            align_var.set(val)
            for v, b in btn_refs.items():
                b.config(
                    bg=C["accent"]  if v == val else "#e8eaf0",
                    fg=C["white"]   if v == val else C["text"],
                )
            update_cb(f)

        btn_frame = tk.Frame(align_row, bg=bg)
        btn_frame.pack(side="left", padx=(6, 0))

        for opt in _ALIGN_OPTIONS:
            b = tk.Button(
                btn_frame,
                text=f" {opt.capitalize()} ",
                command=lambda v=opt: _set_align(v),
                bg=C["accent"] if align_var.get() == opt else "#e8eaf0",
                fg=C["white"]  if align_var.get() == opt else C["text"],
                relief="flat", bd=0, cursor="hand2",
                font=("Segoe UI", 8),
                activebackground=C["accent2"],
                activeforeground=C["white"],
                padx=6, pady=2,
            )
            b.pack(side="left", padx=(0, 2))
            btn_refs[opt] = b
