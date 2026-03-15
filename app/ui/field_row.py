"""
Scrollable field editor list in the sidebar.
Each field card: name+visibility, size+font+colour, alignment,
                 opacity slider, shadow toggle, outline toggle.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C
from app.ui.widgets import label

_ALIGN_OPTIONS = [("L", "left"), ("C", "center"), ("R", "right")]


class FieldList(tk.Frame):

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

        cv  = tk.Canvas(self, height=360, bg=C["surface"], highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=cv.yview,
                            style="Dark.Vertical.TScrollbar")
        inner = tk.Frame(cv, bg=C["surface"])
        inner.bind("<Configure>",
                   lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=vsb.set)

        cv.bind("<Enter>", lambda e, c=cv: c.bind_all(
            "<MouseWheel>",
            lambda ev: c.yview_scroll(int(-1*(ev.delta/120)), "units")))
        cv.bind("<Leave>", lambda e, c=cv: c.unbind_all("<MouseWheel>"))

        for i, field in enumerate(fields):
            _FieldCard(inner, field, field_vars, font_settings,
                       available_fonts, update_cb, color_cb,
                       alt=(i % 2 == 1))

        vsb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)


class _FieldCard(tk.Frame):

    def __init__(self, parent, field, field_vars, font_settings,
                 available_fonts, update_cb, color_cb, alt=False):
        bg = C["row_alt"] if alt else C["surface"]
        super().__init__(parent, bg=bg)
        self.pack(fill="x")
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", side="bottom")
        self._build(field, field_vars, font_settings,
                    available_fonts, update_cb, color_cb, bg)

    def _build(self, field, field_vars, font_settings,
               available_fonts, update_cb, color_cb, bg):
        pad = tk.Frame(self, bg=bg, padx=12, pady=8)
        pad.pack(fill="x")

        # ─ Row 1: field name + visible toggle ───────────────
        r1 = tk.Frame(pad, bg=bg)
        r1.pack(fill="x", pady=(0, 5))
        tk.Label(r1, text=field.upper(), font=("Segoe UI", 7, "bold"),
                 fg=C["accent"], bg=bg).pack(side="left")
        vis_frame = tk.Frame(r1, bg=bg)
        vis_frame.pack(side="right")
        tk.Label(vis_frame, text="show", font=("Segoe UI", 7),
                 fg=C["subtext"], bg=bg).pack(side="left", padx=(0, 3))
        _ToggleSwitch(vis_frame, field_vars[field],
                      lambda f=field: update_cb(f), bg=bg).pack(side="left")

        # ─ Row 2: size + font ─────────────────────────────
        r2 = tk.Frame(pad, bg=bg)
        r2.pack(fill="x", pady=(0, 4))
        tk.Label(r2, text="Sz", font=("Segoe UI", 7),
                 fg=C["subtext"], bg=bg).pack(side="left")
        spin = tk.Spinbox(
            r2, from_=6, to=300, width=4,
            textvariable=font_settings[field]["size"],
            font=("Segoe UI", 8),
            bg=C["surface3"], fg=C["text"],
            buttonbackground=C["surface3"],
            insertbackground=C["text"],
            relief="flat", bd=0,
            highlightthickness=1, highlightbackground=C["border"],
            command=lambda f=field: update_cb(f),
        )
        spin.bind("<Return>", lambda e, f=field: update_cb(f))
        spin.pack(side="left", padx=(3, 8))
        tk.Label(r2, text="Font", font=("Segoe UI", 7),
                 fg=C["subtext"], bg=bg).pack(side="left")
        cb = ttk.Combobox(
            r2, values=available_fonts,
            textvariable=font_settings[field]["font_name"],
            width=11, state="readonly", style="Flat.TCombobox",
        )
        cb.bind("<<ComboboxSelected>>", lambda e, f=field: update_cb(f))
        cb.pack(side="left", padx=(3, 0))

        # ─ Row 3: colour + alignment ────────────────────────
        r3 = tk.Frame(pad, bg=bg)
        r3.pack(fill="x", pady=(0, 4))
        swatch = tk.Button(
            r3, width=3, height=1,
            bg=font_settings[field]["color"].get(),
            relief="flat", bd=0, cursor="hand2",
            highlightthickness=2, highlightbackground=C["border"],
            command=lambda f=field: color_cb(f),
        )
        swatch.pack(side="left", padx=(0, 5))
        font_settings[field]["_swatch"] = swatch
        tk.Label(r3, text="Color", font=("Segoe UI", 7),
                 fg=C["subtext"], bg=bg).pack(side="left", padx=(0, 8))

        tk.Label(r3, text="Align", font=("Segoe UI", 7),
                 fg=C["subtext"], bg=bg).pack(side="left", padx=(0, 3))
        align_var = font_settings[field]["align"]
        pill_refs = {}

        def _set_align(val, f=field):
            align_var.set(val)
            for v, b in pill_refs.items():
                b.config(
                    bg=C["btn_active"] if v == val else C["btn_idle"],
                    fg=C["white"]      if v == val else C["subtext"],
                )
            update_cb(f)

        for sym, val in _ALIGN_OPTIONS:
            active = align_var.get() == val
            b = tk.Button(
                r3, text=sym,
                command=lambda v=val: _set_align(v),
                bg=C["btn_active"] if active else C["btn_idle"],
                fg=C["white"]      if active else C["subtext"],
                font=("Segoe UI", 8, "bold"),
                relief="flat", bd=0, cursor="hand2",
                activebackground=C["accent2"],
                activeforeground=C["white"],
                width=2, pady=1,
            )
            b.pack(side="left", padx=(0, 2))
            pill_refs[val] = b

        # ─ Row 4: opacity slider ───────────────────────────
        r4 = tk.Frame(pad, bg=bg)
        r4.pack(fill="x", pady=(0, 4))
        tk.Label(r4, text="Opacity", font=("Segoe UI", 7),
                 fg=C["subtext"], bg=bg).pack(side="left")
        opac_lbl = tk.Label(r4, text="100%", font=("Segoe UI", 7),
                            fg=C["text"], bg=bg, width=4)
        opac_lbl.pack(side="right")

        def _opac_changed(val, f=field):
            opac_lbl.config(text=f"{int(float(val))}%")
            update_cb(f)

        tk.Scale(
            r4, from_=0, to=100, orient="horizontal",
            variable=font_settings[field]["opacity"],
            command=_opac_changed,
            bg=bg, fg=C["text"],
            troughcolor=C["surface3"],
            highlightthickness=0, relief="flat",
            showvalue=False, length=140,
        ).pack(side="left", padx=(4, 0), fill="x", expand=True)

        # ─ Row 5: shadow + outline toggles ──────────────────
        r5 = tk.Frame(pad, bg=bg)
        r5.pack(fill="x")

        def _toggle_btn(parent, label_text, var, f):
            def _toggle():
                var.set(not var.get())
                b.config(
                    bg=C["btn_active"] if var.get() else C["btn_idle"],
                    fg=C["white"]      if var.get() else C["subtext"],
                )
                update_cb(f)
            b = tk.Button(
                parent, text=label_text,
                command=_toggle,
                bg=C["btn_active"] if var.get() else C["btn_idle"],
                fg=C["white"]      if var.get() else C["subtext"],
                font=("Segoe UI", 7),
                relief="flat", bd=0, cursor="hand2",
                activebackground=C["accent2"],
                activeforeground=C["white"],
                padx=6, pady=2,
            )
            b.pack(side="left", padx=(0, 4))
            return b

        _toggle_btn(r5, "Shadow",  font_settings[field]["shadow"],  field)
        tk.Label(r5, text="off", font=("Segoe UI", 6),
                 fg=C["muted"], bg=bg).pack(side="left", padx=(0, 8))
        _toggle_btn(r5, "Outline", font_settings[field]["outline"], field)
        tk.Label(r5, text="off", font=("Segoe UI", 6),
                 fg=C["muted"], bg=bg).pack(side="left")


# ---------------------------------------------------------------------------
class _ToggleSwitch(tk.Canvas):
    W, H, R = 32, 16, 8

    def __init__(self, parent, var, command, bg="#ffffff"):
        super().__init__(parent, width=self.W, height=self.H,
                         bg=bg, highlightthickness=0, cursor="hand2")
        self._var = var
        self._cmd = command
        self.bind("<Button-1>", self._toggle)
        var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _draw(self):
        self.delete("all")
        on    = self._var.get()
        track = C["accent"] if on else C["surface3"]
        self._rounded_rect(0, 0, self.W, self.H, self.R, fill=track)
        kx = self.W - self.R - 2 if on else self.R + 2
        self.create_oval(kx-5, self.H//2-5, kx+5, self.H//2+5,
                         fill=C["white"], outline="")

    def _rounded_rect(self, x1, y1, x2, y2, r, **kw):
        self.create_arc(x1,    y1,    x1+2*r, y1+2*r, start=90,  extent=90,  style="pieslice", outline="", **kw)
        self.create_arc(x2-2*r,y1,    x2,     y1+2*r, start=0,   extent=90,  style="pieslice", outline="", **kw)
        self.create_arc(x1,    y2-2*r,x1+2*r, y2,     start=180, extent=90,  style="pieslice", outline="", **kw)
        self.create_arc(x2-2*r,y2-2*r,x2,     y2,     start=270, extent=90,  style="pieslice", outline="", **kw)
        self.create_rectangle(x1+r, y1, x2-r, y2, outline="", **kw)
        self.create_rectangle(x1, y1+r, x2, y2-r, outline="", **kw)

    def _toggle(self, _=None):
        self._var.set(not self._var.get())
        if self._cmd:
            self._cmd()
