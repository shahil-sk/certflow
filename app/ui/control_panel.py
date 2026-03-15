"""
Left sidebar: field list, actions, filename pattern, progress, log.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C, SIDEBAR_W
from app.ui.widgets import label, hsep


class ControlPanel(tk.Frame):
    """
    Exposes:
      self.fields_frame      -- Frame where FieldList is injected
      self.progress          -- ttk.Progressbar
      self.info_text         -- tk.Text log
      self.filename_pattern  -- tk.StringVar  (e.g. "{name}_{date}")
    """

    def __init__(self, parent, preview_cmd, generate_cmd):
        super().__init__(
            parent,
            width=SIDEBAR_W,
            bg=C["surface"],
            highlightthickness=1,
            highlightbackground=C["border"],
        )
        self.pack(side="left", fill="y", padx=(0, 10), pady=6)
        self.pack_propagate(False)
        self.filename_pattern = tk.StringVar(value="")
        self._build(preview_cmd, generate_cmd)

    # ------------------------------------------------------------------
    def _build(self, preview_cmd, generate_cmd):
        # ── Header
        hdr = tk.Frame(self, bg=C["surface"], pady=12)
        hdr.pack(fill="x", padx=16)
        tk.Label(hdr, text="▣  Fields",
                 font=("Segoe UI", 10, "bold"),
                 fg=C["text"], bg=C["surface"]).pack(side="left")
        hsep(self, padx=0, pady=0)

        # ── Field list injection point
        self.fields_frame = tk.Frame(self, bg=C["surface"])
        self.fields_frame.pack(fill="x")

        hsep(self, padx=0, pady=0)

        # ── Filename pattern row
        fn_wrap = tk.Frame(self, bg=C["surface"], pady=8)
        fn_wrap.pack(fill="x", padx=14)
        fn_top = tk.Frame(fn_wrap, bg=C["surface"])
        fn_top.pack(fill="x", pady=(0, 4))
        tk.Label(fn_top, text="📄  Filename pattern",
                 font=("Segoe UI", 8, "bold"),
                 fg=C["subtext"], bg=C["surface"]).pack(side="left")
        tk.Label(fn_top, text="e.g. {name}_{date}",
                 font=("Segoe UI", 7),
                 fg=C["muted"], bg=C["surface"]).pack(side="right")
        fn_entry = tk.Entry(
            fn_wrap,
            textvariable=self.filename_pattern,
            bg=C["surface3"], fg=C["text"],
            insertbackground=C["accent"],
            relief="flat", bd=0,
            font=("Segoe UI", 9),
            highlightthickness=1, highlightbackground=C["border"],
        )
        fn_entry.pack(fill="x", ipady=5)

        hsep(self, padx=0, pady=0)

        # ── Action buttons
        btn_area = tk.Frame(self, bg=C["surface"], pady=10)
        btn_area.pack(fill="x", padx=14)

        tk.Button(
            btn_area, text="▶  Preview", command=preview_cmd,
            bg=C["surface2"], fg=C["success"],
            relief="flat", cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            activebackground=C["surface3"],
            activeforeground=C["success"],
            bd=0, highlightthickness=1,
            highlightbackground=C["success"],
            padx=14, pady=8,
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        tk.Button(
            btn_area, text="⚡  Generate", command=generate_cmd,
            bg=C["accent"], fg=C["white"],
            relief="flat", cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            activebackground=C["accent2"],
            activeforeground=C["white"],
            bd=0, highlightthickness=0,
            padx=14, pady=8,
        ).pack(side="left", fill="x", expand=True)

        # ── Progress bar
        prog_wrap = tk.Frame(self, bg=C["surface"])
        prog_wrap.pack(fill="x", padx=14, pady=(0, 10))
        self.progress = ttk.Progressbar(
            prog_wrap, orient="horizontal", mode="determinate",
            style="Thin.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x")

        hsep(self, padx=0, pady=0)

        # ── Log header
        log_hdr = tk.Frame(self, bg=C["surface"], pady=8)
        log_hdr.pack(fill="x", padx=14)
        tk.Label(log_hdr, text="⧉  Log",
                 font=("Segoe UI", 8, "bold"),
                 fg=C["subtext"], bg=C["surface"]).pack(side="left")
        tk.Button(
            log_hdr, text="clear",
            command=self._clear_log,
            bg=C["surface"], fg=C["muted"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 7),
            activebackground=C["surface"],
            activeforeground=C["subtext"],
        ).pack(side="right")

        # ── Log text
        log_wrap = tk.Frame(self, bg=C["surface"])
        log_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.info_text = tk.Text(
            log_wrap, wrap=tk.WORD,
            font=("Consolas", 8),
            bg=C["log_bg"], fg=C["subtext"],
            insertbackground=C["accent"],
            relief="flat", bd=0, padx=8, pady=6,
            state="disabled",
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.info_text.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(log_wrap, orient="vertical",
                            command=self.info_text.yview,
                            style="Dark.Vertical.TScrollbar")
        vsb.pack(side="right", fill="y")
        self.info_text.configure(yscrollcommand=vsb.set)

        self.info_text.tag_configure("ok",   foreground=C["success"])
        self.info_text.tag_configure("err",  foreground=C["danger"])
        self.info_text.tag_configure("hdr",  foreground=C["accent"])
        self.info_text.tag_configure("warn", foreground=C["warning"])

    # ------------------------------------------------------------------
    def append_log(self, msg: str, clear: bool = False) -> None:
        self.info_text.configure(state="normal")
        if clear:
            self.info_text.delete("1.0", tk.END)
        low = msg.lower()
        if "[error]" in low:
            tag = "err"
        elif "[warn]" in low or "duplicate" in low:
            tag = "warn"
        elif "done" in low or "saved" in low:
            tag = "ok"
        elif msg.startswith("-") or "starting" in low:
            tag = "hdr"
        else:
            tag = ""
        self.info_text.insert(tk.END, msg + "\n", tag)
        self.info_text.see(tk.END)
        self.info_text.configure(state="disabled")

    def set_progress(self, pct: float) -> None:
        self.progress.configure(value=pct)

    def _clear_log(self) -> None:
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", tk.END)
        self.info_text.configure(state="disabled")
