"""
Left sidebar: field list, actions, progress bar, generation log.
"""
import tkinter as tk
from tkinter import ttk

from app.constants import C, SIDEBAR_W
from app.ui.widgets import flat_button, label, hsep, card


class ControlPanel(tk.Frame):
    """
    Exposes:
      self.fields_frame   -- inner Frame where FieldList is injected
      self.progress       -- ttk.Progressbar
      self.info_text      -- tk.Text (log)
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
        self._build(preview_cmd, generate_cmd)

    # ------------------------------------------------------------------
    def _build(self, preview_cmd, generate_cmd):
        # ── Sidebar header
        hdr = tk.Frame(self, bg=C["surface"], pady=12)
        hdr.pack(fill="x", padx=16)
        tk.Label(
            hdr, text="▣  Fields",
            font=("Segoe UI", 10, "bold"),
            fg=C["text"], bg=C["surface"],
        ).pack(side="left")
        hsep(self, padx=0, pady=0)

        # ── Field list injection point
        self.fields_frame = tk.Frame(self, bg=C["surface"])
        self.fields_frame.pack(fill="x")

        hsep(self, padx=0, pady=0)

        # ── Action buttons
        btn_area = tk.Frame(self, bg=C["surface"], pady=12)
        btn_area.pack(fill="x", padx=14)

        preview_btn = tk.Button(
            btn_area,
            text="▶  Preview",
            command=preview_cmd,
            bg=C["surface2"], fg=C["success"],
            relief="flat", cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            activebackground=C["surface3"],
            activeforeground=C["success"],
            bd=0, highlightthickness=1,
            highlightbackground=C["success"],
            padx=14, pady=8,
        )
        preview_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        generate_btn = tk.Button(
            btn_area,
            text="⚡  Generate",
            command=generate_cmd,
            bg=C["accent"], fg=C["white"],
            relief="flat", cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            activebackground=C["accent2"],
            activeforeground=C["white"],
            bd=0, highlightthickness=0,
            padx=14, pady=8,
        )
        generate_btn.pack(side="left", fill="x", expand=True)

        # ── Progress bar
        prog_wrap = tk.Frame(self, bg=C["surface"])
        prog_wrap.pack(fill="x", padx=14, pady=(0, 10))
        self.progress = ttk.Progressbar(
            prog_wrap, orient="horizontal", mode="determinate",
            style="Thin.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x")

        hsep(self, padx=0, pady=0)

        # ── Log section header
        log_hdr = tk.Frame(self, bg=C["surface"], pady=8)
        log_hdr.pack(fill="x", padx=14)
        tk.Label(
            log_hdr, text="⧉  Log",
            font=("Segoe UI", 8, "bold"),
            fg=C["subtext"], bg=C["surface"],
        ).pack(side="left")

        # clear log button
        tk.Button(
            log_hdr, text="clear",
            command=self._clear_log,
            bg=C["surface"], fg=C["muted"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 7),
            activebackground=C["surface"],
            activeforeground=C["subtext"],
        ).pack(side="right")

        # ── Log text widget
        log_wrap = tk.Frame(self, bg=C["surface"])
        log_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.info_text = tk.Text(
            log_wrap,
            wrap=tk.WORD,
            font=("Consolas", 8),
            bg=C["log_bg"], fg=C["subtext"],
            insertbackground=C["accent"],
            relief="flat", bd=0, padx=8, pady=6,
            state="disabled",
            highlightthickness=1,
            highlightbackground=C["border"],
        )
        self.info_text.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(
            log_wrap, orient="vertical",
            command=self.info_text.yview,
            style="Dark.Vertical.TScrollbar",
        )
        vsb.pack(side="right", fill="y")
        self.info_text.configure(yscrollcommand=vsb.set)

        # tag for success / error colouring
        self.info_text.tag_configure("ok",  foreground=C["success"])
        self.info_text.tag_configure("err", foreground=C["danger"])
        self.info_text.tag_configure("hdr", foreground=C["accent"])

    # ------------------------------------------------------------------
    def append_log(self, msg: str, clear: bool = False) -> None:
        self.info_text.configure(state="normal")
        if clear:
            self.info_text.delete("1.0", tk.END)
        # simple tag detection
        tag = ""
        if msg.startswith("[") and "error" in msg.lower():
            tag = "err"
        elif "done" in msg.lower() or "saved" in msg.lower():
            tag = "ok"
        elif msg.startswith("-") or msg.startswith("Start"):
            tag = "hdr"
        self.info_text.insert(tk.END, msg + "\n", tag)
        self.info_text.see(tk.END)
        self.info_text.configure(state="disabled")

    def set_progress(self, pct: float) -> None:
        self.progress.configure(value=pct)

    def _clear_log(self) -> None:
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", tk.END)
        self.info_text.configure(state="disabled")
