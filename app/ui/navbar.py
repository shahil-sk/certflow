"""
Top navigation bar.
"""
import tkinter as tk

from app.constants import C, APP_TITLE, APP_VERSION
from app.ui.widgets import flat_button, label


class NavBar(tk.Frame):

    def __init__(self, parent, save_cmd, load_cmd,
                 load_template_cmd, load_excel_cmd):
        super().__init__(parent, bg=C["nav"], height=52)
        self.pack(fill="x")
        self.pack_propagate(False)
        self._build(save_cmd, load_cmd, load_template_cmd, load_excel_cmd)

    def _build(self, save_cmd, load_cmd, load_template_cmd, load_excel_cmd):
        # Brand
        brand = tk.Frame(self, bg=C["nav"])
        brand.pack(side="left", padx=18)

        tk.Label(
            brand, text="◆", font=("Segoe UI", 14),
            fg=C["accent"], bg=C["nav"],
        ).pack(side="left", padx=(0, 6))
        tk.Label(
            brand, text=APP_TITLE,
            font=("Segoe UI", 12, "bold"),
            fg=C["text"], bg=C["nav"],
        ).pack(side="left")
        tk.Label(
            brand, text=f" v{APP_VERSION}",
            font=("Segoe UI", 8),
            fg=C["subtext"], bg=C["nav"],
        ).pack(side="left", pady=(3, 0))

        # Right-side actions
        right = tk.Frame(self, bg=C["nav"])
        right.pack(side="right", padx=14)

        # Project dropdown
        proj_btn = tk.Menubutton(
            right, text="☰  Project",
            bg=C["surface"], fg=C["subtext"],
            relief="flat", font=("Segoe UI", 9),
            padx=12, pady=6,
            cursor="hand2",
            activebackground=C["surface2"],
            activeforeground=C["text"],
            bd=0, highlightthickness=1,
            highlightbackground=C["border"],
        )
        proj_btn.menu = tk.Menu(
            proj_btn, tearoff=0,
            bg=C["surface2"], fg=C["text"],
            activebackground=C["accent"],
            activeforeground=C["white"],
            font=("Segoe UI", 9),
            bd=0, relief="flat",
        )
        proj_btn["menu"] = proj_btn.menu
        proj_btn.menu.add_command(label="📂  Save Project", command=save_cmd)
        proj_btn.menu.add_command(label="📁  Load Project", command=load_cmd)
        proj_btn.pack(side="left", padx=(0, 8))

        # Primary action buttons
        for text, cmd, bg, abg in (
            ("+ Template", load_template_cmd, C["surface"],  C["surface2"]),
            ("+ Excel",    load_excel_cmd,    C["accent"],   C["accent2"]),
        ):
            b = tk.Button(
                right, text=text, command=cmd,
                bg=bg, fg=C["text"] if bg == C["surface"] else C["white"],
                relief="flat", cursor="hand2",
                font=("Segoe UI", 9),
                activebackground=abg,
                activeforeground=C["white"] if bg != C["surface"] else C["text"],
                bd=0, highlightthickness=1,
                highlightbackground=C["border"],
                padx=14, pady=6,
            )
            b.pack(side="left", padx=(0, 6))
