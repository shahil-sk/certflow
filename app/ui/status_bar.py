"""
Thin status bar at the bottom of the window.
"""
import tkinter as tk

from app.constants import C


class StatusBar(tk.Frame):

    def __init__(self, parent):
        super().__init__(parent, bg=C["nav"], height=26)
        self.pack(side="bottom", fill="x")
        self.pack_propagate(False)

        # Left dot indicator
        self._dot = tk.Label(
            self, text="●", font=("Segoe UI", 8),
            fg=C["success"], bg=C["nav"],
        )
        self._dot.pack(side="left", padx=(12, 4))

        self._var = tk.StringVar(value="Ready")
        tk.Label(
            self, textvariable=self._var,
            bg=C["nav"], fg=C["subtext"],
            font=("Segoe UI", 8), anchor="w",
        ).pack(side="left", fill="x", expand=True)

        # Right: version badge
        tk.Label(
            self, text="CertWizard",
            bg=C["nav"], fg=C["muted"],
            font=("Segoe UI", 7), anchor="e", padx=12,
        ).pack(side="right")

    def set(self, msg: str, ok: bool = True) -> None:
        self._var.set(msg)
        self._dot.config(fg=C["success"] if ok else C["warning"])
