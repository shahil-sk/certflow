"""
Right-side canvas frame: template + draggable text placeholders.
Features:
  - Scroll (mousewheel + scrollbars)
  - Row preview switcher in toolbar
  - Ctrl+Scroll zoom (30% - 200%)
  - Undo / Redo for placeholder positions (Ctrl+Z / Ctrl+Y or Ctrl+Shift+Z)
"""
import platform
import tkinter as tk
from tkinter import ttk
from collections import deque

from PIL import Image, ImageTk

from app.constants import C, CANVAS_MAX_W, CANVAS_MAX_H
from app.image_renderer import render_placeholder

_LANCZOS   = Image.Resampling.LANCZOS
_MAX_UNDO  = 50


class CanvasArea(tk.Frame):

    def __init__(self, parent):
        super().__init__(
            parent, bg=C["surface"],
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.pack(side="left", fill="both", expand=True, pady=6)

        self._scale_x = self._scale_y = 1.0
        self._zoom    = 1.0
        self._pil_image     = None
        self._display_image = None
        self._ph_images:    dict = {}
        self._placeholders: dict = {}
        self._drag:         dict = {}
        self._preview_idx   = 0

        # undo / redo stacks: each entry is a snapshot {field: (x,y)}
        self._undo_stack: deque = deque(maxlen=_MAX_UNDO)
        self._redo_stack: deque = deque(maxlen=_MAX_UNDO)

        # injected by App
        self.font_settings:   dict = {}
        self.available_fonts: dict = {}
        self.excel_data:      list = []
        self.fields:          list = []

        self._build_toolbar()
        self._build_canvas()

    # ------------------------------------------------------------------
    # Toolbar
    # ------------------------------------------------------------------
    def _build_toolbar(self) -> None:
        tb = tk.Frame(self, bg=C["nav"], height=32)
        tb.pack(fill="x")
        tb.pack_propagate(False)

        self._zoom_label = tk.Label(
            tb, text="100%",
            font=("Segoe UI", 8), fg=C["subtext"], bg=C["nav"],
        )
        self._zoom_label.pack(side="left", padx=10)

        # undo / redo buttons
        for txt, cmd in (("\u21b6", self.undo), ("\u21b7", self.redo)):
            tk.Button(
                tb, text=txt, command=cmd,
                bg=C["nav"], fg=C["subtext"],
                relief="flat", bd=0, cursor="hand2",
                font=("Segoe UI", 11),
                activebackground=C["surface"],
                activeforeground=C["text"],
                padx=6, pady=2,
            ).pack(side="left", padx=(0, 2))

        nav = tk.Frame(tb, bg=C["nav"])
        nav.pack(side="left", expand=True)

        _btn = lambda txt, cmd: tk.Button(
            nav, text=txt, command=cmd,
            bg=C["nav"], fg=C["subtext"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9),
            activebackground=C["surface"],
            activeforeground=C["text"],
            padx=8, pady=2,
        )
        _btn("\u25c4", self._prev_row).pack(side="left")
        self._row_label = tk.Label(
            nav, text="row \u2014",
            font=("Segoe UI", 8), fg=C["text"], bg=C["nav"], width=12,
        )
        self._row_label.pack(side="left", padx=4)
        _btn("\u25ba", self._next_row).pack(side="left")

        tk.Label(
            tb, text="Ctrl+scroll=zoom  Ctrl+Z=undo",
            font=("Segoe UI", 7), fg=C["muted"], bg=C["nav"],
        ).pack(side="right", padx=10)

    # ------------------------------------------------------------------
    # Canvas
    # ------------------------------------------------------------------
    def _build_canvas(self) -> None:
        self._h_scroll = ttk.Scrollbar(
            self, orient="horizontal", style="Dark.Horizontal.TScrollbar")
        self._v_scroll = ttk.Scrollbar(
            self, orient="vertical",   style="Dark.Vertical.TScrollbar")

        self._canvas = tk.Canvas(
            self, bg=C["canvas_bg"], highlightthickness=0,
            xscrollcommand=self._h_scroll.set,
            yscrollcommand=self._v_scroll.set,
        )
        self._h_scroll.config(command=self._canvas.xview)
        self._v_scroll.config(command=self._canvas.yview)

        self._h_scroll.pack(side="bottom", fill="x")
        self._v_scroll.pack(side="right",  fill="y")
        self._canvas.pack(fill="both", expand=True, padx=1, pady=1)

        self._canvas.bind("<Enter>", self._bind_scroll)
        self._canvas.bind("<Leave>", self._unbind_scroll)

        # Keyboard shortcuts
        self._canvas.bind_all("<Control-z>", lambda e: self.undo())
        self._canvas.bind_all("<Control-Z>", lambda e: self.undo())
        self._canvas.bind_all("<Control-y>", lambda e: self.redo())
        self._canvas.bind_all("<Control-Y>", lambda e: self.redo())
        self._canvas.bind_all("<Control-Shift-z>", lambda e: self.redo())

    # ------------------------------------------------------------------
    @property
    def scale_x(self): return self._scale_x

    @property
    def scale_y(self): return self._scale_y

    # ------------------------------------------------------------------
    # Image loading
    # ------------------------------------------------------------------
    def load_image(self, pil_image) -> None:
        self._pil_image   = pil_image
        self._preview_idx = 0
        self._redraw_image()

    def _redraw_image(self) -> None:
        if not self._pil_image:
            return
        ow, oh = self._pil_image.size
        base_ratio = min(CANVAS_MAX_W / ow, CANVAS_MAX_H / oh)
        base_nw    = max(int(ow * base_ratio), 1)
        base_nh    = max(int(oh * base_ratio), 1)
        # scale_x/scale_y are pure image-to-base-display ratios, zoom-independent
        self._scale_x = ow / base_nw
        self._scale_y = oh / base_nh

        nw = max(int(base_nw * self._zoom), 1)
        nh = max(int(base_nh * self._zoom), 1)

        self._display_image = ImageTk.PhotoImage(
            self._pil_image.resize((nw, nh), _LANCZOS))
        self._canvas.config(
            width=min(nw, CANVAS_MAX_W),
            height=min(nh, CANVAS_MAX_H),
            scrollregion=(0, 0, nw, nh),
        )
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, image=self._display_image, anchor="nw")
        self._zoom_label.config(text=f"{int(self._zoom * 100)}%")

        for field, data in list(self._placeholders.items()):
            self.draw_placeholder(field, data["x"], data["y"])

    # ------------------------------------------------------------------
    # Row switcher
    # ------------------------------------------------------------------
    def _update_row_label(self) -> None:
        total = len(self.excel_data)
        self._row_label.config(
            text=f"row {self._preview_idx + 1} / {total}" if total else "row \u2014")

    def _prev_row(self) -> None:
        if not self.excel_data:
            return
        self._preview_idx = (self._preview_idx - 1) % len(self.excel_data)
        self._update_row_label()
        self._refresh_all_placeholders()

    def _next_row(self) -> None:
        if not self.excel_data:
            return
        self._preview_idx = (self._preview_idx + 1) % len(self.excel_data)
        self._update_row_label()
        self._refresh_all_placeholders()

    def _refresh_all_placeholders(self) -> None:
        for field, data in list(self._placeholders.items()):
            self.draw_placeholder(field, data["x"], data["y"])

    def _current_row(self) -> dict:
        return self.excel_data[self._preview_idx] if self.excel_data else {}

    # ------------------------------------------------------------------
    # Undo / Redo
    # ------------------------------------------------------------------
    def _snapshot(self) -> dict:
        return {f: (d["x"], d["y"]) for f, d in self._placeholders.items()}

    def _push_undo(self) -> None:
        self._undo_stack.append(self._snapshot())
        self._redo_stack.clear()

    def undo(self) -> None:
        if not self._undo_stack:
            return
        self._redo_stack.append(self._snapshot())
        state = self._undo_stack.pop()
        self._apply_snapshot(state)

    def redo(self) -> None:
        if not self._redo_stack:
            return
        self._undo_stack.append(self._snapshot())
        state = self._redo_stack.pop()
        self._apply_snapshot(state)

    def _apply_snapshot(self, state: dict) -> None:
        for field, (x, y) in state.items():
            if field in self._placeholders:
                self._placeholders[field]["x"] = x
                self._placeholders[field]["y"] = y
        self._redraw_image()

    # ------------------------------------------------------------------
    # Placeholder management
    # ------------------------------------------------------------------
    def draw_placeholder(self, field: str, x: float, y: float) -> None:
        if field in self._placeholders:
            self._canvas.delete(self._placeholders[field]["item"])

        saved = self.excel_data
        row   = self._current_row()
        self.excel_data = [row] if row else saved

        ph_img = render_placeholder(
            field, self.font_settings, self.available_fonts,
            self.excel_data, self._scale_x, self._scale_y, self._zoom,
        )
        self.excel_data = saved

        photo = ImageTk.PhotoImage(ph_img)
        self._ph_images[field] = photo

        item = self._canvas.create_image(x, y, image=photo, anchor="center")
        self._canvas.tag_bind(item, "<Button-1>",
                              lambda e, i=item: self._drag_start(e, i))
        self._canvas.tag_bind(item, "<B1-Motion>",
                              lambda e, i=item: self._drag_move(e, i))
        self._placeholders[field] = {"item": item, "x": x, "y": y}
        self._update_row_label()

    def create_placeholder(self, field: str, x=None, y=None) -> None:
        if field not in self.fields:
            return
        if x is None or y is None:
            cw  = self._canvas.winfo_width() or 800
            idx = self.fields.index(field)
            x, y = cw // 2, 50 + idx * 60
        self.draw_placeholder(field, x, y)

    def update_placeholder(self, field: str) -> None:
        if field in self._placeholders:
            p = self._placeholders[field]
            self.draw_placeholder(field, p["x"], p["y"])

    def get_scaled_positions(self) -> dict:
        """
        Convert canvas (display) coordinates back to original image coordinates.

        Canvas positions are stored at zoomed display space. We need to undo
        the zoom first, then apply the base scale to get pixel coordinates in
        the original full-resolution image.

            original_px = canvas_pos / zoom * scale
        """
        return {
            f: (
                d["x"] / self._zoom * self._scale_x,
                d["y"] / self._zoom * self._scale_y,
            )
            for f, d in self._placeholders.items()
        }

    def clear(self) -> None:
        self._canvas.delete("all")
        self._placeholders.clear()
        self._ph_images.clear()
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._preview_idx = 0
        self._update_row_label()

    # ------------------------------------------------------------------
    # Scroll + Zoom
    # ------------------------------------------------------------------
    def _bind_scroll(self, _event=None) -> None:
        system = platform.system()
        if system == "Windows":
            self._canvas.bind_all("<MouseWheel>",       self._on_mousewheel)
            self._canvas.bind_all("<Shift-MouseWheel>", self._scroll_x)
        elif system == "Darwin":
            self._canvas.bind_all("<MouseWheel>",       self._on_mousewheel_mac)
            self._canvas.bind_all("<Shift-MouseWheel>", self._scroll_x_mac)
        else:
            self._canvas.bind_all("<Button-4>",         self._scroll_up)
            self._canvas.bind_all("<Button-5>",         self._scroll_down)
            self._canvas.bind_all("<Shift-Button-4>",   self._scroll_left)
            self._canvas.bind_all("<Shift-Button-5>",   self._scroll_right)
            self._canvas.bind_all("<Control-Button-4>", self._zoom_in)
            self._canvas.bind_all("<Control-Button-5>", self._zoom_out)

    def _unbind_scroll(self, _event=None) -> None:
        for seq in ("<MouseWheel>", "<Shift-MouseWheel>",
                    "<Button-4>", "<Button-5>",
                    "<Shift-Button-4>", "<Shift-Button-5>",
                    "<Control-Button-4>", "<Control-Button-5>"):
            try:
                self._canvas.unbind_all(seq)
            except Exception:
                pass

    def _on_mousewheel(self, event):
        if event.state & 0x0004:
            self._zoom_in() if event.delta > 0 else self._zoom_out()
        else:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_mac(self, event):
        if event.state & 0x0004:
            self._zoom_in() if event.delta > 0 else self._zoom_out()
        else:
            self._canvas.yview_scroll(int(-1 * event.delta), "units")

    def _scroll_x(self, event):
        self._canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _scroll_x_mac(self, event):
        self._canvas.xview_scroll(int(-1 * event.delta), "units")

    def _scroll_up(self,    _e): self._canvas.yview_scroll(-1, "units")
    def _scroll_down(self,  _e): self._canvas.yview_scroll( 1, "units")
    def _scroll_left(self,  _e): self._canvas.xview_scroll(-1, "units")
    def _scroll_right(self, _e): self._canvas.xview_scroll( 1, "units")

    def _zoom_in(self, _e=None):
        if self._zoom < 2.0:
            self._zoom = round(min(self._zoom + 0.1, 2.0), 1)
            self._redraw_image()

    def _zoom_out(self, _e=None):
        if self._zoom > 0.3:
            self._zoom = round(max(self._zoom - 0.1, 0.3), 1)
            self._redraw_image()

    # ------------------------------------------------------------------
    # Drag  (pushes undo on button-release)
    # ------------------------------------------------------------------
    def _drag_start(self, event, item) -> None:
        self._push_undo()   # snapshot before move
        self._drag = {"item": item, "x": event.x, "y": event.y}

    def _drag_move(self, event, item) -> None:
        dx = event.x - self._drag["x"]
        dy = event.y - self._drag["y"]
        self._canvas.move(item, dx, dy)
        self._drag["x"] = event.x
        self._drag["y"] = event.y
        for f, d in self._placeholders.items():
            if d["item"] == item:
                d["x"] += dx
                d["y"] += dy
                break
