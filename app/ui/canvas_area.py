"""
Right-side canvas frame: displays the template and draggable text placeholders.
Supports scroll (mousewheel + scrollbars) for templates larger than the view.
"""
import platform
import tkinter as tk
from tkinter import ttk

from PIL import ImageTk

from app.constants import C, CANVAS_MAX_W, CANVAS_MAX_H
from app.image_renderer import render_placeholder


class CanvasArea(tk.Frame):
    """
    Owns the tk.Canvas and all placeholder management.
    Exposes:
      load_image(pil_image)        -- draw/redraw background template
      draw_placeholder(field, x, y)
      create_placeholder(field, x=None, y=None)
      update_placeholder(field)
      get_scaled_positions() -> dict
      scale_x / scale_y           -- read-only properties
    """

    def __init__(self, parent):
        super().__init__(
            parent, bg=C["surface"],
            highlightthickness=1, highlightbackground=C["border"],
        )
        self.pack(side="left", fill="both", expand=True, pady=6)

        self._build_canvas()

        self._scale_x = self._scale_y = 1.0
        self._display_image = None
        self._ph_images: dict = {}
        self._placeholders: dict = {}
        self._drag: dict = {}

        # state injected by App after Excel load
        self.font_settings:   dict = {}
        self.available_fonts: dict = {}
        self.excel_data:      list = []
        self.fields:          list = []

    # ------------------------------------------------------------------
    def _build_canvas(self) -> None:
        self._h_scroll = ttk.Scrollbar(
            self, orient="horizontal",
            style="Flat.Vertical.TScrollbar")
        self._v_scroll = ttk.Scrollbar(
            self, orient="vertical",
            style="Flat.Vertical.TScrollbar")

        self._canvas = tk.Canvas(
            self, bg=C["bg"], highlightthickness=0,
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

    # ------------------------------------------------------------------
    @property
    def scale_x(self): return self._scale_x

    @property
    def scale_y(self): return self._scale_y

    # ------------------------------------------------------------------
    def load_image(self, pil_image) -> None:
        """Scale image to fit the canvas view and redraw it."""
        ow, oh = pil_image.size
        ratio  = min(CANVAS_MAX_W / ow, CANVAS_MAX_H / oh)
        nw, nh = int(ow * ratio), int(oh * ratio)
        self._scale_x, self._scale_y = ow / nw, oh / nh

        self._display_image = ImageTk.PhotoImage(
            pil_image.resize((nw, nh), pil_image.LANCZOS))
        self._canvas.config(
            width=min(nw, CANVAS_MAX_W),
            height=min(nh, CANVAS_MAX_H),
            scrollregion=(0, 0, nw, nh),
        )
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, image=self._display_image, anchor="nw")

        for field, data in list(self._placeholders.items()):
            self.draw_placeholder(field, data["x"], data["y"])

    def draw_placeholder(self, field: str, x: float, y: float) -> None:
        if field in self._placeholders:
            self._canvas.delete(self._placeholders[field]["item"])

        ph_img = render_placeholder(
            field, self.font_settings, self.available_fonts,
            self.excel_data, self._scale_x, self._scale_y,
        )
        photo = ImageTk.PhotoImage(ph_img)
        self._ph_images[field] = photo

        item = self._canvas.create_image(x, y, image=photo, anchor="center")
        self._canvas.tag_bind(item, "<Button-1>",
                              lambda e, i=item: self._drag_start(e, i))
        self._canvas.tag_bind(item, "<B1-Motion>",
                              lambda e, i=item: self._drag_move(e, i))
        self._placeholders[field] = {"item": item, "x": x, "y": y}

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
        return {
            f: (d["x"] * self._scale_x, d["y"] * self._scale_y)
            for f, d in self._placeholders.items()
        }

    def clear(self) -> None:
        self._canvas.delete("all")
        self._placeholders.clear()
        self._ph_images.clear()

    # ------------------------------------------------------------------
    # Scroll handling
    # ------------------------------------------------------------------
    def _bind_scroll(self, _event=None) -> None:
        system = platform.system()
        if system == "Windows":
            self._canvas.bind_all("<MouseWheel>",    self._scroll_y)
            self._canvas.bind_all("<Shift-MouseWheel>", self._scroll_x)
        elif system == "Darwin":
            self._canvas.bind_all("<MouseWheel>",    self._scroll_y_mac)
            self._canvas.bind_all("<Shift-MouseWheel>", self._scroll_x_mac)
        else:  # Linux
            self._canvas.bind_all("<Button-4>",  self._scroll_up)
            self._canvas.bind_all("<Button-5>",  self._scroll_down)
            self._canvas.bind_all("<Shift-Button-4>", self._scroll_left)
            self._canvas.bind_all("<Shift-Button-5>", self._scroll_right)

    def _unbind_scroll(self, _event=None) -> None:
        for seq in ("<MouseWheel>", "<Shift-MouseWheel>",
                    "<Button-4>", "<Button-5>",
                    "<Shift-Button-4>", "<Shift-Button-5>"):
            try:
                self._canvas.unbind_all(seq)
            except Exception:
                pass

    def _scroll_y(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _scroll_x(self, event):
        self._canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _scroll_y_mac(self, event):
        self._canvas.yview_scroll(int(-1 * event.delta), "units")

    def _scroll_x_mac(self, event):
        self._canvas.xview_scroll(int(-1 * event.delta), "units")

    def _scroll_up(self,    _e): self._canvas.yview_scroll(-1, "units")
    def _scroll_down(self,  _e): self._canvas.yview_scroll( 1, "units")
    def _scroll_left(self,  _e): self._canvas.xview_scroll(-1, "units")
    def _scroll_right(self, _e): self._canvas.xview_scroll( 1, "units")

    # ------------------------------------------------------------------
    # Drag handling
    # ------------------------------------------------------------------
    def _drag_start(self, event, item) -> None:
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
