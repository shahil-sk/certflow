"""
Project file save / load (.certwiz JSON).
Version 2.4 — adds opacity, shadow, outline.
"""
import json
from datetime import datetime


def serialise(
    template_path, excel_path, color_space,
    positions, fields, font_settings, field_vars,
    filename_pattern="",
) -> dict:
    def _get(var, default=None):
        try:    return var.get()
        except: return default

    return {
        "version":          "2.4",
        "last_modified":    datetime.now().isoformat(),
        "template_path":    template_path,
        "excel_path":       excel_path,
        "color_space":      color_space,
        "filename_pattern": filename_pattern,
        "positions":        positions,
        "field_settings":   {
            f: {
                "size":          _get(font_settings[f]["size"],         32),
                "color":         _get(font_settings[f]["color"],        "#000000"),
                "visible":       _get(field_vars[f],                    True),
                "font_name":     _get(font_settings[f]["font_name"],    ""),
                "align":         _get(font_settings[f]["align"],        "center"),
                "opacity":       _get(font_settings[f].get("opacity"),  100),
                "shadow":        _get(font_settings[f].get("shadow"),   False),
                "shadow_offset": _get(font_settings[f].get("shadow_offset"), 4),
                "outline":       _get(font_settings[f].get("outline"),  False),
                "outline_width": _get(font_settings[f].get("outline_width"), 2),
            }
            for f in fields
        },
    }


def save(path: str, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)


def load(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)
