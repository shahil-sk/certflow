"""
Project file save / load (the .certwiz JSON format).
Version 2.3 — adds filename_pattern.
"""
import json
from datetime import datetime


def serialise(
    template_path,
    excel_path,
    color_space: str,
    positions: dict,
    fields: list,
    font_settings: dict,
    field_vars: dict,
    filename_pattern: str = "",
) -> dict:
    return {
        "version":          "2.3",
        "last_modified":    datetime.now().isoformat(),
        "template_path":    template_path,
        "excel_path":       excel_path,
        "color_space":      color_space,
        "filename_pattern": filename_pattern,
        "positions":        positions,
        "field_settings":   {
            f: {
                "size":      font_settings[f]["size"].get(),
                "color":     font_settings[f]["color"].get(),
                "visible":   field_vars[f].get(),
                "font_name": font_settings[f]["font_name"].get(),
                "align":     font_settings[f]["align"].get(),
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
