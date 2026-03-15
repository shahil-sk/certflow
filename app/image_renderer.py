"""
All PIL drawing operations live here.
No tkinter widgets are created in this module.
"""
import io

from PIL import Image, ImageDraw

from app.helpers import hex_to_rgb
from app.font_manager import resolve_font

_LANCZOS = Image.Resampling.LANCZOS

# PIL anchor: first char = horizontal (l/m/r), second = vertical (m = middle)
_ANCHOR = {
    "left":   "lm",
    "center": "mm",
    "right":  "rm",
}


def _get(var, default=""):
    """Read a tkinter StringVar/IntVar or plain value safely."""
    try:
        return var.get()
    except AttributeError:
        return var if var is not None else default


def draw_text_on_image(
    img: Image.Image,
    fields: list,
    field_vars: dict,
    font_settings: dict,
    available_fonts: dict,
    student: dict,
    positions: dict,
) -> Image.Image:
    """Render all visible field text onto img and return it."""
    draw = ImageDraw.Draw(img)
    for field in fields:
        var = field_vars.get(field)
        if var is None or not _get(var):
            continue
        if field not in positions:
            continue
        try:
            x, y   = positions[field]
            s      = font_settings[field]
            size   = _get(s["size"], 32)
            color  = _get(s["color"], "#000000")
            fname  = _get(s["font_name"], "")
            align  = _get(s.get("align"), "center")
            font   = resolve_font(available_fonts, fname, size)
            text   = student.get(field, "")
            anchor = _ANCHOR.get(align, "mm")
            draw.text(
                (x, y),
                text,
                font=font,
                fill=hex_to_rgb(color),
                anchor=anchor,
            )
        except Exception as exc:
            print(f"[renderer] {field}: {exc}")
    return img


def render_placeholder(
    field: str,
    font_settings: dict,
    available_fonts: dict,
    excel_data: list,
    scale_x: float,
    scale_y: float,
) -> Image.Image:
    """Return a PIL Image of the sample text scaled for canvas display."""
    s     = font_settings[field]
    size  = _get(s["size"], 32)
    color = _get(s["color"], "#000000")
    fname = _get(s["font_name"], "")
    font  = resolve_font(available_fonts, fname, size)
    text  = (excel_data[0].get(field, field) if excel_data else field) or field

    tmp  = Image.new("RGBA", (1, 1))
    draw = ImageDraw.Draw(tmp)
    tw   = max(int(draw.textlength(text, font=font)), 1)
    try:
        asc, desc = font.getmetrics()
        th = max(asc + desc, 1)
    except Exception:
        th = max(size, 1)

    img  = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    try:
        bbox = font.getbbox(text)
        yo   = (th - (bbox[3] - bbox[1])) // 2
    except Exception:
        yo = 0
    draw.text((0, yo), text, font=font, fill=hex_to_rgb(color))

    sw = max(int(tw / scale_x), 1)
    sh = max(int(th / scale_y), 1)
    return img.resize((sw, sh), _LANCZOS)


def image_to_bytes(img: Image.Image) -> bytes:
    """Encode a PIL image as PNG bytes in memory (no temp files)."""
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG", optimize=True)
    buf.seek(0)
    return buf.read()
