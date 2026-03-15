"""
All PIL drawing operations live here.
No tkinter widgets are created in this module.
"""
import io

from PIL import Image, ImageDraw

from app.helpers import hex_to_rgb
from app.font_manager import resolve_font

_LANCZOS = Image.Resampling.LANCZOS

# Mapping: alignment string -> PIL anchor string
# PIL anchor: first char = horizontal (l/m/r), second = vertical (a/m/d)
_ANCHOR = {
    "left":   "lm",
    "center": "mm",
    "right":  "rm",
}


def _draw_offset(align: str, tw: float, x: float) -> float:
    """
    Return the x coordinate to pass to draw.text().
    PIL anchors handle the horizontal shift themselves, but we still need
    to map the placeholder pin-point (x) correctly:
      left   -> x is the left edge of the text
      center -> x is the centre of the text
      right  -> x is the right edge of the text
    """
    if align == "left":
        return x
    if align == "right":
        return x
    return x   # center: anchor="mm" handles it


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
        if var is None or not var.get():
            continue
        if field not in positions:
            continue
        try:
            x, y   = positions[field]
            s      = font_settings[field]
            size   = s["size"].get()
            color  = s["color"].get()
            fname  = s["font_name"].get()
            align  = s.get("align", tk_str_or_default(s, "align", "center"))
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


def tk_str_or_default(settings: dict, key: str, default: str) -> str:
    """Safely read a StringVar or plain str from a font_settings sub-dict."""
    val = settings.get(key)
    if val is None:
        return default
    try:
        return val.get()
    except AttributeError:
        return str(val)


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
    size  = s["size"].get()
    color = s["color"].get()
    fname = s["font_name"].get()
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
