"""
All PIL drawing operations live here.
No tkinter widgets are created in this module.

Per-field controls:
  size, color, font_name, align     -- existing
  opacity   (0-100 int)             -- new: text layer alpha
  shadow    (bool)                  -- new: drop shadow
  shadow_offset (int, px)           -- new: shadow distance
  outline   (bool)                  -- new: stroke around text
  outline_width (int, px)           -- new: stroke width
"""
import io

from PIL import Image, ImageDraw, ImageFilter

from app.helpers import hex_to_rgb
from app.font_manager import resolve_font

_LANCZOS = Image.Resampling.LANCZOS

_ANCHOR = {
    "left":   "lm",
    "center": "mm",
    "right":  "rm",
}


def _get(var, default=""):
    """Read a tkinter Var or plain value safely."""
    try:
        return var.get()
    except AttributeError:
        return var if var is not None else default


def _draw_text_layer(
    size_px: int,
    text: str,
    font,
    color_rgb: tuple,
    opacity: int,
    shadow: bool,
    shadow_offset: int,
    outline: bool,
    outline_width: int,
    anchor: str,
    x: float,
    y: float,
    canvas_w: int,
    canvas_h: int,
    draw: ImageDraw.ImageDraw,
    base_img: Image.Image,
) -> None:
    """
    Draw text (with optional shadow / outline / opacity) onto draw/base_img.
    Uses an RGBA overlay so opacity blends correctly.
    """
    r, g, b = color_rgb[:3]
    alpha   = max(0, min(255, int(opacity / 100 * 255)))

    # -- shadow layer
    if shadow and shadow_offset > 0:
        sh_layer = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
        sh_draw  = ImageDraw.Draw(sh_layer)
        sh_draw.text(
            (x + shadow_offset, y + shadow_offset),
            text, font=font,
            fill=(0, 0, 0, min(alpha, 180)),
            anchor=anchor,
        )
        sh_layer = sh_layer.filter(
            ImageFilter.GaussianBlur(radius=max(1, shadow_offset // 2)))
        base_img.paste(
            Image.new("RGBA", (canvas_w, canvas_h), (0,0,0,0)),
            mask=sh_layer.split()[3],
        )
        # composite shadow under text
        base_img.alpha_composite(sh_layer)

    # -- outline layer
    if outline and outline_width > 0:
        out_layer = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
        out_draw  = ImageDraw.Draw(out_layer)
        for ox in range(-outline_width, outline_width + 1):
            for oy in range(-outline_width, outline_width + 1):
                if ox == 0 and oy == 0:
                    continue
                out_draw.text(
                    (x + ox, y + oy), text, font=font,
                    fill=(0, 0, 0, alpha),
                    anchor=anchor,
                )
        base_img.alpha_composite(out_layer)

    # -- main text layer
    txt_layer = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    txt_draw  = ImageDraw.Draw(txt_layer)
    txt_draw.text(
        (x, y), text, font=font,
        fill=(r, g, b, alpha),
        anchor=anchor,
    )
    base_img.alpha_composite(txt_layer)


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
    # Work in RGBA for compositing, convert at end
    img   = img.convert("RGBA")
    iw, ih = img.size

    for field in fields:
        var = field_vars.get(field)
        if var is None or not _get(var):
            continue
        if field not in positions:
            continue
        try:
            x, y  = positions[field]
            s     = font_settings[field]
            size  = _get(s["size"],         32)
            color = _get(s["color"],         "#000000")
            fname = _get(s["font_name"],     "")
            align = _get(s.get("align"),     "center")
            opac  = _get(s.get("opacity"),   100)
            shad  = _get(s.get("shadow"),    False)
            shoff = _get(s.get("shadow_offset"), 4)
            outl  = _get(s.get("outline"),   False)
            outw  = _get(s.get("outline_width"), 2)

            font   = resolve_font(available_fonts, fname, size)
            text   = student.get(field, "")
            anchor = _ANCHOR.get(align, "mm")

            _draw_text_layer(
                size_px=size,
                text=text, font=font,
                color_rgb=hex_to_rgb(color),
                opacity=opac,
                shadow=shad,  shadow_offset=shoff,
                outline=outl, outline_width=outw,
                anchor=anchor, x=x, y=y,
                canvas_w=iw, canvas_h=ih,
                draw=None, base_img=img,
            )
        except Exception as exc:
            print(f"[renderer] {field}: {exc}")

    return img.convert("RGB")


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
    size  = _get(s["size"],     32)
    color = _get(s["color"],    "#000000")
    fname = _get(s["font_name"], "")
    opac  = _get(s.get("opacity"), 100)
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
    r, g, b = hex_to_rgb(color)[:3]
    alpha   = max(0, min(255, int(opac / 100 * 255)))
    draw.text((0, yo), text, font=font, fill=(r, g, b, alpha))

    sw = max(int(tw / scale_x), 1)
    sh = max(int(th / scale_y), 1)
    return img.resize((sw, sh), _LANCZOS)


def image_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG", optimize=True)
    buf.seek(0)
    return buf.read()
