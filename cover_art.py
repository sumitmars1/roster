from __future__ import annotations

import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

A4_PORTRAIT_PX = (1654, 2339)  # 200 DPI A4

_TITLE_FONT_CANDIDATES = [
    "/Users/zhangyu/Library/Fonts/FZXBSK.ttf",
    "/Users/zhangyu/Library/Fonts/FZXBSJW.ttf",
    "/Users/zhangyu/Library/Fonts/FZHTK.ttf",
    "/Library/Fonts/msyhbd.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]

_SUB_FONT_CANDIDATES = [
    "/Users/zhangyu/Library/Fonts/FZHTK.ttf",
    "/Users/zhangyu/Library/Fonts/FZKTK.ttf",
    "/Users/zhangyu/Library/Fonts/FZFSK.ttf",
    "/Library/Fonts/msyh.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]


def _pick_font(candidates: list[str], size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in candidates:
        p = Path(path)
        if p.is_file():
            try:
                return ImageFont.truetype(str(p), size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def _fit_cover_canvas(img: Image.Image, target_size: tuple[int, int]) -> Image.Image:
    tw, th = target_size
    sw, sh = img.size
    if sw <= 0 or sh <= 0:
        return Image.new("RGB", target_size, (246, 248, 252))
    scale = max(tw / sw, th / sh)
    nw, nh = int(round(sw * scale)), int(round(sh * scale))
    resized = img.resize((nw, nh), Image.Resampling.LANCZOS)
    left = max(0, (nw - tw) // 2)
    top = max(0, (nh - th) // 2)
    return resized.crop((left, top, left + tw, top + th)).convert("RGBA")


def _region_luminance(img: Image.Image, box: tuple[int, int, int, int]) -> float:
    crop = img.crop(box).convert("L")
    pixels = list(crop.get_flattened_data() if hasattr(crop, "get_flattened_data") else crop.getdata())
    if not pixels:
        return 255.0
    return sum(pixels) / len(pixels)


def _draw_centered_text(
    draw: ImageDraw.ImageDraw,
    center_x: int,
    y: int,
    text: str,
    font: ImageFont.ImageFont,
    fill: tuple[int, int, int, int],
) -> int:
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    draw.text((center_x - w // 2, y), text, font=font, fill=fill)
    return h


def generate_cover_from_background(
    bg_path: Path | None,
    out_path: Path,
    event_title: str,
    *,
    subtitle: str = "秩  序  册",
    footer_lines: list[str] | None = None,
) -> Path:
    """Compose a print-ready A4 cover PNG from an event key-visual background or clean backdrop."""
    tw, th = A4_PORTRAIT_PX
    if bg_path and Path(bg_path).is_file():
        with Image.open(bg_path) as raw:
            base = _fit_cover_canvas(raw, A4_PORTRAIT_PX)
    else:
        base = Image.new("RGBA", A4_PORTRAIT_PX, (245, 247, 250, 255))
        draw_bg = ImageDraw.Draw(base)
        draw_bg.rectangle([(0, 0), (tw, 48)], fill=(21, 78, 138, 255))
        draw_bg.rectangle([(0, th - 48), (tw, th)], fill=(21, 78, 138, 255))

    # Measure luminance in the upper-middle title band to guarantee crisp readability
    lum = _region_luminance(base, (120, int(th * 0.16), tw - 120, int(th * 0.54)))
    overlay = Image.new("RGBA", A4_PORTRAIT_PX, (0, 0, 0, 0))
    odraw = ImageDraw.Draw(overlay)

    if bg_path and Path(bg_path).is_file():
        if 85 < lum < 190:
            # Mid-tone busy key visual: add soft translucent bright panel around title
            odraw.rounded_rectangle(
                [(110, int(th * 0.17)), (tw - 110, int(th * 0.54))],
                radius=28,
                fill=(255, 255, 255, 205),
            )
            text_color = (18, 34, 64, 255)
            accent_color = (184, 34, 38, 255)
        elif lum <= 85:
            text_color = (255, 255, 255, 255)
            accent_color = (255, 214, 102, 255)
        else:
            text_color = (18, 34, 64, 255)
            accent_color = (184, 34, 38, 255)
    else:
        text_color = (18, 34, 64, 255)
        accent_color = (184, 34, 38, 255)

    composed = Image.alpha_composite(base, overlay)
    draw = ImageDraw.Draw(composed)

    title_font = _pick_font(_TITLE_FONT_CANDIDATES, 82)
    sub_font = _pick_font(_TITLE_FONT_CANDIDATES, 118)
    meta_font = _pick_font(_SUB_FONT_CANDIDATES, 46)

    clean_title = (event_title or "体育赛事").strip()
    lines = textwrap.wrap(clean_title, width=13) or [clean_title]
    cur_y = int(th * 0.21)
    for line in lines[:3]:
        h = _draw_centered_text(draw, tw // 2, cur_y, line, title_font, text_color)
        cur_y += h + 28

    cur_y += 45
    _draw_centered_text(draw, tw // 2, cur_y, subtitle, sub_font, accent_color)

    if footer_lines:
        foot_lum = _region_luminance(composed, (120, int(th * 0.78), tw - 120, int(th * 0.94)))
        foot_color = (255, 255, 255, 255) if foot_lum <= 105 else (30, 41, 59, 255)
        fy = int(th * 0.82)
        for fl in footer_lines[:3]:
            fl_clean = fl.strip()
            if not fl_clean:
                continue
            h = _draw_centered_text(draw, tw // 2, fy, fl_clean, meta_font, foot_color)
            fy += h + 20

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    composed.convert("RGB").save(str(out_path), format="PNG", dpi=(200, 200))
    return out_path
