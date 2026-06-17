"""Рисование последовательной схемы: символ сопротивления с подписью (условное обозначение)."""

from __future__ import annotations

import io
import os
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PIL import ImageFont

from .models import EQUIPMENT_LABEL_BY_KEY, SoptEquipmentItem
from .resistance import item_resistance_ohm

_FIG_STANDARD_W = 560
_FIG_STANDARD_H = 300


def _trim_white_borders(img: "Image.Image", *, margin: int = 0) -> "Image.Image":
    px = img.load()
    w, h = img.size
    min_x, min_y, max_x, max_y = w, h, -1, -1
    for y in range(h):
        for x in range(w):
            if px[x, y] != (255, 255, 255):
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)
    if max_x < 0:
        return img
    min_x = max(0, min_x - margin)
    min_y = max(0, min_y - margin)
    max_x = min(w - 1, max_x + margin)
    max_y = min(h - 1, max_y + margin)
    return img.crop((min_x, min_y, max_x + 1, max_y + 1))


def _fit_image_to_standard_canvas(
    img: "Image.Image",
    *,
    width: int = _FIG_STANDARD_W,
    height: int = _FIG_STANDARD_H,
    margin: int = 10,
) -> "Image.Image":
    from PIL import Image

    trimmed = _trim_white_borders(img)
    inner_w = max(1, width - 2 * margin)
    inner_h = max(1, height - 2 * margin)
    src_w, src_h = trimmed.size
    scale = min(inner_w / src_w, inner_h / src_h)
    dst_w = max(1, int(src_w * scale))
    dst_h = max(1, int(src_h * scale))
    if (dst_w, dst_h) != trimmed.size:
        trimmed = trimmed.resize((dst_w, dst_h), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (width, height), "white")
    canvas.paste(trimmed, ((width - dst_w) // 2, (height - dst_h) // 2))
    return canvas


def _load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    from PIL import ImageFont

    if bold:
        candidates = ("arialbd.ttf", "Arial Bold.ttf", "segoeuib.ttf", "DejaVuSans-Bold.ttf")
    else:
        candidates = ("arial.ttf", "Arial.ttf", "segoeui.ttf", "DejaVuSans.ttf")
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _lightning_icon_candidates() -> list[Path]:
    env_path = os.environ.get("ELECTROPROJECT_LIGHTNING_ICON", "").strip()
    candidates: list[Path] = []
    if env_path:
        candidates.append(Path(env_path))
    candidates.append(Path(__file__).with_name("assets") / "lightning.png")
    return candidates


def _load_lightning_icon() -> "Image.Image | None":
    from PIL import Image

    for path in _lightning_icon_candidates():
        try:
            icon = Image.open(path).convert("RGBA")
            px = icon.load()
            for y in range(icon.height):
                for x in range(icon.width):
                    r, g, b, a = px[x, y]
                    red_dominance = r - max(g, b)
                    if r < 95 or red_dominance < 35:
                        px[x, y] = (r, g, b, 0)
                        continue
                    alpha_scale = min(1.0, max(0.35, red_dominance / 170.0))
                    px[x, y] = (r, g, b, int(a * alpha_scale))
            bb = icon.getbbox()
            if bb is not None:
                icon = icon.crop(bb)
            return icon
        except OSError:
            continue
    return None


def _akb_ready_schematic_candidates() -> list[Path]:
    env_path = os.environ.get("ELECTROPROJECT_AKB_SCHEMATIC", "").strip()
    candidates: list[Path] = []
    if env_path:
        candidates.append(Path(env_path))
    candidates.append(Path(__file__).with_name("assets") / "akb_input_schematic.png")
    return candidates


def _load_akb_ready_schematic() -> "Image.Image | None":
    from PIL import Image

    for path in _akb_ready_schematic_candidates():
        try:
            return Image.open(path).convert("RGB")
        except OSError:
            continue
    return None


def _label_text(item: SoptEquipmentItem) -> str:
    text = item.designation.strip()
    if text:
        return text
    return EQUIPMENT_LABEL_BY_KEY.get(item.item_type, "—")[:24]


def render_akb_input_schematic() -> "Image.Image":
    from PIL import Image, ImageDraw

    pad_left = 52
    pad_top = 40
    img = Image.new("RGB", (560 + pad_left, 300 + pad_top), "white")
    draw = ImageDraw.Draw(img)
    line_color = "black"
    wire_w = 2
    box_w = 44
    box_h = 20
    left_res_w = box_h
    left_res_h = box_w
    font_fu = _load_font(22, bold=False)
    font_polarity = _load_font(26, bold=True)
    font_left_r = _load_font(24, bold=False)
    font_left = _load_font(22, bold=False)
    font_left_sub = _load_font(14, bold=False)

    x_left = 92 + pad_left
    x_right = 452 + pad_left
    y_top = 78 + pad_top
    y_bottom = 228 + pad_top
    y_mid = (y_top + y_bottom) // 2

    # Верхняя и нижняя шины с предохранителями (координаты как на рис.2).
    fu1_left = 228 + pad_left
    fu2_left = fu1_left
    fu_cx = fu1_left + box_w // 2
    draw.line([(x_left, y_top), (fu1_left, y_top)], fill=line_color, width=wire_w)
    draw.rectangle([fu1_left, y_top - box_h // 2, fu1_left + box_w, y_top + box_h // 2], outline=line_color, width=2)
    draw.line([(fu1_left + box_w, y_top), (x_right, y_top)], fill=line_color, width=wire_w)

    draw.line([(x_left, y_bottom), (fu2_left, y_bottom)], fill=line_color, width=wire_w)
    draw.rectangle(
        [fu2_left, y_bottom - box_h // 2, fu2_left + box_w, y_bottom + box_h // 2],
        outline=line_color,
        width=2,
    )
    draw.line([(fu2_left + box_w, y_bottom), (x_right, y_bottom)], fill=line_color, width=wire_w)

    # Вертикальная левая ветвь.
    draw.line([(x_left, y_top), (x_left, y_bottom)], fill=line_color, width=wire_w)

    # Левый вертикальный резистор Rаб+Rпер (как на рис.2–3, без числовых значений).
    left_rect_top = y_mid - left_res_h // 2
    draw.rectangle(
        [x_left - left_res_w // 2, left_rect_top, x_left + left_res_w // 2, left_rect_top + left_res_h],
        outline=line_color,
        width=2,
        fill="white",
    )

    # Клеммы +/- справа.
    draw.line([(x_right, y_top - 10), (x_right, y_top + 10)], fill=line_color, width=2)
    draw.line([(x_right, y_bottom - 10), (x_right, y_bottom + 10)], fill=line_color, width=2)
    draw.text((x_right + 12, y_top - 18), "+", fill=line_color, font=font_polarity)
    draw.text((x_right + 11, y_bottom - 18), "-", fill=line_color, font=font_polarity)

    # Подписи FU1 / FU2 — слитно, по центру над прямоугольниками.
    fu_gap = 10
    for label, y_line in (("FU1", y_top), ("FU2", y_bottom)):
        fu_bbox = draw.textbbox((0, 0), label, font=font_fu)
        fu_tw = fu_bbox[2] - fu_bbox[0]
        fu_th = fu_bbox[3] - fu_bbox[1]
        draw.text(
            (fu_cx - fu_tw // 2, y_line - box_h // 2 - fu_gap - fu_th),
            label,
            fill=line_color,
            font=font_fu,
        )

    _draw_fig1_left_res_label_pil(
        img,
        x_left - 46,
        y_mid,
        font_r=font_left_r,
        font_sub=font_left_sub,
        font_main=font_left,
    )

    return _fit_image_to_standard_canvas(img)


def akb_input_schematic_png_bytes() -> io.BytesIO:
    img = _load_akb_ready_schematic()
    if img is None:
        img = render_akb_input_schematic()
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _fmt_ohm(value: float) -> str:
    return f"{value:.3f}".replace(".", ",")


def _fmt_ohm4(value: float) -> str:
    return f"{value:.4f}".replace(".", ",")


def _draw_r_subscript_label_pil(
    draw: "ImageDraw.ImageDraw",
    cx: int,
    y: int,
    suffix: str,
    *,
    color: str,
    font_r: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    font_sub: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
) -> int:
    r_box = draw.textbbox((0, 0), "R", font=font_r)
    r_w = r_box[2] - r_box[0]
    r_h = r_box[3] - r_box[1]
    sub_box = draw.textbbox((0, 0), suffix, font=font_sub)
    sub_w = sub_box[2] - sub_box[0]
    sub_h = sub_box[3] - sub_box[1]
    total_w = r_w + sub_w + 2
    x0 = cx - total_w // 2
    draw.text((x0, y), "R", fill=color, font=font_r)
    draw.text((x0 + r_w + 2, y + r_h // 2 - sub_h // 2 + 3), suffix, fill=color, font=font_sub)
    return total_w


def _draw_rpr_caption_pil(
    draw: "ImageDraw.ImageDraw",
    cx: int,
    y: int,
    suffix: str,
    *,
    color: str,
    font_r: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    font_sub: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
) -> None:
    _draw_r_subscript_label_pil(
        draw, cx, y, suffix, color=color, font_r=font_r, font_sub=font_sub
    )


def _draw_left_res_label_pil(
    img: "Image.Image",
    draw: "ImageDraw.ImageDraw",
    x: int,
    y: int,
    *,
    color: str,
    font_r: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    font_sub: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    font_main: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    value_line: str | None = None,
) -> None:
    from PIL import Image, ImageDraw

    show_value = value_line is not None and value_line != ""

    tmp = Image.new("RGBA", (1, 1), (255, 255, 255, 0))
    td = ImageDraw.Draw(tmp)
    r_box = td.textbbox((0, 0), "R", font=font_r)
    sub_ab_box = td.textbbox((0, 0), "АБ", font=font_sub)
    plus_box = td.textbbox((0, 0), "+", font=font_main)
    sub_per_box = td.textbbox((0, 0), "ПЕР", font=font_sub)
    line2_box = td.textbbox((0, 0), value_line or "", font=font_main) if show_value else (0, 0, 0, 0)

    r_w = r_box[2] - r_box[0]
    r_h = r_box[3] - r_box[1]
    sub_ab_w = sub_ab_box[2] - sub_ab_box[0]
    sub_ab_h = sub_ab_box[3] - sub_ab_box[1]
    plus_w = plus_box[2] - plus_box[0]
    plus_h = plus_box[3] - plus_box[1]
    sub_per_w = sub_per_box[2] - sub_per_box[0]
    sub_per_h = sub_per_box[3] - sub_per_box[1]
    line2_w = line2_box[2] - line2_box[0]
    line2_h = line2_box[3] - line2_box[1]

    line1_w = r_w + 2 + sub_ab_w + 8 + plus_w + 8 + r_w + 2 + sub_per_w
    line1_h = max(r_h, plus_h, sub_ab_h + 12, sub_per_h + 12)
    sep_above = 4
    sep_below = 4
    sep_pad = 4
    value_block_h = (sep_above + 1 + sep_below + line2_h + 6) if show_value else 0
    block_w = (max(line1_w, line2_w) if show_value else line1_w) + 8
    block_h = line1_h + value_block_h

    block = Image.new("RGBA", (block_w, block_h), (255, 255, 255, 0))
    bd = ImageDraw.Draw(block)

    x0 = (block_w - line1_w) // 2
    y0 = 0
    bd.text((x0, y0), "R", fill=(0, 0, 0, 255), font=font_r)
    x0 += r_w + 2
    bd.text((x0, y0 + 12), "АБ", fill=(0, 0, 0, 255), font=font_sub)
    x0 += sub_ab_w + 8
    bd.text((x0, y0), "+", fill=(0, 0, 0, 255), font=font_main)
    x0 += plus_w + 8
    bd.text((x0, y0), "R", fill=(0, 0, 0, 255), font=font_r)
    x0 += r_w + 2
    bd.text((x0, y0 + 12), "ПЕР", fill=(0, 0, 0, 255), font=font_sub)

    if show_value:
        sep_w = max(line1_w, line2_w) + sep_pad * 2
        sep_x0 = (block_w - sep_w) // 2
        sep_y = line1_h + sep_above
        bd.line([(sep_x0, sep_y), (sep_x0 + sep_w, sep_y)], fill=(0, 0, 0, 255), width=1)
        line2_x = (block_w - line2_w) // 2
        line2_y = sep_y + 1 + sep_below
        bd.text((line2_x, line2_y), value_line or "", fill=(0, 0, 0, 255), font=font_main)

    bb = block.getbbox()
    if bb is not None:
        block = block.crop(bb)
    rot = block.rotate(90, expand=True)
    paste_x = max(2, min(img.width - rot.width - 2, x - rot.width // 2))
    paste_y = max(2, min(img.height - rot.height - 2, y - rot.height // 2))
    img.paste(rot, (paste_x, paste_y), rot)


def _draw_fig1_left_res_label_pil(
    img: "Image.Image",
    x: int,
    y: int,
    *,
    font_r: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    font_sub: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
    font_main: "ImageFont.FreeTypeFont | ImageFont.ImageFont",
) -> None:
    """Вертикальная подпись Rаб+Rпер только для рис.1 (без числовых значений)."""
    from PIL import Image, ImageDraw

    tmp = Image.new("RGBA", (1, 1), (255, 255, 255, 0))
    td = ImageDraw.Draw(tmp)
    r_box = td.textbbox((0, 0), "R", font=font_r)
    sub_ab_box = td.textbbox((0, 0), "АБ", font=font_sub)
    plus_box = td.textbbox((0, 0), "+", font=font_main)
    sub_per_box = td.textbbox((0, 0), "ПЕР", font=font_sub)

    r_w = r_box[2] - r_box[0]
    r_h = r_box[3] - r_box[1]
    sub_ab_w = sub_ab_box[2] - sub_ab_box[0]
    sub_ab_h = sub_ab_box[3] - sub_ab_box[1]
    plus_w = plus_box[2] - plus_box[0]
    plus_h = plus_box[3] - plus_box[1]
    sub_per_w = sub_per_box[2] - sub_per_box[0]
    sub_per_h = sub_per_box[3] - sub_per_box[1]

    line1_w = r_w + 2 + sub_ab_w + 8 + plus_w + 8 + r_w + 2 + sub_per_w
    line1_h = max(r_h, plus_h, sub_ab_h + 12, sub_per_h + 12)
    ink_pad = 8
    block_w = line1_w + ink_pad * 2
    block_h = line1_h + ink_pad * 2

    block = Image.new("RGBA", (block_w, block_h), (255, 255, 255, 0))
    bd = ImageDraw.Draw(block)
    x0 = (block_w - line1_w) // 2
    y0 = ink_pad
    bd.text((x0, y0), "R", fill=(0, 0, 0, 255), font=font_r)
    x0 += r_w + 2
    bd.text((x0, y0 + 12), "АБ", fill=(0, 0, 0, 255), font=font_sub)
    x0 += sub_ab_w + 8
    bd.text((x0, y0), "+", fill=(0, 0, 0, 255), font=font_main)
    x0 += plus_w + 8
    bd.text((x0, y0), "R", fill=(0, 0, 0, 255), font=font_r)
    x0 += r_w + 2
    bd.text((x0, y0 + 12), "ПЕР", fill=(0, 0, 0, 255), font=font_sub)

    bb = block.getbbox()
    if bb is not None:
        block = block.crop(bb)
    edge_pad = 10
    padded = Image.new(
        "RGBA",
        (block.width + edge_pad * 2, block.height + edge_pad * 2),
        (255, 255, 255, 0),
    )
    padded.paste(block, (edge_pad, edge_pad), block)
    rot = padded.rotate(90, expand=True)
    margin = 4
    paste_x = max(margin, min(img.width - rot.width - margin, x - rot.width // 2))
    paste_y = max(margin, min(img.height - rot.height - margin, y - rot.height // 2))
    img.paste(rot, (paste_x, paste_y), rot)


def render_k1_substitution_schematic(
    r_ab_ohm: float,
    r_per_ohm: float,
    r_pr_ohm: float,
    *,
    kz_point_label: str = "K1-1(2)",
) -> "Image.Image":
    return _render_k1_substitution_schematic_pil(r_ab_ohm, r_per_ohm, r_pr_ohm, kz_point_label=kz_point_label)


def _render_k1_substitution_schematic_pil(
    r_ab_ohm: float,
    r_per_ohm: float,
    r_pr_ohm: float,
    *,
    kz_point_label: str,
) -> "Image.Image":
    from PIL import Image, ImageDraw

    img = Image.new("RGB", (_FIG_STANDARD_W, _FIG_STANDARD_H), "white")
    draw = ImageDraw.Draw(img)
    line_color = "black"
    wire_w = 2
    box_w = 44
    box_h = 20

    font_r = _load_font(22, bold=False)
    font_r_sub = _load_font(12, bold=False)
    font_value = _load_font(22, bold=False)
    font_left = _load_font(20, bold=False)
    font_left_sub = _load_font(12, bold=False)
    font_kz = _load_font(16 if len(kz_point_label) > 2 else 22, bold=False)

    x_left = 92
    x_right = 452
    y_top = 78
    y_bottom = 228
    y_mid = (y_top + y_bottom) // 2

    draw.line([(x_left, y_top), (x_right, y_top)], fill=line_color, width=wire_w)
    draw.line([(x_right, y_top), (x_right, y_bottom)], fill=line_color, width=wire_w)
    draw.line([(x_right, y_bottom), (x_left, y_bottom)], fill=line_color, width=wire_w)
    draw.line([(x_left, y_bottom), (x_left, y_top)], fill=line_color, width=wire_w)

    left_rect_w = box_h
    left_rect_h = box_w
    left_rect_top = y_mid - left_rect_h // 2
    draw.rectangle(
        [x_left - left_rect_w // 2, left_rect_top, x_left + left_rect_w // 2, left_rect_top + left_rect_h],
        outline=line_color,
        width=2,
        fill="white",
    )

    fu1_x = 228
    draw.rectangle([fu1_x, y_top - box_h // 2, fu1_x + box_w, y_top + box_h // 2], outline=line_color, width=2, fill="white")
    fu2_x = fu1_x
    draw.rectangle(
        [fu2_x, y_bottom - box_h // 2, fu2_x + box_w, y_bottom + box_h // 2],
        outline=line_color,
        width=2,
        fill="white",
    )

    top_val = f"{_fmt_ohm(r_pr_ohm)} Ом"
    bottom_val = f"{_fmt_ohm(r_pr_ohm)} Ом"

    top_val_box = draw.textbbox((0, 0), top_val, font=font_value)
    top_val_w = top_val_box[2] - top_val_box[0]
    _draw_rpr_caption_pil(
        draw,
        fu1_x + box_w // 2,
        y_top - 74,
        "ПР.FU1",
        color=line_color,
        font_r=font_r,
        font_sub=font_r_sub,
    )
    draw.line(
        [(fu1_x + box_w // 2 - top_val_w // 2 - 4, y_top - 52), (fu1_x + box_w // 2 + top_val_w // 2 + 4, y_top - 52)],
        fill=line_color,
        width=1,
    )
    draw.text((fu1_x + box_w // 2 - top_val_w // 2, y_top - 46), top_val, fill=line_color, font=font_value)

    bottom_val_box = draw.textbbox((0, 0), bottom_val, font=font_value)
    bottom_val_w = bottom_val_box[2] - bottom_val_box[0]
    _draw_rpr_caption_pil(
        draw,
        fu2_x + box_w // 2,
        y_bottom - 74,
        "ПР.FU2",
        color=line_color,
        font_r=font_r,
        font_sub=font_r_sub,
    )
    draw.line(
        [(fu2_x + box_w // 2 - bottom_val_w // 2 - 4, y_bottom - 52), (fu2_x + box_w // 2 + bottom_val_w // 2 + 4, y_bottom - 52)],
        fill=line_color,
        width=1,
    )
    draw.text((fu2_x + box_w // 2 - bottom_val_w // 2, y_bottom - 46), bottom_val, fill=line_color, font=font_value)

    _draw_left_res_label_pil(
        img,
        draw,
        x_left - 56,
        y_mid,
        color=line_color,
        font_r=font_r,
        font_sub=font_left_sub,
        font_main=font_left,
        value_line=f"({_fmt_ohm4(r_ab_ohm)}+{_fmt_ohm4(r_per_ohm)}) Ом",
    )

    draw.text((x_right + 24, y_mid - 36), kz_point_label, fill=line_color, font=font_kz)
    icon = _load_lightning_icon()
    if icon is not None:
        target_h = 26
        target_w = max(10, int(icon.width * (target_h / max(icon.height, 1))))
        icon = icon.resize((target_w, target_h))
        img.paste(icon, (x_right + 4, y_mid - 12), icon)
    else:
        draw.line(
            [
                (x_right + 18, y_mid - 8),
                (x_right + 6, y_mid + 2),
                (x_right + 14, y_mid + 2),
                (x_right + 1, y_mid + 8),
            ],
            fill="#7a7a7a",
            width=1,
        )
    return img


def k1_substitution_schematic_png_bytes(
    r_ab_ohm: float,
    r_per_ohm: float,
    r_pr_ohm: float,
    *,
    kz_point_label: str = "K1-1(2)",
) -> io.BytesIO:
    img = render_k1_substitution_schematic(r_ab_ohm, r_per_ohm, r_pr_ohm, kz_point_label=kz_point_label)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def render_subsection_kz_schematic(items: list[SoptEquipmentItem], r_ab_ohm: float, r_per_ohm: float) -> "Image.Image":
    from PIL import Image, ImageDraw

    components: list[SoptEquipmentItem] = []
    kz_points: list[tuple[str, int]] = []
    for item in items:
        if item.item_type == "kz_point":
            kz_points.append((item.designation.strip() or f"KЗ-{len(kz_points) + 1}", len(components)))
        else:
            components.append(item)
    if not kz_points:
        kz_points.append(("K1", len(components)))

    count = max(1, len(components))
    step = 96
    box_w = 44
    box_h = 20
    x_left = 92
    x_start = 190
    x_positions = [x_start + i * step for i in range(count)]
    points_x: list[int] = []
    for _, after_idx in kz_points:
        if after_idx <= 0:
            points_x.append(x_start - 70)
        elif after_idx >= len(components):
            points_x.append(x_positions[-1] + 86)
        else:
            points_x.append((x_positions[after_idx - 1] + x_positions[after_idx]) // 2)
    x_right = max(points_x + [x_positions[-1] + 50])
    img = Image.new("RGB", (x_right + 120, 310), "white")
    draw = ImageDraw.Draw(img)
    y_top = 90
    y_bottom = 226
    y_mid = (y_top + y_bottom) // 2
    line_color = "black"
    wire_w = 2

    font_r_lbl = _load_font(16, bold=False)
    font_r_sub = _load_font(11, bold=False)
    font_val = _load_font(13, bold=False)
    font_kz = _load_font(16, bold=False)
    font_left_r = _load_font(22, bold=False)
    font_left = _load_font(20, bold=False)
    font_left_sub = _load_font(12, bold=False)

    draw.line([(x_left, y_top), (x_right, y_top)], fill=line_color, width=wire_w)
    draw.line([(x_right, y_top), (x_right, y_bottom)], fill=line_color, width=wire_w)
    draw.line([(x_right, y_bottom), (x_left, y_bottom)], fill=line_color, width=wire_w)
    draw.line([(x_left, y_bottom), (x_left, y_top)], fill=line_color, width=wire_w)

    left_rect_w = box_h
    left_rect_h = 44
    draw.rectangle(
        [x_left - left_rect_w // 2, y_mid - left_rect_h // 2, x_left + left_rect_w // 2, y_mid + left_rect_h // 2],
        outline=line_color,
        width=2,
        fill="white",
    )
    # Левая подпись и координаты как в схеме рис.2.
    _draw_left_res_label_pil(
        img,
        draw,
        x_left - 56,
        y_mid,
        color=line_color,
        font_r=font_left_r,
        font_sub=font_left_sub,
        font_main=font_left,
        value_line=f"({_fmt_ohm4(r_ab_ohm)}+{_fmt_ohm4(r_per_ohm)}) Ом",
    )

    for idx, item in enumerate(components):
        cx = x_positions[idx]
        draw.rectangle([cx - box_w // 2, y_top - box_h // 2, cx + box_w // 2, y_top + box_h // 2], outline=line_color, width=2, fill="white")
        draw.rectangle([cx - box_w // 2, y_bottom - box_h // 2, cx + box_w // 2, y_bottom + box_h // 2], outline=line_color, width=2, fill="white")
        dsg = item.designation.strip() or f"X{idx + 1}"
        r_text = f"{_fmt_ohm(item_resistance_ohm(item))} Ом"
        tw = _draw_r_subscript_label_pil(
            draw,
            cx,
            y_top - 56,
            dsg,
            color=line_color,
            font_r=font_r_lbl,
            font_sub=font_r_sub,
        )
        draw.line([(cx - tw // 2 - 2, y_top - 36), (cx + tw // 2 + 2, y_top - 36)], fill=line_color, width=1)
        vb = draw.textbbox((0, 0), r_text, font=font_val)
        vw = vb[2] - vb[0]
        draw.text((cx - vw // 2, y_top - 32), r_text, fill=line_color, font=font_val)

        tw = _draw_r_subscript_label_pil(
            draw,
            cx,
            y_bottom - 56,
            dsg,
            color=line_color,
            font_r=font_r_lbl,
            font_sub=font_r_sub,
        )
        draw.line([(cx - tw // 2 - 2, y_bottom - 36), (cx + tw // 2 + 2, y_bottom - 36)], fill=line_color, width=1)
        draw.text((cx - vw // 2, y_bottom - 32), r_text, fill=line_color, font=font_val)

    icon = _load_lightning_icon()
    for idx, ((label, _), xkz) in enumerate(zip(kz_points, points_x)):
        if xkz < x_right:
            draw.line([(xkz, y_top), (xkz, y_bottom)], fill="#808080", width=1)
        y_lbl = y_mid - 26 - idx * 4
        draw.text((xkz + 24, y_lbl), label, fill=line_color, font=font_kz)
        if icon is not None:
            ih = 22
            iw = max(9, int(icon.width * (ih / max(icon.height, 1))))
            bolt = icon.resize((iw, ih))
            img.paste(bolt, (xkz + 5, y_mid - 10), bolt)
        else:
            draw.line(
                [
                    (xkz + 18, y_mid - 8),
                    (xkz + 6, y_mid + 2),
                    (xkz + 14, y_mid + 2),
                    (xkz + 1, y_mid + 8),
                ],
                fill="#7a7a7a",
                width=1,
            )

    return img


def subsection_kz_schematic_png_bytes(items: list[SoptEquipmentItem], r_ab_ohm: float, r_per_ohm: float) -> io.BytesIO:
    img = render_subsection_kz_schematic(items, r_ab_ohm, r_per_ohm)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf
