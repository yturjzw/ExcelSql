# -*- coding: utf-8 -*-
"""生成与网页版 .brand-mark 风格一致的应用图标 app.ico。

网页图标（frontend/index.html）：
  近黑圆角方块（#16181d，圆角 6/30）内嵌白色描边图形：
    左：表格外框（圆角矩形 rect x3 y4 w13 h16 rx1）+ 三条行线
    右：">" 箭头（M18 13 l1.5 2.5 L18 18）
  均 stroke=白色、fill=none、stroke-width=1.6、viewBox 0 0 24 24。

关键：Windows 任务栏用的是 16/24 px 小档图标。若用大图（1024）LANCZOS
缩小，1.6 的细描边会被抗锯齿稀释成灰白，任务栏上看起来"细小无力"。
因此每个尺寸都用「超采样矢量绘制 + 仅一步下采样」独立渲染，并按尺寸
加粗描边，保证小档也有足量纯白像素。
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
OUT = HERE / "assets" / "app.ico"

BG = (22, 24, 29, 255)      # #16181d（网页 --accent）
FG = (255, 255, 255, 255)   # 白色描边

SIZES = [256, 128, 64, 48, 32, 24, 16]

# 白色图形占方块边长的比例（已放大，配合任务栏小档更饱满）
GLYPH_SCALE = 0.72

# 超采样倍数：先在 N 倍画布上矢量绘制，再一步缩到目标尺寸，兼顾抗锯齿
OVERSAMPLE = 4


def _stroke_factor(size: int) -> float:
    """按最终尺寸加粗描边：尺寸越小，线相对越粗，保证小档不糊。"""
    if size <= 16:
        return 2.4
    if size <= 24:
        return 2.0
    if size <= 32:
        return 1.6
    if size <= 48:
        return 1.3
    if size <= 64:
        return 1.1
    return 1.0


def draw_glyph(size: int) -> Image.Image:
    """独立渲染一个目标尺寸（含超采样），返回该尺寸的 RGBA 图。"""
    S = size * OVERSAMPLE
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 背景：近黑圆角方块（圆角 0.20，同网页）
    d.rounded_rectangle([0, 0, S - 1, S - 1], radius=S * 0.20, fill=BG)

    # 白色图形：viewBox 0 0 24 24，居中，占比 GLYPH_SCALE
    g = GLYPH_SCALE * S
    ox = (S - g) / 2.0
    oy = (S - g) / 2.0
    k = g / 24.0
    sw = max(1.0, 1.6 * k * _stroke_factor(size))

    def P(x: float, y: float):
        return (ox + x * k, oy + y * k)

    # 表格外框（rounded rect）
    d.rounded_rectangle(
        [P(3, 4), P(16, 20)], radius=1.0 * k, outline=FG, width=int(round(sw))
    )
    # 三条行线
    for x1, y1, x2, y2 in [(6, 8, 13, 8), (6, 12, 13, 12), (6, 16, 10, 16)]:
        d.line([P(x1, y1), P(x2, y2)], fill=FG, width=int(round(sw)))
    # 右侧 ">" 箭头
    d.line(
        [P(18, 13), P(19.5, 15.5), P(18, 18)],
        fill=FG,
        width=int(round(sw)),
        joint="curve",
    )

    if OVERSAMPLE == 1:
        return img
    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    images = [draw_glyph(s) for s in SIZES]
    images[0].save(
        OUT,
        format="ICO",
        append_images=images[1:],
        sizes=[(s, s) for s in SIZES],
    )
    print(f"written: {OUT}  ({OUT.stat().st_size} bytes)")

    # 导出各尺寸 PNG 便于肉眼/量化核对
    for s, im in zip(SIZES, images):
        im.save(OUT.parent / f"preview_{s}.png", format="PNG")
    print(f"previews: {OUT.parent}/preview_*.png")


if __name__ == "__main__":
    main()