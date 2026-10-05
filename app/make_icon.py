# -*- coding: utf-8 -*-
"""產生應用程式圖示 app.ico（多尺寸）。只有改圖示時才需要重跑。

需要：pip install pillow
"""

from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent / "app.ico"
SIZES = [16, 24, 32, 48, 64, 128, 256]

BG_TOP = (0, 48, 95)        # 深藍
BG_BOTTOM = (0, 86, 158)
BAR = (255, 255, 255)
ACCENT = (219, 0, 17)       # 紅色重點，呼應報表的紅色負數


def draw_icon(size: int) -> Image.Image:
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # 圓角底色＋由上到下的漸層
    radius = int(size * 0.18)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=BG_TOP)
    for y in range(size):
        ratio = y / max(size - 1, 1)
        color = tuple(
            int(BG_TOP[i] + (BG_BOTTOM[i] - BG_TOP[i]) * ratio) for i in range(3)
        )
        draw.line([(0, y), (size, y)], fill=color + (255,))

    # 再蓋一次圓角，把漸層裁成圓角矩形
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, size - 1, size - 1], radius=radius, fill=255
    )
    image.putalpha(mask)
    draw = ImageDraw.Draw(image)

    # 長條圖：三根白柱 + 一根紅柱（代表每日損益）
    margin = size * 0.18
    base = size - margin
    bar_width = size * 0.12
    gap = size * 0.07
    heights = [0.30, 0.46, 0.62, 0.80]
    colors = [BAR, BAR, BAR, ACCENT]

    total = len(heights) * bar_width + (len(heights) - 1) * gap
    x = (size - total) / 2
    for height_ratio, color in zip(heights, colors):
        top = base - (base - margin) * height_ratio
        box = [x, top, x + bar_width, base]
        radius = int(bar_width * 0.25)
        # 小尺寸下柱子只有 1~2 px 寬，圓角會算出負寬度，直接畫方角
        if radius >= 1 and bar_width > 2 * radius + 2:
            draw.rounded_rectangle(box, radius=radius, fill=color + (255,))
        else:
            draw.rectangle(box, fill=color + (255,))
        x += bar_width + gap

    return image


def main():
    # 從最大尺寸存，PIL 會自己縮出 ICO 需要的各種尺寸
    # （不能從 16x16 存，否則每個尺寸都會被壓成 16x16）
    master = draw_icon(max(SIZES))
    master.save(OUT, format="ICO", sizes=[(s, s) for s in SIZES])
    print(f"已產生圖示：{OUT}")


if __name__ == "__main__":
    main()
