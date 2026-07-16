# -*- coding: utf-8 -*-
"""把 AI 生成的图标 PNG 转成多尺寸、透明背景的 Windows .ico。

用法:
    python convert_theme_icon.py <input_png> <output_ico> [padding]
- 自动生成 16/32/48/64/128/256 六档尺寸（Windows 标准）
- 保持 PNG 的透明通道；非正方形输入会居中贴到透明画布
- padding(可选, 0~0.3): 四周留白比例, 避免图标贴边
"""
import sys
import os
from PIL import Image

SIZES = [16, 32, 48, 64, 128, 256]


def to_ico(input_png, output_ico, padding=0.0):
    src = Image.open(input_png).convert("RGBA")
    w, h = src.size
    # 非正方形 -> 居中到正方形透明画布
    if w != h:
        side = max(w, h)
        canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        canvas.paste(src, ((side - w) // 2, (side - h) // 2), src)
        src = canvas
    # 可选留白
    if padding > 0:
        side = src.size[0]
        inner = int(side * (1 - 2 * padding))
        inner_img = src.resize((inner, inner), Image.LANCZOS)
        canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        canvas.paste(inner_img, ((side - inner) // 2, (side - inner) // 2), inner_img)
        src = canvas
    frames = []
    for s in SIZES:
        frames.append(src.resize((s, s), Image.LANCZOS))
    os.makedirs(os.path.dirname(os.path.abspath(output_ico)), exist_ok=True)
    frames[0].save(
        output_ico,
        format="ICO",
        sizes=[(s, s) for s in SIZES],
        append_images=frames[1:],
    )
    print(f"[OK] {output_ico}  ({SIZES} px, transparent)")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: python convert_theme_icon.py <input_png> <output_ico> [padding]")
        sys.exit(1)
    pad = float(sys.argv[3]) if len(sys.argv) > 3 else 0.08
    to_ico(sys.argv[1], sys.argv[2], pad)
