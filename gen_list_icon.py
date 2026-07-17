#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成清单套(B 公文夹对勾)图标资源：
   1) TodoList_清单打勾.ico —— 多尺寸透明(16/32/48/64/128/256)，用于系统标题栏+exe+任务栏
   2) title_清单.png        —— 40x40 透明，用于应用内深蓝标题栏 Label
"""
import os
from PIL import Image, ImageDraw

BASE = os.path.dirname(os.path.abspath(__file__))

# 配色（与稳重骨架色同族）
FOLDER_MAIN = (79, 129, 189)    # 钢蓝 #4F81BD
FOLDER_DARK = (44, 62, 80)      # 深蓝 #2C3E50
FOLDER_EDGE = (28, 40, 54)      # 描边
GOLD = (241, 196, 15)           # 金 #F1C40F


def draw_folder_check(size):
    """在 size x size 透明画布上绘制 公文夹+金勾，返回 PIL.Image(RGBA)。"""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size / 256.0  # 缩放因子

    # 公文夹主体（圆角矩形）
    body_x0, body_y0 = int(40 * s), int(96 * s)
    body_x1, body_y1 = int(216 * s), int(220 * s)
    r = max(2, int(18 * s))
    d.rounded_rectangle([body_x0, body_y0, body_x1, body_y1], radius=r,
                        fill=FOLDER_MAIN, outline=FOLDER_EDGE, width=max(1, int(3 * s)))

    # 夹口（顶部梯形标签）
    tab_x0, tab_y0 = int(40 * s), int(70 * s)
    tab_x1, tab_y1 = int(132 * s), int(108 * s)
    tab_mid = int(120 * s)
    d.polygon([(tab_x0, tab_y1), (tab_x0, tab_y0), (tab_mid, tab_y0),
               (tab_x1, tab_y1)], fill=FOLDER_DARK, outline=FOLDER_EDGE)

    # 金色对勾（粗线，圆角端点）—— 夹面中部偏左
    lw = max(3, int(15 * s))
    cx0, cy0 = int(104 * s), int(150 * s)
    cx1, cy1 = int(140 * s), int(188 * s)
    cx2, cy2 = int(196 * s), int(120 * s)
    d.line([(cx0, cy0), (cx1, cy1), (cx2, cy2)], fill=GOLD,
           width=lw, joint="curve")
    return img


def make_ico(path):
    sizes = [16, 32, 48, 64, 128, 256]
    frames = [draw_folder_check(sz).resize((sz, sz), Image.LANCZOS) for sz in sizes]
    frames[0].save(path, sizes=[(sz, sz) for sz in sizes],
                   format="ICO", append_images=frames[1:])
    print(f"ico 已生成: {path} ({len(frames)} 尺寸)")


def make_png(path, size=40):
    img = draw_folder_check(size)
    img.save(path, format="PNG")
    print(f"png 已生成: {path} ({size}x{size})")


if __name__ == "__main__":
    ico_path = os.path.join(BASE, "TodoList_Themes", "themes", "清单打勾", "TodoList_清单打勾.ico")
    png_path = os.path.join(BASE, "TodoList_Themes", "themes", "清单打勾", "title_清单.png")
    make_ico(ico_path)
    make_png(png_path, 40)
