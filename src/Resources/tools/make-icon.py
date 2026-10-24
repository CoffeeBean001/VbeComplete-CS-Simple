# 生成程序图标：src/app.ico（多尺寸）+ 预览用的 contact-sheet.png
#
# 为什么用脚本画而不是塞一张位图：图标在 16/20/24/32/256 下都要清楚，
# 位图缩放出来的 16px 会糊成一团。这里按"矢量意图"在每个尺寸上重画，
# 小尺寸自动加粗笔画、去掉装饰，保证托盘那一格也认得出来。
#
# 用法（Pillow 在隔离环境里）：
#   python make-icon.py
#
# 图形：深蓝圆角方块 + 白色粗体 V + 下方一条"光标"横条（终端提示符的意象）。

import os
from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_ICO = os.path.join(HERE, "..", "app.ico")
OUT_SHEET = os.path.join(HERE, "..", "icon-preview.png")

S = 1024                      # 主画布
BLUE_TOP = (59, 130, 246)     # #3B82F6
BLUE_BOT = (29, 78, 216)      # #1D4ED8
WHITE = (255, 255, 255, 255)


def rounded_mask(size, radius):
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


def vertical_gradient(size, stops):
    """stops = [(t, rgb), ...]，t 从 0(顶) 到 1(底)。多段插值，接缝不可见。"""
    g = Image.new("RGB", (1, size))
    for y in range(size):
        t = y / max(1, size - 1)
        lo, hi = stops[0], stops[-1]
        for i in range(len(stops) - 1):
            if stops[i][0] <= t <= stops[i + 1][0]:
                lo, hi = stops[i], stops[i + 1]
                break
        span = max(1e-6, hi[0] - lo[0])
        u = min(1.0, max(0.0, (t - lo[0]) / span))
        g.putpixel((0, y), tuple(round(lo[1][c] + (hi[1][c] - lo[1][c]) * u) for c in range(3)))
    return g.resize((size, size), Image.BILINEAR)


def draw_master(size=S, fat=1.0):
    """在 size×size 上按比例重画。fat 用来给小尺寸加粗笔画。"""
    k = size / 1024.0
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    # 底：圆角方块（半径按比例，小尺寸略收一点，避免看起来像圆形）
    radius = int(230 * k) if size >= 64 else int(190 * k)
    body = vertical_gradient(size, [
        (0.00, (96, 165, 250)),   # #60A5FA 顶部提亮
        (0.42, (59, 130, 246)),   # #3B82F6
        (1.00, (30, 64, 175)),    # #1E40AF 底部压深
    ]).convert("RGBA")
    img.paste(body, (0, 0), rounded_mask(size, radius))

    # 左上角的柔光：用一个径向渐变当蒙版，别用硬边矩形（缩到 32px 会看到一条横线）
    if size >= 32:
        glow = Image.new("L", (size, size), 0)
        ImageDraw.Draw(glow).ellipse(
            [-int(360 * k), -int(430 * k), int(780 * k), int(560 * k)], fill=64)
        glow = glow.filter(ImageFilter.GaussianBlur(max(1, int(90 * k))))
        white = Image.new("RGBA", (size, size), (255, 255, 255, 255))
        img = Image.composite(white, img, glow.point(lambda v: int(v * 0.55)))

    d = ImageDraw.Draw(img)

    # 白色 V：一个六边形一次画完（两条竖臂 + 底部尖），不依赖字体。
    # ★不能"两条平行四边形拼起来"：接缝会留在受光面上，缩到 16px 就变成 Y。
    top_y = 236 * k
    bot_y = 790 * k
    left_x = 176 * k
    right_x = 848 * k
    mid_x = 512 * k
    arm = 176 * k * fat          # 每条臂在水平方向上的宽度
    rise = arm * 1.15            # 内侧缺口比外侧尖高出多少

    d.polygon([
        (left_x, top_y),                     # 左上外
        (left_x + arm, top_y),               # 左上内
        (mid_x, bot_y - rise),               # 内侧尖（缺口底）
        (right_x - arm, top_y),              # 右上内
        (right_x, top_y),                    # 右上外
        (mid_x, bot_y),                      # 外侧尖
    ], fill=WHITE)

    # 极轻微的内描边，让方块在浅色任务栏上也有边界
    if size >= 32:
        edge = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        ImageDraw.Draw(edge).rounded_rectangle(
            [1, 1, size - 2, size - 2], radius=radius,
            outline=(255, 255, 255, 40), width=max(1, int(2 * k)))
        img = Image.alpha_composite(img, edge)

    return img


def main():
    sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
    frames = []
    for s in sizes:
        # 小尺寸加粗一点点，补偿缩放损失
        fat = 1.0 if s >= 48 else 1.10
        big = draw_master(s * 4, fat=fat)
        # 4 倍超采样再缩回，边缘干净（比直接画小图锐利）
        frames.append(big.resize((s, s), Image.LANCZOS))

    ico_path = os.path.normpath(OUT_ICO)
    frames[-1].save(ico_path, format="ICO", sizes=[(s, s) for s in sizes],
                    append_images=frames[:-1])
    print("wrote", ico_path, os.path.getsize(ico_path), "bytes")

    # 预览：各尺寸平铺 + 深/浅两种底色，方便肉眼检查
    pad = 16
    cell = 280
    sheet = Image.new("RGBA", (cell * len(sizes) + pad * 2, cell * 2 + pad * 3), (240, 242, 245, 255))
    for i, (s, fr) in enumerate(zip(sizes, frames)):
        for row, bg in enumerate([(240, 242, 245, 255), (32, 34, 38, 255)]):
            x = pad + i * cell
            y = pad + row * (cell + pad)
            tile = Image.new("RGBA", (cell, cell), bg)
            show = fr.resize((min(s * (cell // 64 - 1), 240),) * 2, Image.NEAREST) if s <= 64 else fr
            tile.alpha_composite(show, ((cell - show.width) // 2, (cell - show.height) // 2))
            sheet.alpha_composite(tile, (x, y))
    sheet.save(os.path.normpath(OUT_SHEET))
    print("wrote", os.path.normpath(OUT_SHEET))


if __name__ == "__main__":
    main()
