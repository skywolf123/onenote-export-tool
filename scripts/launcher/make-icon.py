"""生成 OneNote 导出工具的应用图标（多尺寸 .ico）。

设计：紫色圆角方块（OneNote 的配色印象）+ 白色文档页 + 向下的导出箭头。
在 1024px 上绘制再降采样，保证 16px 下仍然认得出。

用法（icon.ico 已随仓库提供，只有重新设计图标时才需要跑）：
    uv run --with pillow python3 scripts/launcher/make-icon.py

输出写到本脚本所在目录下的 icon.ico，并顺带在 /tmp 留两张预览图。
"""
import os

from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_ICO = os.path.join(HERE, "icon.ico")

S = 1024  # 超采样画布

BG_TOP = (168, 85, 247)     # #A855F7
BG_BOTTOM = (109, 40, 217)  # #6D28D9
PAGE = (255, 255, 255)
LINE = (203, 213, 225)      # 页面上的文字线条
ARROW = (124, 58, 237)      # 箭头
ARROW_DARK = (91, 33, 182)


def lerp(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def vertical_gradient(size, top, bottom):
    img = Image.new("RGB", (1, size), top)
    px = img.load()
    for y in range(size):
        px[0, y] = lerp(top, bottom, y / (size - 1))
    return img.resize((size, size), Image.NEAREST)


def rounded_mask(size, radius):
    m = Image.new("L", (size, size), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


def build():
    # ── 背景：圆角方块 + 纵向渐变 ──
    bg = vertical_gradient(S, BG_TOP, BG_BOTTOM).convert("RGBA")
    bg.putalpha(rounded_mask(S, int(S * 0.22)))

    # ── 文档页：居中偏上，带一点投影 ──
    pw, ph = int(S * 0.52), int(S * 0.62)
    px0 = (S - pw) // 2
    py0 = int(S * 0.17)

    shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(
        [px0, py0 + int(S * 0.02), px0 + pw, py0 + ph + int(S * 0.02)],
        radius=int(S * 0.045), fill=(0, 0, 0, 90))
    shadow = shadow.filter(ImageFilter.GaussianBlur(S * 0.018))
    bg.alpha_composite(shadow)

    page = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(page).rounded_rectangle(
        [px0, py0, px0 + pw, py0 + ph], radius=int(S * 0.045), fill=PAGE)
    bg.alpha_composite(page)

    d = ImageDraw.Draw(bg)

    # ── 页首的两条文字线（示意 Markdown 正文）──
    line_x0 = px0 + int(pw * 0.16)
    line_x1 = px0 + int(pw * 0.84)
    for i, (yy, w) in enumerate([(0.14, 1.0), (0.24, 0.72)]):
        ly = py0 + int(ph * yy)
        d.rounded_rectangle([line_x0, ly, line_x0 + int((line_x1 - line_x0) * w), ly + int(S * 0.022)],
                            radius=int(S * 0.011), fill=LINE)

    # ── 导出箭头：竖杆 + 三角头，居中于页面下半部 ──
    cx = px0 + pw // 2
    stem_w = int(S * 0.075)
    stem_top = py0 + int(ph * 0.36)
    stem_bot = py0 + int(ph * 0.60)
    head_w = int(S * 0.20)
    head_top = stem_bot - int(S * 0.012)
    head_bot = py0 + int(ph * 0.83)

    d.rounded_rectangle([cx - stem_w // 2, stem_top, cx + stem_w // 2, stem_bot],
                        radius=stem_w // 3, fill=ARROW)
    d.polygon([(cx - head_w // 2, head_top), (cx + head_w // 2, head_top), (cx, head_bot)],
              fill=ARROW_DARK)

    # ── 箭头下方的收纳托盘：一条横线，强化"导出到某处"的语义 ──
    tray_y = py0 + int(ph * 0.90)
    tray_h = int(S * 0.026)
    d.rounded_rectangle([line_x0, tray_y, line_x1, tray_y + tray_h],
                        radius=tray_h // 2, fill=ARROW)

    return bg


def main():
    base = build()
    sizes = [256, 128, 64, 48, 32, 24, 16]
    frames = [base.resize((s, s), Image.LANCZOS) for s in sizes]

    frames[0].save(OUT_ICO, format="ICO",
                   sizes=[(s, s) for s in sizes], append_images=frames[1:])

    # 顺带留两张预览图，便于肉眼确认小尺寸下是否可辨
    base.resize((512, 512), Image.LANCZOS).save("/tmp/app-512.png")
    sheet = Image.new("RGBA", (16 + 24 + 32 + 48 + 64 + 5 * 16, 80), (240, 240, 240, 255))
    x = 8
    with Image.open(OUT_ICO) as im:
        for s in sizes[::-1][:5]:
            sheet.alpha_composite(im.ico.getimage((s, s)).convert("RGBA"), (x, 40 - s // 2))
            x += s + 16
    sheet.resize((sheet.width * 3, sheet.height * 3), Image.NEAREST).save("/tmp/app-sizes.png")

    # 校验 ico 里确实含全部尺寸
    with Image.open(OUT_ICO) as im:
        print("已生成:", OUT_ICO)
        print("ico 尺寸:", sorted(im.ico.sizes()))


if __name__ == "__main__":
    main()
