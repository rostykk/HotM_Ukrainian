"""Капсула й банер українізатора на основі офіційних артів гри.

Вхід (покладіть поруч у dist_template/art_src/):
  capsule_kr.webp — корейська капсула 920×430 (логотип піднятий, під ним корейська назва)
  keyart.webp     — ключовий арт 2000×646 (місто на заході, розбите серце)
Вихід: capsule_uk.png, banner_uk.png у dist_template/.
"""
import os
import sys

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "art_src")
FONTS = r"C:\Windows\Fonts"
WHITE = (244, 246, 250)
RED = (230, 40, 48)
UA_BLUE, UA_YELLOW = (0, 87, 183), (255, 213, 0)


def font(name, size, axes=None):
    f = ImageFont.truetype(os.path.join(FONTS, name), size)
    if axes:
        try:
            f.set_variation_by_axes(axes)
        except Exception:
            pass
    return f


def spaced(draw, text, fnt, spacing):
    return sum(draw.textlength(c, font=fnt) for c in text) + spacing * (len(text) - 1)


def draw_spaced(img, xy, text, fnt, fill, spacing, shadow=True):
    """Текст із розрядкою + м'яка темна тінь, як у підписів на капсулах."""
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    x, y = xy
    for c in text:
        d.text((x, y), c, font=fnt, fill=fill + (255,))
        x += d.textlength(c, font=fnt) + spacing
    if shadow:
        sh = Image.new("RGBA", img.size, (0, 0, 0, 0))
        sh.putalpha(layer.split()[3].filter(ImageFilter.GaussianBlur(4)).point(lambda v: int(v * 0.9)))
        img.alpha_composite(sh, (2, 3))
    img.alpha_composite(layer)


def extract_logo(capsule):
    """Логотип HEART / OF THE MACHINE з корейської капсули (білий і червоний за кольором)."""
    a = np.asarray(capsule.crop((288, 222, 642, 364))).astype(float)
    R, G, B = a[..., 0], a[..., 1], a[..., 2]
    mn, mx = a.min(axis=2), a.max(axis=2)
    sat = (mx - mn) / np.maximum(mx, 1)
    w_alpha = np.clip((mn - 150) / 60, 0, 1) * np.clip((0.35 - sat) / 0.2, 0, 1)
    r_alpha = np.clip((R - G - 90) / 50, 0, 1) * np.clip((R - 140) / 40, 0, 1) * np.clip((B - G + 40) / 40, 0, 1)
    alpha = np.maximum(w_alpha, r_alpha)
    return Image.fromarray(np.dstack([a, alpha * 255]).astype(np.uint8), "RGBA")


def erase_region(img, box, luma_threshold=70):
    """Стерти світлий текст на темному тлі: світлі пікселі замінюємо оцінкою фону."""
    region = img.crop(box).convert("RGB")
    a = np.asarray(region).astype(float)
    luma = a @ np.array([0.299, 0.587, 0.114])
    mask = Image.fromarray(((luma > luma_threshold) * 255).astype(np.uint8)).filter(ImageFilter.MaxFilter(7))
    bg = region.filter(ImageFilter.MinFilter(15)).filter(ImageFilter.GaussianBlur(6))
    region.paste(bg, (0, 0), mask.filter(ImageFilter.GaussianBlur(1.5)))
    img.paste(region, box[:2])


def make_capsule(variant="bahn"):
    cap = Image.open(os.path.join(SRC, "capsule_kr.webp")).convert("RGB")
    erase_region(cap, (300, 360, 640, 418))
    img = cap.convert("RGBA")
    d = ImageDraw.Draw(img)
    text = "СЕРЦЕ МАШИНИ"
    if variant == "bahn":
        f, sp = font("bahnschrift.ttf", 44, [700, 100]), 4
    else:
        f, sp = font("seguibl.ttf", 36), 3
    w = spaced(d, text, f, sp)
    cx = 465  # центр логотипа
    bbox = d.textbbox((0, 0), text, font=f)
    y = 386 - (bbox[1] + bbox[3]) / 2
    draw_spaced(img, (cx - w / 2, y), text, f, WHITE, sp)
    return img


def make_banner():
    key = Image.open(os.path.join(SRC, "keyart.webp")).convert("RGB")
    cap = Image.open(os.path.join(SRC, "capsule_kr.webp")).convert("RGB")
    W, H = 1100, 355
    img = key.resize((W, int(key.height * W / key.width)), Image.LANCZOS).convert("RGBA")
    H = img.height
    logo = extract_logo(cap)
    # легке затемнення під текстом ліворуч і праворуч, щоб читалося
    shade = Image.new("RGBA", img.size, (0, 0, 0, 0))
    sd = ImageDraw.Draw(shade)
    for x in range(W):
        k = max(0.0, 1 - abs((x - W / 2) / (W / 2)) * 1.0)  # 0 по краях, 1 в центрі
        a = int(110 * (1 - k) ** 1.4)
        sd.line([(x, H * 0.45), (x, H)], fill=(4, 14, 28, a))
    img.alpha_composite(shade.filter(ImageFilter.GaussianBlur(20)))
    lx, ly = 40, H - logo.height - 34
    glow = Image.new("RGBA", img.size, (0, 0, 0, 0))
    gl = Image.new("RGBA", logo.size, (0, 0, 0, 255))
    gl.putalpha(logo.split()[3].filter(ImageFilter.GaussianBlur(6)))
    glow.alpha_composite(gl, (lx + 2, ly + 3))
    img.alpha_composite(glow)
    img.alpha_composite(logo, (lx, ly))
    d = ImageDraw.Draw(img)
    # праворуч: українська назва
    right = W - 40
    f1 = font("bahnschrift.ttf", 40, [700, 100])
    f2 = font("bahnschrift.ttf", 22, [600, 75])
    t1, t2 = "СЕРЦЕ МАШИНИ", "УКРАЇНІЗАТОР · ПОВНИЙ ПЕРЕКЛАД"
    w1, w2 = spaced(d, t1, f1, 4), spaced(d, t2, f2, 5)
    y1 = H - 34 - 92
    draw_spaced(img, (right - w1, y1), t1, f1, WHITE, 4)
    bar_w = 120
    d = ImageDraw.Draw(img)
    d.rectangle([right - bar_w, y1 + 56, right, y1 + 60], fill=UA_BLUE)
    d.rectangle([right - bar_w, y1 + 61, right, y1 + 65], fill=UA_YELLOW)
    draw_spaced(img, (right - w2, y1 + 72), t2, f2, (200, 222, 240), 5)
    return img


def make_preview(capsule):
    """Квадратне прев'ю посібника: вирізка капсули навколо логотипа."""
    side = capsule.height
    left = max(0, min(capsule.width - side, 465 - side // 2))
    return capsule.crop((left, 0, left + side, side)).resize((512, 512), Image.LANCZOS)


if __name__ == "__main__":
    caps = {v: make_capsule(v).convert("RGB") for v in ("segoe", "bahn")}
    caps["segoe"].save(os.path.join(HERE, "capsule_uk.png"))
    caps["bahn"].save(os.path.join(HERE, "capsule_uk_alt.png"))
    make_preview(caps["segoe"]).save(os.path.join(HERE, "preview_uk.png"))
    make_banner().convert("RGB").save(os.path.join(HERE, "banner_uk.png"))
    print("ok")
