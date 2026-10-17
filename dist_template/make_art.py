"""Банер і лого українізатора в стилі Heart of the Machine (Pillow)."""
import math
import os
import random

from PIL import Image, ImageDraw, ImageFilter, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS = r"C:\Windows\Fonts"
NAVY_TOP, NAVY_BOTTOM = (6, 16, 30), (12, 34, 58)
CYAN = (70, 200, 255)
RED = (226, 36, 43)
UA_BLUE, UA_YELLOW = (0, 87, 183), (255, 213, 0)
WHITE = (240, 246, 252)


def font(size, weight=700, width=75, name="bahnschrift.ttf"):
    f = ImageFont.truetype(os.path.join(FONTS, name), size)
    try:
        f.set_variation_by_axes([weight, width])
    except Exception:
        pass
    return f


def gradient(w, h, top, bottom):
    img = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / max(1, h - 1)
        d.line([(0, y), (w, y)], fill=tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return img


def lattice(w, h, step, jitter, rnd):
    """Трикутна ґратка з розкиданими вузлами -> список трикутників."""
    rows = int(h / (step * 0.866)) + 3
    cols = int(w / step) + 3
    pts = {}
    for j in range(-1, rows):
        for i in range(-1, cols):
            x = i * step + (step / 2 if j % 2 else 0) + rnd.uniform(-jitter, jitter)
            y = j * step * 0.866 + rnd.uniform(-jitter, jitter)
            pts[(i, j)] = (x, y)
    tris = []
    for j in range(-1, rows - 1):
        for i in range(-1, cols - 1):
            if j % 2 == 0:
                a, b, c, d = (i, j), (i + 1, j), (i, j + 1), (i - 1, j + 1)
                tris.append((pts[a], pts[b], pts[c]))
                if d in pts:
                    tris.append((pts[a], pts[c], pts[d]))
            else:
                a, b, c, d = (i, j), (i + 1, j), (i + 1, j + 1), (i, j + 1)
                tris.append((pts[a], pts[b], pts[c]))
                tris.append((pts[a], pts[c], pts[d]))
    return tris


def wire_background(img, rnd, density=70, alpha_range=(18, 70), mask=None):
    w, h = img.size
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    for tri in lattice(w, h, density, density * 0.38, rnd):
        cx = sum(p[0] for p in tri) / 3
        cy = sum(p[1] for p in tri) / 3
        k = 1.0 if mask is None else mask(cx / w, cy / h)
        if k <= 0:
            continue
        a = int(rnd.uniform(*alpha_range) * k)
        if rnd.random() < 0.18:
            d.polygon(tri, fill=(CYAN[0], CYAN[1], CYAN[2], int(a * 0.35)))
        d.line(list(tri) + [tri[0]], fill=(CYAN[0], CYAN[1], CYAN[2], a), width=1)
        if rnd.random() < 0.12:
            x, y = tri[0]
            d.ellipse([x - 1.6, y - 1.6, x + 1.6, y + 1.6], fill=(170, 235, 255, min(255, a * 3)))
    glow = layer.filter(ImageFilter.GaussianBlur(3))
    img.alpha_composite(glow)
    img.alpha_composite(layer)


def heart_polygon(size, scale=0.92, n=360):
    """Параметрична крива серця, вписана в квадрат size×size."""
    pts = []
    for i in range(n):
        t = 2 * math.pi * i / n
        x = 16 * math.sin(t) ** 3
        y = 13 * math.cos(t) - 5 * math.cos(2 * t) - 2 * math.cos(3 * t) - math.cos(4 * t)
        pts.append((x, y))
    k = size * scale / 34.0
    cx, cy = size / 2, size / 2 - 1.5 * k
    return [(cx + x * k, cy - y * k) for x, y in pts]


def lowpoly_heart(size, rnd, step=None):
    """Низькополігональне серце: чіткий контур, верх синій, низ жовтий, бірюзові ребра."""
    step = step or size / 9.5
    outline = heart_polygon(size)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).polygon(outline, fill=255)
    split = size * 0.50  # рівна межа прапора

    def layer(base):
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        r = random.Random(rnd.random())
        for tri in lattice(size, size, step, step * 0.32, r):
            cx = sum(p[0] for p in tri) / 3
            cy = sum(p[1] for p in tri) / 3
            light = r.uniform(-0.20, 0.20) - (cx / size - 0.5) * 0.25 - (cy / size - 0.5) * 0.15
            if light > 0:
                col = tuple(int(c + (255 - c) * light) for c in base)
            else:
                col = tuple(int(c * (1 + light)) for c in base)
            d.polygon(tri, fill=col + (255,))
            d.line(list(tri) + [tri[0]], fill=(170, 230, 255, 70), width=max(1, size // 260))
        return img, r

    blue, r1 = layer(UA_BLUE)
    yellow, _ = layer(UA_YELLOW)
    flag_mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(flag_mask).rectangle([0, split, size, size], fill=255)
    body = Image.composite(yellow, blue, flag_mask)
    heart = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    heart.paste(body, (0, 0), mask)
    # світіння й контур
    glow = Image.new("RGBA", (size, size), (60, 170, 255, 0))
    glow.putalpha(mask.filter(ImageFilter.GaussianBlur(size / 22)).point(lambda v: int(v * 0.6)))
    edge = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(edge).line(outline + [outline[0]], fill=(190, 240, 255, 230), width=max(2, size // 110))
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.alpha_composite(glow)
    out.alpha_composite(heart)
    out.alpha_composite(edge)
    return out


def spaced_text(draw, xy, text, fnt, fill, spacing):
    x, y = xy
    for ch in text:
        draw.text((x, y), ch, font=fnt, fill=fill)
        x += draw.textlength(ch, font=fnt) + spacing
    return x


def spaced_width(draw, text, fnt, spacing):
    return sum(draw.textlength(ch, font=fnt) for ch in text) + spacing * (len(text) - 1)


def glitch_text(base, xy, text, fnt, rnd, slices=2):
    """Білий текст із хроматичним зсувом і горизонтальними зрізами, як у лого гри."""
    w, h = base.size
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).text(xy, text, font=fnt, fill=WHITE + (255,))
    red = Image.new("RGBA", base.size, (0, 0, 0, 0))
    ImageDraw.Draw(red).text((xy[0] - 3, xy[1]), text, font=fnt, fill=(255, 40, 60, 140))
    cyan = Image.new("RGBA", base.size, (0, 0, 0, 0))
    ImageDraw.Draw(cyan).text((xy[0] + 3, xy[1] + 1), text, font=fnt, fill=(40, 220, 255, 130))
    bbox = layer.getbbox()
    for _ in range(slices):
        y0 = rnd.randint(bbox[1] + 8, bbox[3] - 10)
        hh = rnd.randint(2, 5)
        dx = rnd.choice([-1, 1]) * rnd.randint(3, 8)
        band = layer.crop((0, y0, w, y0 + hh))
        layer.paste((0, 0, 0, 0), (0, y0, w, y0 + hh))
        layer.paste(band, (dx, y0), band)
    base.alpha_composite(red)
    base.alpha_composite(cyan)
    base.alpha_composite(layer)
    return bbox


def ua_bar(draw, x, y, w, h):
    draw.rectangle([x, y, x + w, y + h // 2], fill=UA_BLUE)
    draw.rectangle([x, y + h // 2, x + w, y + h], fill=UA_YELLOW)


def make_banner(path, w=1280, h=420, seed=7):
    rnd = random.Random(seed)
    img = gradient(w, h, NAVY_TOP, NAVY_BOTTOM).convert("RGBA")
    wire_background(img, rnd, density=74,
                    mask=lambda u, v: max(0.0, min(1.0, 0.25 + abs(u - 0.5) * 1.9)))
    heart = lowpoly_heart(660, random.Random(seed + 1)).resize((330, 330), Image.LANCZOS)
    img.alpha_composite(heart, (w - 330 - 60, (h - 330) // 2 + 6))
    d = ImageDraw.Draw(img)
    left = 70
    f_heart = font(150, 700, 75)
    glitch_text(img, (left - 4, 38), "HEART", f_heart, rnd)
    d = ImageDraw.Draw(img)
    f_mach = font(46, 700, 75)
    spaced_text(d, (left, 200), "OF THE MACHINE", f_mach, RED, 14)
    ua_bar(d, left, 274, 86, 10)
    f_ua = font(58, 600, 75)
    spaced_text(d, (left, 296), "УКРАЇНІЗАТОР", f_ua, WHITE, 8)
    f_sub = font(26, 400, 100)
    d.text((left + 2, 370), "повний переклад гри українською", font=f_sub, fill=(150, 200, 235))
    # тонка рамка-«інтерфейс»
    d.line([(24, 24), (180, 24)], fill=CYAN + (200,), width=2)
    d.line([(24, 24), (24, 90)], fill=CYAN + (200,), width=2)
    d.line([(w - 24, h - 24), (w - 180, h - 24)], fill=CYAN + (200,), width=2)
    d.line([(w - 24, h - 24), (w - 24, h - 90)], fill=CYAN + (200,), width=2)
    img.convert("RGB").save(path, optimize=True)


def make_logo(path, size=512, seed=11):
    rnd = random.Random(seed)
    img = gradient(size, size, NAVY_TOP, NAVY_BOTTOM).convert("RGBA")
    wire_background(img, rnd, density=58, alpha_range=(14, 55))
    heart = lowpoly_heart(660, random.Random(seed + 1)).resize((330, 330), Image.LANCZOS)
    img.alpha_composite(heart, ((size - 330) // 2, 26))
    d = ImageDraw.Draw(img)
    f_t = font(64, 700, 75)
    t = "HEART"
    tw = spaced_width(d, t, f_t, 6)
    spaced_text(d, ((size - tw) / 2, 352), t, f_t, WHITE, 6)
    f_m = font(26, 700, 75)
    m = "OF THE MACHINE"
    mw = spaced_width(d, m, f_m, 7)
    spaced_text(d, ((size - mw) / 2, 424), m, f_m, RED, 7)
    ua_bar(d, (size - 120) // 2, 466, 120, 10)
    img.convert("RGB").save(path, optimize=True)


if __name__ == "__main__":
    make_banner(os.path.join(HERE, "banner.png"))
    make_logo(os.path.join(HERE, "logo.png"))
    print("ok")
