# -*- coding: utf-8 -*-
"""
Вписывание номеров заказов в скан накладной (PIL, без нейросетей).
Портировано из vpisat_zakazy.py; шрифт подбирается переносимо (Windows/Linux).
"""
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

DARK = 170


def find_font():
    for c in (
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ):
        if Path(c).exists():
            return c
    return None


def _cluster(v, gap):
    if not v:
        return []
    b, s, p = [], v[0], v[0]
    for x in v[1:]:
        if x - p > gap:
            b.append((s, p))
            s = x
        p = x
    b.append((s, p))
    return b


def _horiz_proj_var(im):
    """Дисперсия горизонтальной проекции тёмных пикселей.

    У правильно ориентированного документа текст и строки таблицы идут
    горизонтально → плотность тёмных пикселей сильно меняется по строкам
    (большая дисперсия). У повёрнутого на 90° — почти равномерно по строкам.
    Максимум дисперсии = правильная ориентация (надёжнее, чем подсчёт «полос»).
    """
    a = np.array(im)
    h, w = a.shape
    dens = (a < DARK).sum(axis=1) / max(1, w)
    return float(dens.var())


def _auto_orient(im):
    scores = {a: _horiz_proj_var(im.rotate(a, expand=True)) for a in (0, 90, 180, 270)}
    best = max(scores, key=scores.get)
    return im.rotate(best, expand=True), best


def _deskew_angle(im):
    w = 500
    small = im.resize((w, max(1, int(im.height * w / im.width))), Image.LANCZOS)
    sw, sh = small.size

    def score(a):
        r = small.rotate(a, resample=Image.BICUBIC, fillcolor=255)
        px = list(r.getdata())
        rs = [sum(px[y * sw:(y + 1) * sw]) for y in range(sh)]
        m = sum(rs) / sh
        return sum((s - m) ** 2 for s in rs) / sh

    coarse = [round(i * 0.5, 2) for i in range(-12, 13)]
    bc = max(coarse, key=score)
    fine = [round(bc + i * 0.1, 2) for i in range(-5, 6)]
    return max(fine, key=score)


def _detect_rows(gray, n, skip=0):
    w, h = gray.size
    a = gray.load()
    x0, x1 = int(w * 0.06), int(w * 0.94)
    N = x1 - x0
    prof = [sum(1 for x in range(x0, x1) if a[x, y] < DARK) / N for y in range(h)]
    thr = max(0.13, max(prof) * 0.6)
    lines = _cluster([y for y in range(h) if prof[y] >= thr], 4)
    centers = [round((s + e) / 2) for s, e in lines]
    rows = [(centers[i], centers[i + 1]) for i in range(len(centers) - 1)
            if centers[i + 1] - centers[i] >= 28]
    return rows[skip:skip + n]


def _detect_kod_col(gray):
    w, h = gray.size
    a = gray.load()
    y0, y1 = int(h * 0.25), int(h * 0.75)
    M = y1 - y0
    cols = _cluster([x for x in range(w)
                     if sum(1 for y in range(y0, y1) if a[x, y] < DARK) / M >= 0.4], 6)
    centers = [round((s + e) / 2) for s, e in cols]
    if len(centers) < 2:
        return int(w * 0.06), int(w * 0.13)
    return centers[0] + 3, centers[1] - 5


def annotate(scan_path, positions, out_path, angle=None):
    """positions: список (y_center, text); y_center=None или text='' — пропустить.

    angle — уже определённая при OCR ориентация (0/90/180/270). Если None —
    определяется заново через _auto_orient (совместимость со старыми вызовами).
    """
    img = Image.open(scan_path).convert("RGB")
    gray = img.convert("L")

    if angle is None:
        gray2, ang = _auto_orient(gray)
    else:
        ang = angle
        gray2 = gray.rotate(ang, expand=True)
    rgb2 = img.rotate(ang, expand=True)

    a = _deskew_angle(gray2)
    gray3 = gray2.rotate(a, resample=Image.BICUBIC, fillcolor=255)
    rgb3 = rgb2.rotate(a, resample=Image.BICUBIC, fillcolor=(255, 255, 255))

    col_left, col_right = _detect_kod_col(gray3)
    col_width = col_right - col_left

    font_path = find_font()
    draw = ImageDraw.Draw(rgb3)
    for y_center, text in positions:
        if not text or y_center is None:
            continue
        size = 20
        font = ImageFont.truetype(font_path, size) if font_path else ImageFont.load_default()
        bb = draw.textbbox((0, 0), text, font=font)
        tw = bb[2] - bb[0]
        while tw > col_width and size > 13:
            size -= 1
            font = ImageFont.truetype(font_path, size) if font_path else ImageFont.load_default()
            bb = draw.textbbox((0, 0), text, font=font)
            tw = bb[2] - bb[0]
        th = bb[3] - bb[1]
        x = col_left - bb[0]
        y = int(round(y_center)) - th // 2 - bb[1]
        draw.text((x, y), text, font=font, fill=(0, 0, 0))

    rgb3.save(out_path)
    return out_path
