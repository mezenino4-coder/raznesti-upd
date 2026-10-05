# -*- coding: utf-8 -*-
"""
Мини-приложение: разнести накладную (УПД) по плану кооперации.

Использование:
    python app.py <план.xlsx> <скан1.jpg> [скан2.png ...] \
        [--password 2232] [--upd 527] [--date 08.09.2026] [--out результат.xlsx]

Без аргументов — интерактивный режим (спрашивает пути).

Что делает:
    1. OCR читает сканы накладной (rapidocr v5, кириллица) → позиции (код, название, кол-во)
    2. Расшифровывает план (пароль) → матчит по коду (с нечётким поиском)
    3. Снимает жёлтую заливку, вписывает №УПД и дату, излишек → комментарий «пришло N шт»
    4. Вписывает номера заказов в сканы (копии «*_с_заказами.png»)
    5. Сохраняет и заново шифрует план паролем
"""
import argparse
import io
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
from PIL import Image

from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR

import annotate
import core

DEFAULT_PASSWORD = "2232"

DET_MODEL = "PP-OCRv6_det_small.onnx"
REC_MODEL = "cyrillic_PP-OCRv5_rec_mobile.onnx"
CLS_MODEL = "ch_ppocr_mobile_v2.0_cls_mobile.onnx"


def models_dir() -> Path:
    """Папка с .onnx-моделями OCR. Работает и из исходников, и в PyInstaller-сборке.

    Приоритет:
      1) папка `models/` рядом с app.py (или внутри _MEIPASS у exe);
      2) `ocr_rus/models/` на уровень выше (для запуска из исходников).
    """
    candidates = []
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        candidates += [base / "models", base / "ocr_rus" / "models"]
    here = Path(__file__).resolve().parent
    candidates += [
        here / "models",
        here.parent / "ocr_rus" / "models",
    ]
    for cand in candidates:
        if (cand / DET_MODEL).exists() and (cand / REC_MODEL).exists():
            return cand
    return here / "models"


def build_engine():
    md = models_dir()
    return RapidOCR(params={
        "Global.model_root_dir": str(md),
        "Global.use_cls": False,
        "Det.model_path": str(md / DET_MODEL),
        "Rec.model_path": str(md / REC_MODEL),
        "Cls.model_path": str(md / CLS_MODEL),
        "Rec.lang_type": LangRec.CYRILLIC,
        "Rec.ocr_version": OCRVersion.PPOCRV5,
        "Rec.model_type": ModelType.MOBILE,
    })


def _score_text(txts):
    """Оценка 'осмысленности' текста: кириллические буквы + коды с точками.
    Повёрнутый на 180° текст OCR читает как мусор («962», «0008») — оценка ~0."""
    score = 0
    for t in (txts or []):
        t = (t or "").strip()
        if not t:
            continue
        cyr = sum(1 for ch in t if ('а' <= ch <= 'я') or ('А' <= ch <= 'Я') or ch in 'ёЁ')
        if cyr:
            score += 1 + cyr
        if t.count(".") >= 3:
            score += 5
    return score


def _resolve_orientation(engine, img, gray):
    """Определяет правильный угол поворота скана (0/90/180/270).

    Дисперсия горизонтальной проекции отличает «текст горизонтален» от
    «текст вертикален», но НЕ отличает верх от низа (90° и 270° дают одинаковую
    дисперсию). Поэтому среди двух кандидатов (best и best+180) выбираем тот,
    где OCR на уменьшенной копии читает больше кириллицы и кодов.
    """
    scores = {a: annotate._horiz_proj_var(gray.rotate(a, expand=True))
              for a in (0, 90, 180, 270)}
    best = max(scores, key=scores.get)
    candidates = [best, (best + 180) % 360]
    best_ang, best_score = best, -1
    for a in candidates:
        small = img.rotate(a, expand=True)
        small.thumbnail((1400, 1400), Image.LANCZOS)
        buf = io.BytesIO()
        small.save(buf, format="PNG")
        res = engine(buf.getvalue())
        sc = _score_text(res.txts)
        if sc > best_score:
            best_score, best_ang = sc, a
    return best_ang


def preprocess_scan(path, engine=None):
    """Автоповорот + выравнивание наклона (deskew) скана.
    Возвращает (PNG-байты, угол поворота)."""
    img = Image.open(path).convert("RGB")
    gray = img.convert("L")

    ang = _resolve_orientation(engine, img, gray)
    if ang:
        img = img.rotate(ang, expand=True)
        gray = gray.rotate(ang, expand=True)

    # небольшой наклон (deskew) — проекционный метод из annotate
    try:
        a = annotate._deskew_angle(gray)
        if abs(a) >= 0.1:
            img = img.rotate(a, resample=Image.BICUBIC,
                             fillcolor=(255, 255, 255))
    except Exception:
        pass

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue(), ang


def ocr_scan(path, engine=None):
    engine = engine or build_engine()
    data, ang = preprocess_scan(path, engine)
    res = engine(data)
    txts = res.txts if res.txts is not None else []
    boxes = res.boxes if res.boxes is not None else []
    return txts, boxes, ang


def _find_qty(txts, i, n):
    for j in range(i + 1, min(i + 6, n)):
        tj = (txts[j] or "").strip()
        m = re.fullmatch(r"(\d+)[.,]?", tj)
        if m and m.group(1) != "796":
            return int(m.group(1))
    nums = []
    for j in range(i + 1, min(i + 8, n)):
        tj = (txts[j] or "").strip().replace(" ", "").replace("\u00a0", "")
        m = re.fullmatch(r"(\d+)[.,](\d+)", tj)
        if m:
            nums.append(float(m.group(1) + "." + m.group(2)))
    for a in range(len(nums) - 1):
        if nums[a] > 0:
            q = nums[a + 1] / nums[a]
            if abs(q - round(q)) < 0.001:
                return int(round(q))
    return None


def _y_centers(boxes):
    yc = []
    if boxes is None:
        return yc
    for b in boxes:
        if b is None:
            yc.append(None)
            continue
        ys = [float(p[1]) for p in b]
        yc.append((min(ys) + max(ys)) / 2)
    return yc


def _name_nearby(txts, yc, i, y0):
    """Если код на своей строке, а название перенесено на соседнюю — ищем его ниже."""
    if y0 is None:
        return ""
    for j in range(i + 1, min(i + 14, len(txts))):
        tj = (txts[j] or "").strip()
        if not tj:
            continue
        yj = yc[j] if j < len(yc) else None
        if yj is not None and (yj - y0) > 45:
            break
        if tj.count(".") >= 3:
            continue
        if re.fullmatch(r"[\d\s.,]+", tj):
            continue
        if re.search(r"Без НДС|Без НДC|акциз|Итого|Всего|ШТ|Шт", tj):
            continue
        if re.search(r"[А-Яа-яЁё]", tj):
            return tj
    return ""


def parse_items(txts, boxes=None):
    items = []
    n = len(txts)
    yc = _y_centers(boxes)
    yc += [None] * (n - len(yc))
    for i, t in enumerate(txts):
        t = (t or "").strip()
        if not t:
            continue
        tokens = t.split()
        code = None
        k = -1
        for k, tok in enumerate(tokens):
            if tok.count(".") >= 3:
                code = tok.rstrip(".,")
                break
        if code is None:
            continue
        name = " ".join(tokens[k + 1:]).strip()
        if not name:
            name = _name_nearby(txts, yc, i, yc[i])
        qty = _find_qty(txts, i, n)
        items.append((code, name, qty, yc[i]))
    return items


def extract_upd(txts):
    full = " ".join(txts)
    # Предпочитаем именно «УПД №NNN», а не ссылки внутри документа
    # вроде «Счет на оплату №245» (это счёт, не УПД).
    m = re.search(r"УПД\s*№\s*(\d+)", full, re.IGNORECASE)
    if m:
        num = m.group(1)
    else:
        m = re.search(r"№\s*(\d+)", full)
        num = m.group(1) if m else None
    m = re.search(r"от\s*(\d{2}\.\d{2}\.\d{4})", full)
    date = m.group(1) if m else None
    return num, date


def main():
    if len(sys.argv) == 1:
        plan = input("Путь к плану (.xlsx): ").strip().strip('"')
        scans = input("Сканы накладной (через пробел): ").strip().strip('"').split()
        password = input("Пароль плана [2232]: ").strip() or "2232"
        a = argparse.Namespace(plan=plan, scans=scans, password=password,
                               upd=None, date=None, out="результат.xlsx")
    else:
        p = argparse.ArgumentParser(description=__doc__,
                                    formatter_class=argparse.RawDescriptionHelpFormatter)
        p.add_argument("plan")
        p.add_argument("scans", nargs="+")
        p.add_argument("--password", default=DEFAULT_PASSWORD)
        p.add_argument("--upd")
        p.add_argument("--date")
        p.add_argument("--out", default="результат.xlsx")
        a = p.parse_args()

    plan_path = Path(a.plan)
    today = datetime.now()

    print("Читаю сканы (OCR)...")
    scan_items = []  # [(scan_path, items, angle)]
    upd_no, upd_date = a.upd, a.date
    for sc in a.scans:
        sc = Path(sc)
        txts, boxes, ang = ocr_scan(sc)
        if upd_no is None or upd_date is None:
            n, d = extract_upd(txts)
            upd_no = upd_no or n
            upd_date = upd_date or d
        its = parse_items(txts, boxes)
        print(f"  {sc.name}: {len(its)} позиций (поворот {ang}°)")
        scan_items.append((sc, its, ang))

    if upd_no is None or upd_date is None:
        print("Не определил №УПД/дату — укажи --upd и --date")
        return

    try:
        upd_date_short = datetime.strptime(upd_date, "%d.%m.%Y").strftime("%d.%m.%y")
    except ValueError:
        upd_date_short = upd_date

    all_items = [(c, nm, q) for _, its, _ in scan_items for (c, nm, q, y) in its]
    print(f"\nУПД №{upd_no} от {upd_date}. Разношу...\n")

    report, orders_list = core.process(
        plan_path, a.password, all_items, upd_no, upd_date_short, today, Path(a.out)
    )
    print("\n".join(report))

    # вписывание номеров заказов в сканы
    print("\nВписываю номера заказов в сканы...")
    offset = 0
    for sc, its, ang in scan_items:
        n = len(its)
        orders = orders_list[offset:offset + n]
        offset += n
        positions = [(its[i][3], ", ".join(orders[i])) for i in range(n)]
        out_scan = sc.with_name(sc.stem + "_с_заказами.png")
        annotate.annotate(sc, positions, out_scan, angle=ang)
        print(f"  {sc.name} → {out_scan.name}")

    print(f"\nГотово → {a.out} (зашифрован паролем)")


if __name__ == "__main__":
    main()
