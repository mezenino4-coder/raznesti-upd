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

def _base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


MODEL_DIR = _base_dir() / "ocr_rus" / "models"
DEFAULT_PASSWORD = "2232"


def build_engine():
    return RapidOCR(params={
        "Global.model_root_dir": str(MODEL_DIR),
        "Global.use_cls": False,
        "Rec.lang_type": LangRec.CYRILLIC,
        "Rec.ocr_version": OCRVersion.PPOCRV5,
        "Rec.model_type": ModelType.MOBILE,
    })


def preprocess_scan(path):
    """Автоповорот (90°) + выравнивание наклона (deskew) скана.
    Возвращает PNG-байты выровненного изображения для OCR."""
    img = Image.open(path).convert("RGB")
    gray = img.convert("L")

    # та же ориентация, что и при вписывании заказов (annotate) —
    # чтобы координаты OCR совпадали с координатами вписывания.
    gray2, best = annotate._auto_orient(gray)
    if best:
        img = img.rotate(best, expand=True)
        gray = gray2

    # небольшой наклон (deskew) — проекционный метод из annotate
    try:
        ang = annotate._deskew_angle(gray)
        if abs(ang) >= 0.1:
            img = img.rotate(ang, resample=Image.BICUBIC,
                             fillcolor=(255, 255, 255))
    except Exception:
        pass

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def ocr_scan(path, engine=None):
    engine = engine or build_engine()
    data = preprocess_scan(path)
    res = engine(data)
    txts = res.txts if res.txts is not None else []
    boxes = res.boxes if res.boxes is not None else []
    return txts, boxes


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


def parse_items(txts, boxes=None):
    items = []
    n = len(txts)
    for i, t in enumerate(txts):
        t = (t or "").strip()
        if not t:
            continue
        tokens = t.split()
        code = None
        for k, tok in enumerate(tokens):
            if tok.count(".") >= 3:
                code = tok
                break
        if code is None:
            continue
        name = " ".join(tokens[k + 1:])
        if not name.strip():
            continue
        qty = _find_qty(txts, i, n)
        y_center = None
        if boxes is not None and i < len(boxes) and boxes[i] is not None:
            ys = [float(p[1]) for p in boxes[i]]
            y_center = (min(ys) + max(ys)) / 2
        items.append((code, name, qty, y_center))
    return items


def extract_upd(txts):
    full = " ".join(txts)
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
    scan_items = []  # [(scan_path, items)]
    upd_no, upd_date = a.upd, a.date
    for sc in a.scans:
        sc = Path(sc)
        txts, boxes = ocr_scan(sc)
        if upd_no is None or upd_date is None:
            n, d = extract_upd(txts)
            upd_no = upd_no or n
            upd_date = upd_date or d
        its = parse_items(txts, boxes)
        print(f"  {sc.name}: {len(its)} позиций")
        scan_items.append((sc, its))

    if upd_no is None or upd_date is None:
        print("Не определил №УПД/дату — укажи --upd и --date")
        return

    try:
        upd_date_short = datetime.strptime(upd_date, "%d.%m.%Y").strftime("%d.%m.%y")
    except ValueError:
        upd_date_short = upd_date

    all_items = [(c, nm, q) for _, its in scan_items for (c, nm, q, y) in its]
    print(f"\nУПД №{upd_no} от {upd_date}. Разношу...\n")

    report, orders_list = core.process(
        plan_path, a.password, all_items, upd_no, upd_date_short, today, Path(a.out)
    )
    print("\n".join(report))

    # вписывание номеров заказов в сканы
    print("\nВписываю номера заказов в сканы...")
    offset = 0
    for sc, its in scan_items:
        n = len(its)
        orders = orders_list[offset:offset + n]
        offset += n
        positions = [(its[i][3], ", ".join(orders[i])) for i in range(n)]
        out_scan = sc.with_name(sc.stem + "_с_заказами.png")
        annotate.annotate(sc, positions, out_scan)
        print(f"  {sc.name} → {out_scan.name}")

    print(f"\nГотово → {a.out} (зашифрован паролем)")


if __name__ == "__main__":
    main()
