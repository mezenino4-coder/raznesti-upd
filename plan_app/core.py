# -*- coding: utf-8 -*-
"""
Ядро разноса УПД по плану кооперации: расшифровка → матч → запись → шифрование.

Логика (ТЗ Максима):
- жёлтый (FFFF00) = «ожидает поставки» — при приходе снимаем заливку
- «не заказаны» = только оранжевые строки (тема accent6 = F79646)
- разнос по «сроку исполнения» (раньше срок — первым)
- строки со сроком, истёкшим >=30 дней назад — не трогаем (менеджер забыл)
- пришло > суммарно жёлтого → излишек в оранжевые, комментарий «пришло N шт»
- код не найден точно → нечёткий поиск по имени+коду (thefuzz)

Функции:
    process(plan_path, password, items, upd_no, upd_date, today, out_path) -> report
"""
import io
import re
from datetime import datetime, timedelta
from pathlib import Path

import msoffcrypto, openpyxl
from openpyxl.styles import PatternFill

try:
    from thefuzz import fuzz
except ImportError:
    fuzz = None

CYR = {c: l for c, l in zip(
    "АВЕКМНОРСТУХавекмнорстух", "ABEKMHOPCTYXabekmnhopctyx")}

NO_FILL = PatternFill(fill_type=None)


def norm_code(s):
    s = (s or "").replace(" ", "").replace("\u00a0", "")
    return "".join(CYR.get(ch, ch) for ch in s).upper()


def short_order(o):
    """Короткий № заказа для вписывания в скан.

    Отбрасываем «пункт/часть» («Р2134 п.1» → «Р2134», «Р2130 ч.2» → «Р2130»,
    «Р2130 (ч.2» → «Р2130»). «Д №49», «Р2144» и т.п. остаются как есть.
    """
    if not o:
        return ""
    o = str(o).strip()
    o = re.split(r"\s*\(?\s*[пПчЧ]\.", o, maxsplit=1)[0]
    return o.strip().rstrip("( ").strip()


def _rgb(cell):
    f = cell.fill
    fc = f.fgColor if f else None
    t = getattr(fc, "type", None) if fc else None
    try:
        if t == "rgb":
            return str(fc.rgb)
        if t == "theme":
            return f"theme:{fc.theme}"
    except Exception:
        pass
    return None


def status_of(ws, r):
    c = _rgb(ws.cell(row=r, column=4))
    if c == "FFFFFF00":
        return "yellow"
    if c == "theme:9":
        return "orange"
    if c == "FF00B050":
        return "green"
    if c == "FFFF0000":
        return "red"
    return "none"


def load_plain(path, password):
    with open(path, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        of.load_key(password=password)
        buf = io.BytesIO()
        of.decrypt(buf)
        buf.seek(0)
    return openpyxl.load_workbook(buf)


def save_encrypted(wb, out_path, password):
    """Сохранить workbook и зашифровать паролем через msoffcrypto."""
    tmp = Path(out_path).with_suffix(".plain.xlsx")
    wb.save(tmp)
    from msoffcrypto.format.ooxml import OOXMLFile
    with open(tmp, "rb") as fin, open(out_path, "wb") as fout:
        of = OOXMLFile(fin)
        of.encrypt(password, fout)
    tmp.unlink()
    return out_path


def build_index(ws):
    idx = {}
    # № Заказа пишется один раз на все позиции одной служебки: в колонке A
    # он стоит только у первой строки блока, дальше пусто. Поэтому делаем
    # forward-fill — пустая ячейка наследует последний непустой номер заказа сверху.
    last_order = None
    for r in range(2, ws.max_row + 1):
        v4 = ws.cell(row=r, column=4).value
        if not isinstance(v4, str):
            continue
        raw_order = ws.cell(row=r, column=1).value
        if isinstance(raw_order, str) and raw_order.strip():
            last_order = raw_order.strip()
        code = norm_code(v4.strip().split(" ", 1)[0])
        name = v4.strip().split(" ", 1)[1] if " " in v4.strip() else ""
        idx.setdefault(code, []).append({
            "r": r, "code": code, "name": name,
            "qty": ws.cell(row=r, column=5).value,
            "deadline": ws.cell(row=r, column=2).value,
            "order": last_order,
            "status": status_of(ws, r),
            "upd": ws.cell(row=r, column=10).value,
        })
    return idx


def deadline_key(row):
    d = row["deadline"]
    return d if isinstance(d, datetime) else datetime.max


def fuzzy_find(ncode, name, idx):
    if fuzz is None:
        return None
    best, best_ratio = None, 0
    for c, rows in idx.items():
        rname = rows[0]["name"] if rows else ""
        if not (name and rname):
            continue
        if fuzz.partial_ratio(name, rname) < 70:
            continue
        r = fuzz.ratio(ncode, c)
        if r > best_ratio:
            best, best_ratio = c, r
    return (best, best_ratio) if best and best_ratio >= 60 else None


def process(plan_path, password, items, upd_no, upd_date, today, out_path,
            stale_days=30):
    """items: list of (code, name, qty). Возвращает (report, orders_list)."""
    wb = load_plain(plan_path, password)
    ws = wb["План"]
    idx = build_index(ws)
    stale_before = today - timedelta(days=stale_days)
    upd_text = f"упд №{upd_no} от {upd_date}"

    report = []
    orders_list = []
    excess_list = []
    for code, name, arrived in items:
        if not isinstance(arrived, int):
            report.append(f"[{code}] кол-во не определено — пропущено")
            orders_list.append([])
            continue
        ncode = norm_code(code)
        rows = idx.get(ncode, [])
        note = ""
        if not rows and fuzz is not None:
            hit = fuzzy_find(ncode, name, idx)
            if hit:
                rows = idx[hit[0]]
                note = f" [FUZZY {code}→{hit[0]}]"

        active = [r for r in rows if not (
            isinstance(r["deadline"], datetime) and r["deadline"] < stale_before)]
        yellows = sorted([r for r in active if r["status"] == "yellow"], key=deadline_key)
        oranges = sorted([r for r in active if r["status"] == "orange"], key=deadline_key)
        yellow_total = sum(r["qty"] or 0 for r in yellows)

        remaining = arrived
        closed = 0
        item_orders = []
        _seen = set()
        for r in yellows:
            if remaining <= 0:
                break
            need = r["qty"] or 0
            alloc = min(need, remaining)
            remaining -= alloc
            # снять жёлтую заливку по всей строке (колонки 4..10)
            for c in range(4, 11):
                cell = ws.cell(row=r["r"], column=c)
                if _rgb(cell) == "FFFFFF00":
                    cell.fill = NO_FILL
            ws.cell(row=r["r"], column=10).value = upd_text
            closed += 1
            o = short_order(r["order"])
            if o and o not in _seen:
                item_orders.append(o)
                _seen.add(o)
            report.append(f"  r{r['r']} жёлт→снять заливку + {upd_text} ({alloc}/{need})")
        for r in oranges:
            if remaining <= 0:
                break
            need = r["qty"] or 0
            alloc = min(need, remaining)
            remaining -= alloc
            ws.cell(row=r["r"], column=10).value = f"пришло {alloc} шт ({upd_text})"
            report.append(f"  r{r['r']} оранж→«пришло {alloc} шт» ({upd_text})")
        orders_list.append(item_orders)
        if arrived > yellow_total:
            disp_name = (rows[0]["name"] if rows else "") or name
            excess_list.append((code, disp_name, arrived, yellow_total,
                                arrived - yellow_total))
        status = f"[{code}]{note} пришло {arrived}: закрыто жёлтых {closed}, "
        if remaining > 0:
            status += f"ИЗЛИШЕК {remaining} некуда"
        else:
            status += "разнесено полностью"
        report.insert(0, status)

    if excess_list:
        report.append("")
        report.append("=== Излишек (пришло больше, чем ожидалось) ===")
        for code, disp_name, arrived, expected, excess in excess_list:
            report.append(
                f"[{code}] {disp_name}: пришло {arrived}, "
                f"ожидалось {expected}, излишек {excess} шт")

    save_encrypted(wb, out_path, password)
    return report, orders_list
