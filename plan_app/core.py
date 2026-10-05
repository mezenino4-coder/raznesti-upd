# -*- coding: utf-8 -*-
"""
Ядро разноса УПД по плану кооперации: расшифровка → матч → запись → шифрование.

Логика (ТЗ Максима):
- жёлтый (FFFF00) = «ожидает поставки» — при приходе снимаем заливку
- оранжевый («не заказано», theme:9 = F79646) — при приходе снимаем так же
- разнос по «сроку исполнения» (раньше срок — первым)
- строки со сроком, истёкшим >=30 дней назад — не трогаем (менеджер забыл)
- пришло > суммарно жёлтого+оранжевого → излишек
- код не найден точно → нечёткий поиск по имени+коду (thefuzz)

ВАЖНО про сохранение: план содержит диаграммы/картинки. openpyxl при
пересохранении их выбрасывает → Excel ругается «Ошибка в части содержимого».
Поэтому изменения пишутся напрямую в XML внутри xlsx (только нужные ячейки),
все остальные части файла (диаграммы, рисунки, sharedStrings) остаются как есть.

Функции:
    process(plan_path, password, items, upd_no, upd_date, today, out_path) -> report
"""
import copy
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

import msoffcrypto, openpyxl

try:
    from thefuzz import fuzz
except ImportError:
    fuzz = None

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

CYR = {c: l for c, l in zip(
    "АВЕКМНОРСТУХавекмнорстух", "ABEKMHOPCTYXabekmnhopctyx")}


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


# --------------------------------------------------------------------------
# Прямое редактирование xlsx (без потери диаграмм/картинок)
# --------------------------------------------------------------------------

def decrypt_to_bytes(path, password):
    with open(path, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        of.load_key(password=password)
        buf = io.BytesIO()
        of.decrypt(buf)
        buf.seek(0)
    return buf.getvalue()


def _col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _resolve_sheet_path(z, sheet_name):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wb)
    rid = m.group(1)
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m2 = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % re.escape(rid), rels)
    target = m2.group(1)
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def _build_twin_map(z):
    """Карта «цветной стиль → такой же, но без заливки».

    Обрабатываем жёлтые (fillId 2 = FFFFFF00) И оранжевые «не заказано»
    (fillId 5 = theme:9) — обе при приходе снимаем.

    Возвращает (twin, clones, xf_count):
      twin   — {index_цветного_xf: index_xf_без_заливки}
      clones — [(new_index, xml_строка)] новые xf, которые надо добавить в styles.xml
      xf_count — исходное число xf (для пересчёта count)
    """
    ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    styles = ET.fromstring(z.read("xl/styles.xml"))
    xe = styles.find(NS + "cellXfs")
    xfs = xe.findall(NS + "xf")

    def sig(xf):
        al = xf.find(NS + "alignment")
        return (xf.get("numFmtId"), xf.get("fontId"), xf.get("borderId"), xf.get("xfId"),
                xf.get("applyFont"), xf.get("applyBorder"), xf.get("applyNumberFormat"),
                xf.get("applyAlignment"), ET.tostring(al) if al is not None else "")

    colored = [i for i, x in enumerate(xfs) if x.get("fillId") in ("2", "5")]
    twin = {}
    clones = []
    next_idx = len(xfs)
    for i in colored:
        si = sig(xfs[i])
        found = None
        for j, x in enumerate(xfs):
            if j != i and x.get("fillId") == "0" and sig(x) == si:
                found = j
                break
        if found is not None:
            twin[i] = found
        else:
            c = copy.deepcopy(xfs[i])
            c.set("fillId", "0")
            xmlstr = re.sub(r"\bns0:", "", ET.tostring(c, encoding="unicode"))
            twin[i] = next_idx
            clones.append((next_idx, xmlstr))
            next_idx += 1
    return twin, clones, len(xfs)


def _apply_edits(sheet_xml, twin, cell_ops):
    """cell_ops: {(row, col): {"unfill": bool, "text": str|None}}"""
    def cell_re(ref):
        return re.compile(r'<c r="%s"(?:[^>]*/>|[^>]*>.*?</c>)' % ref, re.S)

    def find_s(cell_xml):
        m = re.search(r'\bs="(\d+)"', cell_xml)
        return int(m.group(1)) if m else None

    for (row, col), op in sorted(cell_ops.items()):
        ref = _col_letter(col) + str(row)
        m = cell_re(ref).search(sheet_xml)
        cur_s = None
        cell_txt = m.group(0) if m else None
        if cell_txt:
            cur_s = find_s(cell_txt)
        new_s = cur_s
        if op.get("unfill") and cur_s is not None and cur_s in twin:
            new_s = twin[cur_s]
        if op.get("text") is not None:
            txt = (op["text"].replace("&", "&amp;")
                               .replace("<", "&lt;")
                               .replace(">", "&gt;"))
            s_attr = (' s="%d"' % new_s) if new_s is not None else ""
            new_cell = ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                        % (ref, s_attr, txt))
            if m:
                sheet_xml = sheet_xml[:m.start()] + new_cell + sheet_xml[m.end():]
            else:
                row_re = re.compile(r'(<row r="%d"[^>]*>)' % row)
                rm = row_re.search(sheet_xml)
                if rm:
                    end = sheet_xml.index("</row>", rm.start())
                    sheet_xml = sheet_xml[:end] + new_cell + sheet_xml[end:]
                else:
                    sheet_xml = sheet_xml + new_cell
        elif op.get("unfill") and m and cur_s in twin:
            new_cell = re.sub(r'\bs="%d"' % cur_s, 's="%d"' % twin[cur_s], cell_txt, count=1)
            sheet_xml = sheet_xml[:m.start()] + new_cell + sheet_xml[m.end():]
    return sheet_xml


def save_edits_xml(plain_bytes, cell_ops, out_path, password):
    """Применяет правки к расшифрованному xlsx и шифрует результат паролем."""
    z = zipfile.ZipFile(io.BytesIO(plain_bytes))
    sheet_path = _resolve_sheet_path(z, "План")
    twin, clones, xf_count = _build_twin_map(z)

    sheet_xml = _apply_edits(z.read(sheet_path).decode("utf-8"), twin, cell_ops)

    styles_xml = None
    if clones:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        for _idx, xmlstr in clones:
            xmlstr = re.sub(r"\bns0:", "", xmlstr)
            styles_xml = styles_xml.replace("</cellXfs>", xmlstr + "</cellXfs>", 1)
        tag_m = re.search(r"<cellXfs [^>]*>", styles_xml)
        new_tag = re.sub(r'count="\d+"', 'count="%d"' % (xf_count + len(clones)), tag_m.group(0))
        styles_xml = styles_xml.replace(tag_m.group(0), new_tag, 1)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            data = z.read(item.filename)
            if item.filename == sheet_path:
                data = sheet_xml.encode("utf-8")
            elif item.filename == "xl/styles.xml" and styles_xml is not None:
                data = styles_xml.encode("utf-8")
            zo.writestr(item, data)

    tmp = Path(out_path).with_suffix(".plain.xlsx")
    tmp.write_bytes(out.getvalue())
    from msoffcrypto.format.ooxml import OOXMLFile
    with open(tmp, "rb") as fin, open(out_path, "wb") as fout:
        OOXMLFile(fin).encrypt(password, fout)
    tmp.unlink()
    return out_path


# --------------------------------------------------------------------------
# Основная логика разноса
# --------------------------------------------------------------------------

def process(plan_path, password, items, upd_no, upd_date, today, out_path,
            stale_days=30):
    """items: list of (code, name, qty). Возвращает (report, orders_list)."""
    plain = decrypt_to_bytes(plan_path, password)
    wb = openpyxl.load_workbook(io.BytesIO(plain))
    ws = wb["План"]
    idx = build_index(ws)
    stale_before = today - timedelta(days=stale_days)
    upd_text = f"упд №{upd_no} от {upd_date}"

    report = []
    orders_list = []
    excess_list = []
    cell_ops = {}  # {(row, col): {"unfill": bool, "text": str|None}}
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
        # Жёлтые («ожидает поставки») и оранжевые («не заказано») обрабатываем
        # одинаково: снять заливку + вписать №УПД + заказ в скан.
        targets = sorted(
            [r for r in active if r["status"] in ("yellow", "orange")],
            key=deadline_key)
        expected_total = sum(r["qty"] or 0 for r in targets)

        remaining = arrived
        closed = 0
        item_orders = []
        _seen = set()
        for r in targets:
            if remaining <= 0:
                break
            need = r["qty"] or 0
            alloc = min(need, remaining)
            remaining -= alloc
            # снять заливку по всей строке (колонки 4..10)
            for c in range(4, 11):
                op = cell_ops.setdefault((r["r"], c), {"unfill": False, "text": None})
                op["unfill"] = True
            op = cell_ops.setdefault((r["r"], 10), {"unfill": False, "text": None})
            op["text"] = upd_text
            closed += 1
            o = short_order(r["order"])
            if o and o not in _seen:
                item_orders.append(o)
                _seen.add(o)
            tag = "жёлт" if r["status"] == "yellow" else "оранж"
            report.append(f"  r{r['r']} {tag}→снять заливку + {upd_text} ({alloc}/{need})")
        orders_list.append(item_orders)
        if arrived > expected_total:
            disp_name = (rows[0]["name"] if rows else "") or name
            excess_list.append((code, disp_name, arrived, expected_total,
                                arrived - expected_total))
        status = f"[{code}]{note} пришло {arrived}: закрыто {closed}, "
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

    save_edits_xml(plain, cell_ops, out_path, password)
    return report, orders_list
