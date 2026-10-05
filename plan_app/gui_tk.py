# -*- coding: utf-8 -*-
"""
Графическое приложение: разнос УПД по плану кооперации (Tkinter).

То же окно, что и GTK-версия, но на Tkinter — встроен в Python и на Windows,
и на Ubuntu, без дополнительных пакетов (GTK на Windows не нужен).

Окно: выбрать план (.xlsx) → добавить сканы → ввести пароль → «Запустить».
Дальше OCR → матч по коду → снять заливку/вписать УПД → вписать номера заказов
в сканы → сохранить план (зашифрован тем же паролем).

Запуск:
    python  gui_tk.py      (с консолью, для отладки)
    pythonw gui_tk.py      (на Windows — без окна консоли, для ярлыка)
"""
import queue
import threading
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import core
import annotate
from app import build_engine, parse_items, extract_upd, ocr_scan

PASSWORD_DEFAULT = "2232"


class App:
    def __init__(self, root):
        self.root = root
        root.title("Разнос УПД по плану кооперации")
        root.geometry("720x560")
        root.minsize(640, 480)

        self.plan_path = None
        self.scan_paths = []
        self.engine = None
        self.busy = False
        self.log_q = queue.Queue()

        pad = {"padx": 10, "pady": 5}
        frm = tk.Frame(root)
        frm.pack(fill="both", expand=True, padx=12, pady=10)

        # --- План ---
        tk.Label(frm, text="План кооперации (.xlsx):", anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", **pad)
        self.plan_label = tk.Label(frm, text="не выбран", anchor="w", fg="#555")
        self.plan_label.grid(row=1, column=0, sticky="we", **pad)
        tk.Button(frm, text="Загрузить план", command=self.on_choose_plan).grid(
            row=1, column=1, sticky="e", **pad)

        # --- Сканы ---
        tk.Label(frm, text="Сканы накладной:", anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="w", **pad)
        self.scans_label = tk.Label(frm, text="не выбраны", anchor="w", fg="#555")
        self.scans_label.grid(row=3, column=0, sticky="we", **pad)
        box = tk.Frame(frm)
        box.grid(row=3, column=1, sticky="e")
        tk.Button(box, text="Добавить сканы", command=self.on_choose_scans).pack(
            side="left", padx=3)
        tk.Button(box, text="Очистить", command=self.on_clear_scans).pack(
            side="left", padx=3)

        # --- Пароль ---
        tk.Label(frm, text="Пароль плана:", anchor="w").grid(
            row=4, column=0, sticky="w", **pad)
        self.pass_entry = tk.Entry(frm, width=28)
        self.pass_entry.insert(0, PASSWORD_DEFAULT)
        self.pass_entry.grid(row=4, column=1, sticky="e", **pad)

        # --- Запуск ---
        self.run_btn = tk.Button(frm, text="Запустить", command=self.on_run,
                                 height=2, bg="#dcefff")
        self.run_btn.grid(row=5, column=0, columnspan=2, sticky="we", **pad)

        # --- Журнал ---
        tk.Label(frm, text="Журнал:", anchor="w").grid(
            row=6, column=0, columnspan=2, sticky="w", **pad)
        self.log_view = scrolledtext.ScrolledText(frm, height=18, state="disabled",
                                                  wrap="word", font=("Consolas", 9))
        self.log_view.grid(row=7, column=0, columnspan=2, sticky="nsew", **pad)

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(7, weight=1)

        # периодически выкачиваем сообщения из рабочего потока
        self._drain_log()

    # ---- журнал ----
    def log(self, msg):
        self.log_q.put(str(msg))

    def _drain_log(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.log_view.configure(state="normal")
                self.log_view.insert("end", msg + "\n")
                self.log_view.see("end")
                self.log_view.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(120, self._drain_log)

    def _set_busy(self, busy):
        self.busy = busy
        self.run_btn.configure(state="disabled" if busy else "normal")

    # ---- обработчики ----
    def on_choose_plan(self):
        p = filedialog.askopenfilename(
            title="Выбери план кооперации",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")])
        if p:
            self.plan_path = Path(p)
            self.plan_label.configure(text=str(self.plan_path), fg="#000")

    def on_choose_scans(self):
        files = filedialog.askopenfilenames(
            title="Выбери сканы накладной",
            filetypes=[("Изображения", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                       ("Все файлы", "*.*")])
        for f in files:
            p = Path(f)
            if p not in self.scan_paths:
                self.scan_paths.append(p)
        self._refresh_scans_label()

    def on_clear_scans(self):
        self.scan_paths = []
        self._refresh_scans_label()

    def _refresh_scans_label(self):
        if not self.scan_paths:
            self.scans_label.configure(text="не выбраны", fg="#555")
        else:
            self.scans_label.configure(
                text="; ".join(p.name for p in self.scan_paths), fg="#000")

    def on_run(self):
        if self.busy:
            return
        if not self.plan_path:
            messagebox.showwarning("Разнос УПД", "Сначала загрузи план кооперации.")
            return
        if not self.scan_paths:
            messagebox.showwarning("Разнос УПД", "Добавь хотя бы один скан накладной.")
            return
        # пароль читаем в главном потоке (Tkinter не потокобезопасен)
        password = self.pass_entry.get().strip() or PASSWORD_DEFAULT
        plan_path = self.plan_path
        scan_paths = list(self.scan_paths)
        self._set_busy(True)
        threading.Thread(target=self._work,
                         args=(plan_path, scan_paths, password), daemon=True).start()

    def _work(self, plan_path, scan_paths, password):
        try:
            self.log("Читаю сканы (OCR)...")
            if self.engine is None:
                self.engine = build_engine()
            all_items = []
            scan_items = []
            upd_no = upd_date = None
            for sc in scan_paths:
                txts, boxes, ang = ocr_scan(sc, self.engine)
                if upd_no is None or upd_date is None:
                    n, d = extract_upd(txts)
                    upd_no = upd_no or n
                    upd_date = upd_date or d
                its = parse_items(txts, boxes)
                self.log(f"  {sc.name}: {len(its)} позиций")
                scan_items.append((sc, its, ang))
                all_items += [(c, nm, q) for (c, nm, q, y) in its]

            if upd_no is None or upd_date is None:
                self.log("⚠️ Не удалось определить №УПД/дату из сканов.")
                return

            try:
                upd_short = datetime.strptime(upd_date, "%d.%m.%Y").strftime("%d.%m.%y")
            except ValueError:
                upd_short = upd_date

            self.log(f"УПД №{upd_no} от {upd_date}. Разношу...")

            out = plan_path.with_name(plan_path.stem + "_разнесён.xlsx")
            report, orders_list = core.process(
                plan_path, password,
                all_items, upd_no, upd_short, datetime.now(), out)

            for line in report:
                self.log(line)

            self.log("Вписываю номера заказов в сканы...")
            offset = 0
            for sc, its, ang in scan_items:
                n = len(its)
                orders = orders_list[offset:offset + n]
                offset += n
                positions = [(its[i][3], ", ".join(orders[i])) for i in range(n)]
                out_scan = sc.with_name(sc.stem + "_с_заказами.png")
                annotate.annotate(sc, positions, out_scan, angle=ang)
                self.log(f"  {sc.name} → {out_scan.name}")

            self.log(f"✅ Готово → {out} (зашифрован паролем)")
        except Exception as e:
            self.log(f"❌ Ошибка: {e}")
        finally:
            self.root.after(0, lambda: self._set_busy(False))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
