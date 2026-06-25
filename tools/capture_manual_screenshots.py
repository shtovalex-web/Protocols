# -*- coding: utf-8 -*-
"""
Снимки окон программы для подробной инструкции пользователя.

    py -3 tools/capture_manual_screenshots.py

Результат: bundle/manual_screenshots/*.png
"""

from __future__ import annotations

import os
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "bundle" / "manual_screenshots"


def _grab_bbox(bbox: tuple[int, int, int, int], filename: str) -> bool:
    from PIL import ImageGrab

    x1, y1, x2, y2 = bbox
    if x2 - x1 < 80 or y2 - y1 < 80:
        return False
    img = ImageGrab.grab(bbox=bbox)
    path = OUT / filename
    OUT.mkdir(parents=True, exist_ok=True)
    img.save(path)
    print(f"  OK {path.name} ({x2 - x1}x{y2 - y1})")
    return True


def _grab_win(win: tk.Misc, filename: str) -> bool:
    try:
        if not win.winfo_exists():
            print(f"  пропуск {filename}: окно закрыто")
            return False
    except tk.TclError:
        print(f"  пропуск {filename}: окно недоступно")
        return False

    win.update_idletasks()
    win.update()
    time.sleep(0.4)
    try:
        x = win.winfo_rootx()
        y = win.winfo_rooty()
        w = win.winfo_width()
        h = win.winfo_height()
    except tk.TclError:
        return False
    if w < 80 or h < 80:
        print(f"  пропуск {filename}: окно мало ({w}x{h})")
        return False
    return _grab_bbox((x, y, x + w, y + h), filename)


def _shot(win: tk.Misc, filename: str, label: str) -> bool:
    ok = _grab_win(win, filename)
    if not ok:
        print(f"  нет снимка: {label}")
    return ok


def _find_toplevel(app: tk.Misc, title_part: str) -> tk.Toplevel | None:
    for w in app.winfo_children():
        if isinstance(w, tk.Toplevel) and title_part.lower() in w.title().lower():
            return w
    return None


def _find_widget(parent: tk.Misc, cls: type) -> tk.Misc | None:
    for w in parent.winfo_children():
        if isinstance(w, cls):
            return w
        found = _find_widget(w, cls)
        if found is not None:
            return found
    return None


def _close_toplevels(app: tk.Misc) -> None:
    for w in list(app.winfo_children()):
        if isinstance(w, tk.Toplevel):
            try:
                w.destroy()
            except tk.TclError:
                pass
    app.update()


def _find_cascade_submenu(app: tk.Tk, label: str) -> tk.Menu | None:
    mbar = app.nametowidget(app["menu"])
    end = mbar.index("end")
    if end is None:
        return None
    for i in range(end + 1):
        if mbar.type(i) != "cascade":
            continue
        if mbar.entrycget(i, "label") == label:
            return app.nametowidget(mbar.entrycget(i, "menu"))
    return None


def _shot_menu(app: tk.Tk, menu_label: str, filename: str, x_offset: int) -> bool:
    submenu = _find_cascade_submenu(app, menu_label)
    if submenu is None:
        print(f"  нет меню: {menu_label}")
        return False
    app.lift()
    app.update()
    x = app.winfo_rootx() + x_offset
    y = app.winfo_rooty() + 55
    try:
        submenu.post(x, y)
        app.update()
        time.sleep(0.45)
        rx = app.winfo_rootx()
        ry = app.winfo_rooty()
        rw = min(max(app.winfo_width(), 420), 560)
        rh = 380 if menu_label == "Справка" else 320
        ok = _grab_bbox((rx, ry, rx + rw, ry + rh), filename)
    finally:
        try:
            submenu.unpost()
        except tk.TclError:
            pass
        app.update()
    return ok


def _shot_hotkeys_reference(app: tk.Tk) -> bool:
    win = tk.Toplevel(app)
    win.title("Горячие клавиши")
    win.transient(app)
    frm = ttk.Frame(win, padding=14)
    frm.pack(fill=tk.BOTH, expand=True)
    txt = (
        "F5 — обновить базы с диска (сотрудники и справочник программ).\n"
        "Кнопка «Обновить протокол из реестра Минтруда» — номера из выгрузки с портала.\n"
        "Ctrl+F — поле поиска по списку сотрудников.\n"
        "Ctrl+V — вставка из буфера в поля ввода; ПКМ — меню «Вставить».\n\n"
        "В списке сотрудников: Ctrl и Shift — выбор нескольких строк."
    )
    ttk.Label(frm, text=txt, justify=tk.LEFT, wraplength=420).pack(anchor=tk.W)
    ttk.Button(frm, text="OK", command=win.destroy).pack(anchor=tk.E, pady=(12, 0))
    win.update_idletasks()
    ok = _shot(win, "16_goryachie_klavishi.png", "горячие клавиши")
    win.destroy()
    return ok


def _shot_update_dialog_reference(app: tk.Tk) -> bool:
    win = tk.Toplevel(app)
    win.title("Обновление программы")
    win.transient(app)
    frm = ttk.Frame(win, padding=14)
    frm.pack(fill=tk.BOTH, expand=True)
    txt = (
        "Доступна новая версия 1.5.3.\n\n"
        "• Автообновление exe + data/\n\n"
        "Будут обновлены программа и файлы в папке data/ (шаблоны, справка).\n"
        "Файлы в корне папки (базы, protocols.db) не изменяются.\n\n"
        "Установить обновление сейчас?"
    )
    ttk.Label(frm, text=txt, justify=tk.LEFT, wraplength=420).pack(anchor=tk.W)
    bar = ttk.Frame(frm)
    bar.pack(anchor=tk.E, pady=(12, 0))
    ttk.Button(bar, text="Да", command=win.destroy).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(bar, text="Нет", command=win.destroy).pack(side=tk.LEFT)
    win.update_idletasks()
    ok = _shot(win, "20_obnovlenie_dialog.png", "диалог обновления")
    win.destroy()
    return ok


def main() -> int:
    sys.path.insert(0, str(ROOT / "ProtocolOHT_next"))
    sys.path.insert(0, str(ROOT))
    os.chdir(ROOT)

    from faq_viewer import open_faq_window
    from protocol_journal import PROTOCOL_JOURNAL_KIND_TECH
    from protocol_ui import ProtocolApp

    print("Запуск программы для снимков экрана…")
    app = ProtocolApp()
    app.update()
    ok = 0

    if _shot(app, "01_glavnoe_okno.png", "главное окно"):
        ok += 1

    if _shot_menu(app, "Администрирование", "12_menu_administraciya.png", 24):
        ok += 1
    if _shot_menu(app, "Минтруд", "13_menu_mintrud.png", 148):
        ok += 1
    if _shot_menu(app, "Справка", "14_menu_spravka.png", 228):
        ok += 1
    if _shot_update_dialog_reference(app):
        ok += 1

    app.var_emp_search.set("ов")
    app._refilter_employee_list()
    app.update()
    if _shot(app, "02_poisk_sotrudnikov.png", "поиск"):
        ok += 1

    app.var_emp_search.set("")
    app._refilter_employee_list()
    for key, var in app._prog_vars.items():
        var.set(key == "B")
    app.var_technical_protocol.set(False)
    sample = (
        "ПРОТОКОЛ проверки знаний\n\n"
        "Дата: 09.06.2026    № 5-06-2026\n"
        "ФИО: Иванов Иван Иванович\n"
        "Рег. № в реестре: (подставится из файла Минтруда)\n"
    )
    app._set_protocol_preview(None, sample, show_window=False)
    app.update()
    if _shot(app, "03_programmy_i_knopki.png", "программы и кнопки"):
        ok += 1
    if _shot(app, "15_knopka_reestr_mintrud.png", "кнопка реестра"):
        ok += 1

    app._open_admin_window()
    if app._admin_win is not None and _shot(app._admin_win, "04_nastroyki_i_dannye.png", "настройки"):
        ok += 1
    app._close_admin_window()

    app.show_template_variables_help()
    tpl = _find_toplevel(app, "Переменные")
    if tpl is not None and _shot(tpl, "17_peremennye_shablona.png", "переменные шаблона"):
        ok += 1
    _close_toplevels(app)

    app._open_commission_window()
    if app._commission_win is not None:
        nb = _find_widget(app._commission_win, ttk.Notebook)
        if nb is not None:
            nb.select(1)
            app._commission_win.update()
        if _shot(app._commission_win, "05_prikaz_i_komissiya.png", "комиссия ОТ"):
            ok += 1
        if nb is not None:
            nb.select(1)
            app._commission_win.update()
            if _shot(app._commission_win, "18_komissiya_tech.png", "комиссия тех."):
                ok += 1
    if app._commission_win is not None:
        app._commission_win.withdraw()

    app._show_preview_toplevel()
    if app._preview_win is not None and _shot(app._preview_win, "06_predprosmotr.png", "предпросмотр"):
        ok += 1
    if app._preview_win is not None:
        app._preview_win.destroy()
        app._preview_win = None

    app.show_protocol_journal()
    journal = _find_toplevel(app, "Журнал протоколов")
    if journal is not None and _shot(journal, "07_zhurnal.png", "журнал"):
        ok += 1
    _close_toplevels(app)

    app.show_protocol_journal(PROTOCOL_JOURNAL_KIND_TECH)
    journal_tech = _find_toplevel(app, "технических")
    if journal_tech is not None and _shot(journal_tech, "19_zhurnal_tech.png", "журнал тех."):
        ok += 1
    _close_toplevels(app)

    app._open_mintrud_employer_window()
    employer = _find_toplevel(app, "реквизиты")
    if employer is not None and _shot(employer, "08_mintrod_rekvizity.png", "реквизиты"):
        ok += 1
    _close_toplevels(app)

    app._open_mintrud_export_window()
    export = _find_toplevel(app, "Шаблон для сайта")
    if export is not None and _shot(export, "09_mintrod_vygruzka.png", "выгрузка"):
        ok += 1
    _close_toplevels(app)

    open_faq_window(app)
    faq = _find_toplevel(app, "Справка")
    if faq is not None and _shot(faq, "10_spravka_faq.png", "FAQ"):
        ok += 1
    _close_toplevels(app)

    if _shot_hotkeys_reference(app):
        ok += 1

    app._show_about_window()
    about = _find_toplevel(app, "О программе")
    if about is not None and _shot(about, "11_o_programme.png", "о программе"):
        ok += 1
    _close_toplevels(app)

    app.quit()
    app.destroy()
    print(f"\nГотово: {ok} снимков в {OUT}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
