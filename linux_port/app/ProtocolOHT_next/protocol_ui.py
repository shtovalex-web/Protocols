# -*- coding: utf-8 -*-
"""Главное окно tkinter: меню, форма, предпросмотр, сохранение протоколов."""

from __future__ import annotations

import os
import sqlite3
import sys
import traceback
from datetime import date
from pathlib import Path
from typing import Any

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx_template_protection import (
    protect_standard_protocol_templates,
    save_formed_protocol_docx,
    unprotect_standard_protocol_templates,
)

from app_paths import application_bundle_dir, application_exe_dir
from clipboard_ui import install_clipboard_support, register_clipboard_window
from commission_admin import (
    COMMISSION_KIND_OT,
    COMMISSION_KIND_TECH,
    CommissionAdminPanel,
    CommissionState,
    load_mintrud_employer_from_db,
    load_mintrud_trained_registry_path,
    load_protect_bundle_templates_enabled,
    load_tech_protocol_template_docx_path,
    save_protect_bundle_templates_enabled,
    refresh_commission_pool_from_excel,
    save_mintrud_employer_to_db,
    save_mintrud_trained_registry_path,
    save_tech_protocol_template_docx_path,
)
from employees_io import (
    EMPLOYEES_EXCEL_FILENAME,
    EmployeeExcelError,
    EmployeeRecord,
    PROGRAMS_EXCEL_FILENAME,
    TechVProgramInfo,
    copy_program_sheets_from_workbook,
    format_fio_filename_surname_initials,
    listbox_label_for_employee,
    listbox_subdivision_header,
    subdivision_group_key,
    load_all_tech_v_programs_from_excel,
    load_employees_from_excel,
    sort_employees_by_subdivision_then_fio,
)
from excel_data_cache import (
    get_cached_b_program_title,
    get_cached_pp_table_title,
    get_cached_siz_table_title,
    invalidate_employees_cache_for_path,
    invalidate_program_catalog_cache_for_path,
    save_employees_cache,
    try_load_employees_from_cache,
)
from faq_viewer import open_changelog_window, open_faq_window
from programs_v_prof import (
    VProfProfessionCandidate,
    match_profession_in_v_prof,
    similar_professions_in_v_prof,
    v_prof_candidates_for_profession_list,
    v_prof_search_prefix_display,
)
from v_prof_combinations import (
    VProfCombinationConfig,
    VProfCombinationsDialog,
    needs_combinations_dialog,
    professions_by_fio,
    selection_signature,
)
from mintrud_export import write_mintrud_template_xlsx
from mintrud_trained_registry import load_trained_registry_index
from protocol_docx import (
    B_PROGRAM_SHEET_NAME,
    PROTOCOL_BODY_FONT_PT,
    PROTOCOL_PROGRAM_CHECKBOX_SHORT,
    PROTOCOL_PROGRAM_DEFS,
    PROTOCOL_PROGRAM_UI_LABELS,
    PROTOCOL_TEMPLATE_FILENAME,
    PROTOCOL_TEMPLATE_TECH_FILENAME,
    PROTOCOL_TEMPLATE_VARIABLES_DOC,
    ProtocolTemplateError,
    V_PROF_SHEET_NAME,
    _v_prof_select_best_row,
    _all_document_paragraphs_ordered,
    _find_form_template_bounds,
    _iter_paragraph_runs,
    _table_employees_dedupe_by_fio,
    build_filled_protocol_document,
    default_protocol_save_filename,
    document_to_plain_text,
    existing_per_employee_docx_in_folder,
    format_protocol_number_for_template,
    format_v_program_table_block_title,
    is_word_protocol_template,
    protocol_sequence_start_int,
    protocol_technical_template_path,
    protocol_template_path,
    raw_employee_rows_same_fio_as,
    resolve_protocol_template_path,
    save_protocol_docx_from_template,
    v_program_merged_parts_for_raw_employee,
    v_program_ordered_unique_parts_global,
)
from protocol_errors import append_error_journal
from protocol_journal import (
    PROTOCOL_JOURNAL_KIND_OT,
    PROTOCOL_JOURNAL_KIND_TECH,
    build_protocol_export_meta_json,
    clear_protocol_journal,
    default_journal_registry_export_path,
    export_meta_protocol_no,
    export_protocol_journal_registry,
    format_journal_list_line,
    get_all_protocols,
    get_protocols_journal_display,
    journal_ids_and_error_for_per_employee_batch,
    save_protocol,
)
from protocol_output import (
    build_protocol_text,
    write_protocol_docx,
    write_protocol_pdf,
    write_protocol_pdf_from_docx_template,
)
from protocol_paths import (
    DATABASE_FILENAME,
    database_path,
    employees_excel_default_path,
    load_last_protocol_no,
    mintrud_export_output_dir,
    programs_excel_default_path,
    protocols_output_dir,
    save_last_protocol_no,
)
from protocol_app_info import APP_FULL_NAME, APP_WINDOW_TITLE, populate_application_about_text
from protocol_embedded_assets import embedded_logo_png_bytes
from protocol_recovery import export_recovery_templates_to_folder
from ui_theme import (
    UI,
    Colors,
    FIELD_COMBO_STYLE,
    FIELD_DATE_STYLE,
    FIELD_STYLE,
    apply_color_scheme,
    apply_theme,
    apply_startup_geometry,
    apply_theme_to_window,
    color_scheme_choices,
    configure_listbox,
    configure_readonly_text,
    current_color_scheme_id,
    load_color_scheme_from_settings,
    pad,
    SCHEME_LABELS,
    SPACING,
    SPACING_LG,
    SPACING_SM,
)
from ui_widgets import WidgetTooltip, attach_tooltip


def render_document_to_text_widget(widget: tk.Text, doc: Document) -> None:
    """Выводит документ Word в tk.Text с базовым сохранением жирного/курсива по run."""
    widget.configure(state=tk.NORMAL)
    widget.delete("1.0", tk.END)
    for tag in (
        "pv_bold",
        "pv_italic",
        "pv_bi",
        "pv_body",
        "pv_body_bold",
        "pv_body_italic",
        "pv_body_bi",
    ):
        try:
            widget.tag_delete(tag)
        except tk.TclError:
            pass
    widget.tag_configure("pv_bold", font=(*UI.font_body, "bold"))
    widget.tag_configure("pv_italic", font=(*UI.font_body, "italic"))
    widget.tag_configure("pv_bi", font=(*UI.font_body, "bold italic"))
    preview_body = (UI.font_preview_body[0], PROTOCOL_BODY_FONT_PT)
    widget.tag_configure("pv_body", font=preview_body)
    widget.tag_configure("pv_body_bold", font=(*preview_body, "bold"))
    widget.tag_configure("pv_body_italic", font=(*preview_body, "italic"))
    widget.tag_configure("pv_body_bi", font=(*preview_body, "bold italic"))

    ordered = _all_document_paragraphs_ordered(doc)
    plines = [p.text for p in ordered]
    try:
        body_start, body_end = _find_form_template_bounds(plines)
    except ValueError:
        body_start, body_end = 0, len(plines)

    for pi, para in enumerate(ordered):
        in_body = body_start <= pi < body_end
        for run in _iter_paragraph_runs(para):
            chunk = run.text
            if not chunk:
                continue
            b = bool(run.bold)
            i = bool(run.italic)
            if in_body:
                if b and i:
                    tags: tuple[str, ...] = ("pv_body_bi",)
                elif b:
                    tags = ("pv_body_bold",)
                elif i:
                    tags = ("pv_body_italic",)
                else:
                    tags = ("pv_body",)
            else:
                if b and i:
                    tags = ("pv_bi",)
                elif b:
                    tags = ("pv_bold",)
                elif i:
                    tags = ("pv_italic",)
                else:
                    tags = ()
            widget.insert(tk.END, chunk, tags)
        widget.insert(tk.END, "\n")

    if doc.tables:
        widget.insert(tk.END, "\n── Таблица (текст ячеек) ──\n", ("pv_body",))
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.replace("\n", " ").strip() for c in row.cells]
                if not any(cells):
                    continue
                widget.insert(tk.END, "  |  ".join(cells) + "\n", ("pv_body",))


GRADE_OPTIONS = ("удовлетворительно", "неудовлетворительно")
CHECK_TYPE_OPTIONS = ("плановая", "внеплановая")
# Ширина полей формы (символы): длинные должности из V_PROF не обрезаются.
MAIN_FORM_ENTRY_CHARS = 72
MAIN_WINDOW_MIN_WIDTH = 900
MAIN_WINDOW_MIN_HEIGHT = 560
EMPLOYEE_LIST_HEIGHT_NORMAL = 16
EMPLOYEE_LIST_HEIGHT_TECH = 8
# Текст в окне журнала, если в записи не сохранён content (намеренно).
JOURNAL_PLACEHOLDER_NO_BODY = (
    "Полный текст протокола в журнал не сохраняется (меньше персональных данных в базе).\n\n"
    "Сформируйте протокол заново в главном окне или откройте сохранённый файл DOCX/PDF."
)

def _shrink_photoimage_to_max_height(img: tk.PhotoImage, max_h: int) -> tk.PhotoImage:
    cur = img
    while cur.height() > max_h:
        cur = cur.subsample(2, 2)
    return cur


def _default_user_data_hint() -> str:
    """Где лежат рабочие Excel/БД по умолчанию (для подписей в интерфейсе)."""
    if getattr(sys, "frozen", False):
        return "папка с программой (рядом с .exe)"
    return "папка с main.py"


class _WidgetTooltip(WidgetTooltip):
    """Совместимость: динамическое обновление текста подсказки V_PROF."""

    def set_text(self, text: str) -> None:
        self._text = (text or "").strip()


def _attach_tooltip(widget: tk.Misc, text: str, **kw: Any) -> None:
    attach_tooltip(widget, text, **kw)


class ProtocolApp(tk.Tk):
    def __init__(self, *, journal_duplicates_removed: int = 0) -> None:
        super().__init__()
        # Не показывать окно до финального geometry — иначе при .exe виден «прыжок» размера.
        self.withdraw()
        self.template_path: Path | None = None
        self.employees_excel_path: Path | None = None
        self.programs_excel_path: Path | None = None
        self._employee_records: list[EmployeeRecord] = []
        # None в списке — строка-заголовок подразделения (не выбирается).
        self._employee_list_slot_gi: list[int | None] = []
        self.var_emp_search = tk.StringVar(value="")
        self._prog_vars: dict[str, tk.BooleanVar] = {
            key: tk.BooleanVar(value=False) for key, _, _ in PROTOCOL_PROGRAM_DEFS
        }
        self._admin_win: tk.Toplevel | None = None
        self._commission_state = CommissionState()
        self._tech_commission_state = CommissionState()
        self._commission_win: tk.Toplevel | None = None
        self._commission_panel: CommissionAdminPanel | None = None
        self._tech_commission_panel: CommissionAdminPanel | None = None
        self.var_technical_protocol = tk.BooleanVar(value=False)
        self._tech_v_programs_list: list[TechVProgramInfo] = []
        self.technical_template_path: Path | None = None
        self._sync_technical_template_path_from_settings()
        self._employee_search_blobs: list[str] = []
        self._emp_collapsed_subdivisions: set[str] = set()
        self._list_index_to_sub_header: dict[int, str] = {}
        self._after_refilter_id: str | None = None
        self._last_export_persons: list[EmployeeRecord] | None = None
        self._v_prof_combo_config: VProfCombinationConfig | None = None
        self._v_prof_combo_selection_sig: str | None = None
        self._status_var = tk.StringVar(value="")
        self._preview_plain_text: str = ""
        self._last_preview_doc: Document | None = None
        self._preview_win: tk.Toplevel | None = None
        self._embedded_logo_header: tk.PhotoImage | None = None
        self._window_icon_ico_path: Path | None = None
        self._load_embedded_branding_images()
        self._apply_icon()
        load_color_scheme_from_settings()
        self._ui_scheme_var = tk.StringVar(value=current_color_scheme_id())
        apply_theme(self)
        self.title(APP_WINDOW_TITLE)
        # Предпросмотр — в отдельном окне; размер подбирается по содержимому.
        self.minsize(MAIN_WINDOW_MIN_WIDTH, MAIN_WINDOW_MIN_HEIGHT)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_menu()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_app_quit)
        self._setup_admin_window()
        self._refresh_technical_template_labels()
        self._ensure_bundle_templates_protected()
        self._try_autoload_employees()
        self._refresh_employees_file_label()
        self._refresh_programs_file_label()
        self._bind_keyboard_shortcuts()
        install_clipboard_support(self)
        self._sync_status_bar()
        self.update_idletasks()
        self.deiconify()
        self.after_idle(
            lambda: apply_startup_geometry(
                self,
                min_width=MAIN_WINDOW_MIN_WIDTH,
                min_height=MAIN_WINDOW_MIN_HEIGHT,
            )
        )
        if journal_duplicates_removed > 0:
            self.after(400, lambda: self._maybe_notify_journal_purge(journal_duplicates_removed))

    def _maybe_notify_journal_purge(self, removed: int) -> None:
        if removed <= 0:
            return
        from commission_admin import _app_setting_get, _app_setting_set
        from protocol_app_info import APP_VERSION

        key = "journal_purge_notice_version"
        ver = (APP_VERSION or "").strip()
        if _app_setting_get(key, "").strip() == ver:
            return
        _app_setting_set(key, ver)
        messagebox.showinfo(
            "Журнал протоколов",
            f"Из базы удалено устаревших дублей записей: {removed}.\n\n"
            "При повторном формировании протокола запись в журнале обновляется, "
            "а не добавляется повторно.",
            parent=self,
        )

    def _register_clipboard_for_window(self, win: tk.Misc) -> None:
        """Вставка Ctrl+V / ПКМ в полях дополнительного окна."""
        try:
            win.after_idle(lambda w=win: register_clipboard_window(w))
        except tk.TclError:
            pass

    def _face_sheet_profession(self) -> str:
        """Должность с лицевой части формы (в т.ч. подмена из матрицы V_PROF)."""
        if hasattr(self, "entry_position"):
            return self.entry_position.get().strip()
        return ""

    def _v_prof_combo_kwargs(self) -> dict[str, object]:
        """Параметры совмещений для сборки протокола."""
        cfg = self._v_prof_combo_config
        if cfg is None:
            return {
                "v_prof_enabled_norm_keys": None,
                "v_prof_enabled_by_fio": None,
                "v_prof_main_by_fio": None,
            }
        return {
            "v_prof_enabled_norm_keys": cfg.global_enabled_norm_keys(),
            "v_prof_enabled_by_fio": cfg.enabled_by_fio,
            "v_prof_main_by_fio": cfg.main_by_fio,
        }

    def _invalidate_v_prof_combination_choice(self) -> None:
        self._v_prof_combo_config = None
        self._v_prof_combo_selection_sig = None

    def _professions_for_v_prof_hint(self) -> list[str]:
        """Должности для подсказки V_PROF: по каждому выбранному сотруднику."""
        cfg = self._v_prof_combo_config
        if cfg and cfg.enabled_by_fio:
            out: list[str] = []
            seen: set[str] = set()
            for _fk, main in cfg.main_by_fio.items():
                t = (main or "").strip()
                if t and t.lower() not in seen:
                    seen.add(t.lower())
                    out.append(t)
            for rec in self._collect_table_persons():
                from v_program_registry_match import norm_profession_key

                fk = norm_profession_key(rec.fio or "")
                enabled = cfg.enabled_by_fio.get(fk)
                if enabled is None:
                    continue
                for pr in (rec.profession, rec.profession2):
                    t = (pr or "").strip()
                    if not t or norm_profession_key(t) not in enabled:
                        continue
                    k = t.lower()
                    if k in seen:
                        continue
                    seen.add(k)
                    out.append(t)
            return out
        out2: list[str] = []
        seen2: set[str] = set()
        fs = self._face_sheet_profession()
        if fs:
            out2.append(fs)
            seen2.add(fs.strip().lower())
        for rec in self._collect_table_persons():
            for pr in (rec.profession, rec.profession2):
                t = (pr or "").strip()
                if not t:
                    continue
                k = t.lower()
                if k in seen2:
                    continue
                seen2.add(k)
                out2.append(t)
        return out2

    def _configure_v_prof_combinations(self, persons_raw: list[EmployeeRecord]) -> bool:
        """
        Диалог совмещений при нескольких должностях у одного ФИО.
        False — пользователь отменил формирование.
        """
        if not self._prog_vars["V"].get():
            return True
        if not persons_raw:
            return True
        sig = selection_signature(persons_raw)
        if self._v_prof_combo_selection_sig == sig and self._v_prof_combo_config is not None:
            return True
        if not needs_combinations_dialog(persons_raw):
            self._v_prof_combo_selection_sig = sig
            self._v_prof_combo_config = None
            return True
        groups = professions_by_fio(persons_raw)
        init_cfg = self._v_prof_combo_config
        dlg = VProfCombinationsDialog(
            self,
            groups,
            initial_main_by_fio=init_cfg.main_by_fio if init_cfg else None,
            initial_enabled_by_fio=init_cfg.enabled_by_fio if init_cfg else None,
        )
        result = dlg.run()
        if result is None:
            return False
        self._v_prof_combo_config = result
        self._v_prof_combo_selection_sig = sig
        if len(groups) == 1:
            only_key = groups[0][0]
            main_one = result.main_by_fio.get(only_key, "")
            if main_one:
                self.entry_position.delete(0, tk.END)
                self.entry_position.insert(0, main_one)
                self._refresh_v_prof_profession_hint(main_one)
        else:
            self._refresh_v_prof_profession_hint()
        return True

    def _bind_keyboard_shortcuts(self) -> None:
        """F5 — обновить базы; Ctrl+F — фокус в поле поиска сотрудников."""

        def _on_f5(_event: object | None = None) -> str | None:
            self._refresh_data_bases_clicked()
            return "break"

        def _on_ctrl_f(_event: object | None = None) -> str | None:
            if hasattr(self, "entry_emp_search"):
                self.entry_emp_search.focus_set()
                self.entry_emp_search.select_range(0, tk.END)
            return "break"

        self.bind("<F5>", _on_f5)
        self.bind("<Control-f>", _on_ctrl_f)
        self.bind("<Control-F>", _on_ctrl_f)

    def _sync_status_bar(self) -> None:
        if not hasattr(self, "_status_var"):
            return
        total = len(self._employee_records)
        shown = sum(1 for g in self._employee_list_slot_gi if g is not None)
        q = self.var_emp_search.get().strip()
        if total == 0:
            self._status_var.set(
                "Сотрудники не загружены — проверьте файл Excel и нажмите «Обновить базы с диска» (F5)."
            )
            return
        if q:
            self._status_var.set(f"По поиску показано {shown} из {total} сотрудников.")
        else:
            self._status_var.set(f"Загружено сотрудников: {total}.")

    def _fit_window_geometry(
        self,
        win: tk.Toplevel | tk.Tk | None = None,
        *,
        margin_x: int = 32,
        margin_y: int = 64,
        min_w: int = 400,
        min_h: int = 300,
    ) -> None:
        """Подогнать размер окна под содержимое (без полос прокрутки)."""
        wdg = win if win is not None else self
        wdg.update_idletasks()
        req_w = max(min_w, wdg.winfo_reqwidth() + 12)
        req_h = max(min_h, wdg.winfo_reqheight() + 12)
        root = wdg.winfo_toplevel()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        width = min(req_w, sw - margin_x)
        height = min(req_h, sh - margin_y)
        wdg.minsize(min(req_w, width), min(req_h, height))
        if win is not None and win is not self:
            self.update_idletasks()
            x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
            y = self.winfo_rooty() + 32
            if y + height > sh - 8:
                y = max(0, (sh - height) // 2)
            if x + width > sw - 8:
                x = max(0, sw - width - 8)
        else:
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2)
        wdg.geometry(f"{width}x{height}+{x}+{y}")

    def _fit_main_window_to_form(self) -> None:
        """Подогнать высоту главного окна под форму (расширить или сжать)."""
        if not self.var_technical_protocol.get():
            self.minsize(MAIN_WINDOW_MIN_WIDTH, MAIN_WINDOW_MIN_HEIGHT)
            self._apply_main_window_geometry()
            return
        self.update_idletasks()
        sh = self.winfo_screenheight()
        req_h = max(MAIN_WINDOW_MIN_HEIGHT, self.winfo_reqheight() + 24)
        cap_h = max(MAIN_WINDOW_MIN_HEIGHT, sh - 48)
        target_h = min(req_h, cap_h)
        cur_w = max(self.winfo_width(), MAIN_WINDOW_MIN_WIDTH)
        cur_x = self.winfo_x()
        cur_y = self.winfo_y()
        if abs(self.winfo_height() - target_h) > 6:
            self.geometry(f"{cur_w}x{target_h}+{cur_x}+{cur_y}")
        self.minsize(MAIN_WINDOW_MIN_WIDTH, min(target_h, cap_h))

    def _themed_toplevel(self, parent: tk.Misc | None = None) -> tk.Toplevel:
        """Дочернее окно с той же темой, что и главное (clam, поля, скругление)."""
        win = tk.Toplevel(parent if parent is not None else self)
        apply_theme_to_window(win)
        self._apply_embedded_window_icon(win)
        return win

    def _visible_toplevel(self, win: tk.Misc | None) -> bool:
        if win is None:
            return False
        try:
            return bool(win.winfo_exists()) and win.winfo_viewable()
        except tk.TclError:
            return False

    def _dialog_parent(self) -> tk.Misc:
        if self._visible_toplevel(self._admin_win):
            return self._admin_win  # type: ignore[return-value]
        if self._visible_toplevel(self._commission_win):
            return self._commission_win  # type: ignore[return-value]
        return self

    def _make_modal(self, win: tk.Toplevel, *, parent: tk.Misc | None = None) -> None:
        par = parent if parent is not None else self
        win.transient(par)
        try:
            win.grab_set()
        except tk.TclError:
            pass
        win.lift()
        try:
            win.focus_force()
        except tk.TclError:
            pass

    def _release_modal(self, win: tk.Misc | None) -> None:
        if win is None:
            return
        try:
            if win.winfo_exists():
                win.grab_release()
        except tk.TclError:
            pass

    def _restore_parent_modal(self, parent: tk.Misc | None) -> None:
        if parent is None or parent is self or not isinstance(parent, tk.Toplevel):
            return
        if not self._visible_toplevel(parent):
            return
        try:
            parent.grab_set()
            parent.lift()
            parent.focus_force()
        except tk.TclError:
            pass

    def _close_modal_window(self, win: tk.Toplevel, *, parent: tk.Misc | None = None) -> None:
        self._release_modal(win)
        try:
            win.destroy()
        except tk.TclError:
            pass
        self._restore_parent_modal(parent if parent is not None else self)

    def _ensure_main_window_width_for_text(self, *texts: str) -> None:
        """Расширить главное окно, если длинный текст (должность V_PROF) не помещается."""
        samples = [t for t in texts if (t or "").strip()]
        if not samples:
            return
        try:
            fnt = tkfont.Font(font=self.entry_position.cget("font"))
        except tk.TclError:
            fnt = tkfont.Font(family=UI.family, size=10)
        text_px = max(fnt.measure(t) for t in samples)
        want_w = min(
            max(MAIN_WINDOW_MIN_WIDTH, text_px + 300),
            self.winfo_screenwidth() - 32,
        )
        self.update_idletasks()
        if self.winfo_width() < want_w - 12:
            h = max(self.winfo_height(), self.minsize()[1])
            x = max(0, min(self.winfo_x(), self.winfo_screenwidth() - want_w - 8))
            y = self.winfo_y()
            self.geometry(f"{want_w}x{h}+{x}+{y}")
        self.minsize(min(want_w, self.winfo_screenwidth() - 32), self.minsize()[1])

    def _apply_main_window_geometry(self) -> None:
        self._fit_window_geometry(
            self,
            margin_x=40,
            margin_y=72,
            min_w=MAIN_WINDOW_MIN_WIDTH,
            min_h=520,
        )

    def _persist_protocol_no_field(self) -> None:
        save_last_protocol_no(self.entry_protocol_no.get().strip())

    def _default_protocol_initialfile(self, protocol_no: str, date_str: str, ext: str) -> str:
        suf = ""
        if self._last_export_persons and len(self._last_export_persons) == 1:
            suf = format_fio_filename_surname_initials(self._last_export_persons[0].fio)
        return default_protocol_save_filename(protocol_no, date_str, ext, person_suffix=suf)

    def _on_app_quit(self) -> None:
        self._persist_protocol_no_field()
        self.destroy()

    def report_callback_exception(self, exc, val, tb) -> None:
        try:
            tb_text = "".join(traceback.format_exception(exc, val, tb))
            append_error_journal(
                "Ошибка в обработчике интерфейса (tkinter)",
                str(val),
                traceback_text=tb_text,
            )
        except Exception:
            pass
        super().report_callback_exception(exc, val, tb)

    def _employees_file_resolved(self) -> Path:
        return self.employees_excel_path or employees_excel_default_path()

    def _programs_file_resolved(self) -> Path:
        return self.programs_excel_path or programs_excel_default_path()

    def _load_embedded_branding_images(self) -> None:
        """Логотип в шапке из protocol_embedded_assets (вшитый PNG). Значок окон — bundle/icon.ico."""
        self._embedded_logo_header = None
        try:
            raw = tk.PhotoImage(master=self, data=embedded_logo_png_bytes())
            self._embedded_logo_header = _shrink_photoimage_to_max_height(raw, 72)
        except tk.TclError:
            pass

    def _resolve_window_icon_ico_path(self) -> Path | None:
        """Значок окон: bundle/icon.ico (каталог комплекта), иначе icon.ico в рабочей папке."""
        bundle_ico = application_bundle_dir() / "icon.ico"
        if bundle_ico.is_file():
            return bundle_ico.resolve()
        cwd_ico = Path("icon.ico")
        if cwd_ico.is_file():
            return cwd_ico.resolve()
        return None

    def _apply_embedded_window_icon(self, win: tk.Misc) -> None:
        if self._window_icon_ico_path is None:
            return
        try:
            win.iconbitmap(str(self._window_icon_ico_path))
        except (tk.TclError, OSError):
            pass

    def _apply_icon(self) -> None:
        """Значок заголовка: bundle/icon.ico (или icon.ico в cwd)."""
        self._window_icon_ico_path = self._resolve_window_icon_ico_path()
        if self._window_icon_ico_path is None:
            return
        try:
            self.iconbitmap(str(self._window_icon_ico_path))
        except (tk.TclError, OSError):
            self._window_icon_ico_path = None

    def _attach_program_logo(self, db_bar: ttk.Frame) -> None:
        """Вшитый логотип слева в верхней панели."""
        if self._embedded_logo_header is None:
            return
        ttk.Label(db_bar, image=self._embedded_logo_header).pack(side=tk.LEFT, padx=(0, 10))

    def _build_menu(self) -> None:
        mbar = tk.Menu(self)
        self.config(menu=mbar)
        adm = tk.Menu(mbar, tearoff=0)
        mbar.add_cascade(label="Администрирование", menu=adm)
        adm.add_command(label="Настройки и данные…", command=self._open_admin_window)
        adm.add_command(label="Приказ и комиссия…", command=self._open_commission_window)
        adm.add_command(
            label="Выгрузка шаблонов восстановления…",
            command=self._export_recovery_templates_bundle,
        )

        mintrud = tk.Menu(mbar, tearoff=0)
        mbar.add_cascade(label="Минтруд", menu=mintrud)
        mintrud.add_command(
            label="Реквизиты работодателя…",
            command=self._open_mintrud_employer_window,
        )
        mintrud.add_command(
            label="Файл реестра обученных (.xlsx)…",
            command=lambda: self._pick_mintrud_trained_registry(self),
        )
        mintrud.add_command(
            label="Пересформировать протокол (номера реестра)…",
            command=self.regenerate_protocol_with_mintrud_registry,
        )
        mintrud.add_separator()
        mintrud.add_command(
            label="Шаблон для загрузки на сайт…",
            command=self._open_mintrud_export_window,
        )

        view_m = tk.Menu(mbar, tearoff=0)
        mbar.add_cascade(label="Вид", menu=view_m)
        scheme_m = tk.Menu(view_m, tearoff=0)
        view_m.add_cascade(label="Цветовая схема", menu=scheme_m)
        for sid, label in color_scheme_choices():
            scheme_m.add_radiobutton(
                label=label,
                variable=self._ui_scheme_var,
                value=sid,
                command=self._apply_ui_color_scheme,
            )

        help_m = tk.Menu(mbar, tearoff=0)
        mbar.add_cascade(label="Справка", menu=help_m)
        help_m.add_command(
            label="Справка и FAQ…",
            command=lambda: open_faq_window(self),
        )
        help_m.add_command(
            label="Горячие клавиши…",
            command=self._show_hotkeys_help,
        )
        help_m.add_command(
            label="Журнал доработок…",
            command=lambda: open_changelog_window(self),
        )
        help_m.add_separator()
        help_m.add_command(
            label="О программе…",
            command=self._show_about_window,
        )

    def _show_hotkeys_help(self) -> None:
        messagebox.showinfo(
            "Горячие клавиши",
            "F5 — обновить базы с диска (сотрудники и справочник программ).\n"
            "Кнопка «Обновить протокол из реестра Минтруда» — номера из выгрузки с портала.\n"
            "Ctrl+F — поле поиска по списку сотрудников.\n"
            "Ctrl+V — вставка из буфера в поля ввода; ПКМ — меню «Вставить».\n\n"
            "В списке сотрудников: Ctrl и Shift — выбор нескольких строк.",
            parent=self,
        )

    def _show_about_window(self) -> None:
        win = self._themed_toplevel()
        win.title("О программе")
        win.transient(self)
        win.resizable(True, True)
        win.minsize(380, 280)
        outer = ttk.Frame(win, padding=16)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(1, weight=1)
        col_txt = 0
        if self._embedded_logo_header is not None:
            ttk.Label(outer, image=self._embedded_logo_header).grid(
                row=0, column=0, rowspan=2, padx=(0, 12), sticky=tk.NW
            )
            col_txt = 1
        else:
            outer.columnconfigure(0, weight=1)
        ttk.Label(outer, text=APP_FULL_NAME, style="Title.TLabel", font=UI.font_dialog_title).grid(
            row=0, column=col_txt, sticky=tk.W, pady=(0, 8)
        )
        txt = tk.Text(outer, width=48, height=12, wrap=tk.WORD, font=UI.font_body)
        configure_readonly_text(txt)
        txt.grid(row=1, column=col_txt, sticky=tk.NSEW)
        populate_application_about_text(txt)
        txt.configure(state=tk.DISABLED)
        bf = ttk.Frame(outer)
        span = 2 if self._embedded_logo_header is not None else 1
        bf.grid(row=2, column=0, columnspan=span, pady=(12, 0), sticky=tk.E)
        ttk.Button(bf, text="Закрыть", command=lambda: self._close_modal_window(win)).pack(side=tk.RIGHT)
        self._make_modal(win)
        win.update_idletasks()
        try:
            win.focus_set()
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        g = pad()
        g_sm = pad(small=True)

        main = ttk.Frame(self, padding=SPACING_LG)
        main.grid(row=0, column=0, sticky=tk.NSEW)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=0)
        main.rowconfigure(1, weight=1)
        main.rowconfigure(2, weight=0)

        db_bar = ttk.Frame(main, style="Toolbar.TFrame", padding=(SPACING, SPACING_SM))
        db_bar.grid(row=0, column=0, sticky=tk.EW, **g)
        self._attach_program_logo(db_bar)
        self.btn_refresh_data_bases = ttk.Button(
            db_bar,
            text="Обновить базы с диска",
            command=self._refresh_data_bases_clicked,
            style="Small.TButton",
        )
        self.btn_refresh_data_bases.pack(side=tk.LEFT, padx=(0, 10))
        _attach_tooltip(
            self.btn_refresh_data_bases,
            "Сотрудники и справочник программ с диска. F5.",
        )

        lf = ttk.Labelframe(main, text="Формирование протокола", style="Card.TLabelframe")
        lf.grid(row=1, column=0, sticky=tk.NSEW, **g)
        lf.columnconfigure(0, weight=2, minsize=280)
        lf.columnconfigure(1, weight=3, minsize=360)
        lf.rowconfigure(1, weight=1)
        lf.rowconfigure(2, weight=0)

        filter_fr = ttk.Frame(lf)
        filter_fr.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, SPACING_SM))
        filter_fr.columnconfigure(1, weight=1)
        filter_fr.columnconfigure(3, weight=1)
        ttk.Label(filter_fr, text="Поиск:").grid(row=0, column=0, sticky=tk.W, padx=(0, SPACING_SM))
        self.entry_emp_search = ttk.Entry(
            filter_fr,
            textvariable=self.var_emp_search,
            style="Field.TEntry",
        )
        self.entry_emp_search.grid(row=0, column=1, sticky=tk.EW, padx=(0, SPACING_LG))
        _attach_tooltip(
            self.entry_emp_search,
            "Фильтр по ФИО, должности, подразделению, СНИЛС. Ctrl+F — быстрый переход в это поле.",
        )
        ttk.Label(filter_fr, text="Подразделение:").grid(row=0, column=2, sticky=tk.W, padx=(0, SPACING_SM))
        self.entry_subdivision = ttk.Entry(filter_fr, style="Field.TEntry")
        self.entry_subdivision.grid(row=0, column=3, sticky=tk.EW)
        _attach_tooltip(
            self.entry_subdivision,
            "Подразделение для протокола. При выборе сотрудника из списка подставляется из Excel.",
        )

        emp_lf = ttk.Labelframe(lf, text="Сотрудники", padding=SPACING_SM)
        emp_lf.grid(row=1, column=0, sticky=tk.NSEW, padx=(0, SPACING))
        emp_lf.columnconfigure(0, weight=1)
        emp_lf.rowconfigure(0, weight=1)
        emp_box = ttk.Frame(emp_lf)
        emp_box.grid(row=0, column=0, sticky=tk.NSEW)
        emp_box.columnconfigure(0, weight=1)
        emp_box.rowconfigure(0, weight=1)
        sb_emp = ttk.Scrollbar(emp_box)
        sb_emp.grid(row=0, column=1, sticky=tk.NS)
        self.list_employees = tk.Listbox(
            emp_box,
            height=EMPLOYEE_LIST_HEIGHT_NORMAL,
            selectmode=tk.EXTENDED,
            exportselection=False,
            yscrollcommand=sb_emp.set,
        )
        configure_listbox(self.list_employees)
        self.list_employees.grid(row=0, column=0, sticky=tk.NSEW, **g_sm)
        sb_emp.configure(command=self.list_employees.yview)
        self.list_employees.bind("<<ListboxSelect>>", self._on_employee_list_select)
        self.list_employees.bind("<Button-1>", self._on_employee_list_click, add="+")
        self.var_emp_search.trace_add(
            "write", lambda *_: self._schedule_refilter_employee_list()
        )
        _attach_tooltip(
            self.list_employees,
            "Сотрудники по подразделениям: щёлкните строку «▾ название (N)» — свернуть/развернуть группу. "
            "Ctrl и Shift — выбор нескольких сотрудников.",
        )

        right_col = ttk.Frame(lf)
        right_col.grid(row=1, column=1, sticky=tk.NSEW)
        right_col.columnconfigure(0, weight=1)

        person_lf = ttk.Labelframe(right_col, text="Участник", padding=SPACING_SM)
        person_lf.grid(row=0, column=0, sticky=tk.EW, pady=(0, SPACING_SM))
        person_lf.columnconfigure(1, weight=1)

        ttk.Label(person_lf, text="ФИО:").grid(row=0, column=0, sticky=tk.W, **g_sm)
        self.entry_fio = ttk.Entry(person_lf, style="Field.TEntry")
        self.entry_fio.grid(row=0, column=1, sticky=tk.EW, **g_sm)
        _attach_tooltip(
            self.entry_fio,
            "Если человека нет в списке выше — введите ФИО полностью (фамилия, имя, отчество).",
        )

        ttk.Label(person_lf, text="Должность:").grid(row=1, column=0, sticky=tk.NW, **g_sm)
        pos_col = ttk.Frame(person_lf)
        pos_col.grid(row=1, column=1, sticky=tk.EW, **g_sm)
        pos_col.columnconfigure(0, weight=1)
        self.entry_position = ttk.Entry(pos_col, style="Field.TEntry")
        self.entry_position.grid(row=0, column=0, columnspan=2, sticky=tk.EW)
        self.lbl_v_prof_match = ttk.Label(pos_col, text="", style="Hint.TLabel")
        self.lbl_v_prof_match.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))
        self._v_prof_match_tooltip = _WidgetTooltip(self.lbl_v_prof_match, "")
        self._v_prof_suggest_fr = ttk.Frame(pos_col)
        self._v_prof_suggest_fr.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(2, 0))
        self._v_prof_suggest_fr.columnconfigure(1, weight=1)
        self._v_prof_suggest_professions: dict[str, str] = {}
        ttk.Label(self._v_prof_suggest_fr, text="V_PROF:", style="Hint.TLabel").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 4)
        )
        self.cmb_v_prof_suggest = ttk.Combobox(
            self._v_prof_suggest_fr,
            state="readonly",
            font=UI.font_body,
        )
        self.cmb_v_prof_suggest.grid(row=0, column=1, sticky=tk.EW)
        self.btn_v_prof_apply = ttk.Button(
            self._v_prof_suggest_fr,
            text="Подставить",
            command=self._apply_v_prof_profession_from_combo,
            width=11,
            style="Small.TButton",
        )
        self.btn_v_prof_apply.grid(row=0, column=2, sticky=tk.W, padx=(4, 0))
        _attach_tooltip(
            self.cmb_v_prof_suggest,
            "Похожие профессии из V_PROF (по 1–2 первым словам должности). «Подставить» — в поле.",
        )
        self._v_prof_suggest_fr.grid_remove()
        self.entry_position.bind("<FocusOut>", self._on_position_focus_out)
        _attach_tooltip(
            self.entry_position,
            "Основная должность для программ «В» (для одного выбранного — здесь; "
            "для нескольких — в диалоге у каждого сотрудника своя). "
            "Совмещения из списка учитываются для шапки и таблицы «В». "
            f"Сопоставление с листом {V_PROF_SHEET_NAME} — под полем.",
        )

        meta_lf = ttk.Labelframe(right_col, text="Реквизиты протокола", padding=SPACING_SM)
        meta_lf.grid(row=1, column=0, sticky=tk.EW)
        meta_lf.columnconfigure(1, weight=1)
        meta_lf.columnconfigure(3, weight=1)

        ttk.Label(meta_lf, text="Дата:").grid(row=0, column=0, sticky=tk.W, **g_sm)
        self.entry_date = ttk.Entry(meta_lf, width=14, style="FieldDate.TEntry")
        self.entry_date.grid(row=0, column=1, sticky=tk.W, **g_sm)
        self.entry_date.insert(0, date.today().strftime("%d.%m.%Y"))
        _attach_tooltip(
            self.entry_date,
            "Дата протокола. Формат ДД.ММ.ГГГГ или ДД.ММ.ГГ. Участвует в номере в бланке Word (номер-месяц-год).",
        )

        ttk.Label(meta_lf, text="№ протокола:").grid(row=0, column=2, sticky=tk.W, **g_sm)
        self.entry_protocol_no = ttk.Entry(meta_lf, width=18, style="Field.TEntry")
        self.entry_protocol_no.grid(row=0, column=3, sticky=tk.EW, **g_sm)
        _saved_protocol_no = load_last_protocol_no()
        if _saved_protocol_no:
            self.entry_protocol_no.insert(0, _saved_protocol_no)
        _attach_tooltip(
            self.entry_protocol_no,
            "В бланк Word подставляется строка вида «номер-месяц-год» по этому полю и полю «Дата» "
            "(см. также «Переменные шаблона» в настройках).",
        )

        ttk.Label(meta_lf, text="Оценка:").grid(row=1, column=0, sticky=tk.W, **g_sm)
        self.combo_grade = ttk.Combobox(
            meta_lf,
            values=GRADE_OPTIONS,
            state="readonly",
            width=28,
            style=FIELD_COMBO_STYLE,
        )
        self.combo_grade.grid(row=1, column=1, sticky=tk.EW, **g_sm)
        self.combo_grade.current(0)

        ttk.Label(meta_lf, text="Проверка знаний:").grid(row=1, column=2, sticky=tk.W, **g_sm)
        self.combo_check_type = ttk.Combobox(
            meta_lf,
            values=CHECK_TYPE_OPTIONS,
            state="readonly",
            width=28,
            style=FIELD_COMBO_STYLE,
        )
        self.combo_check_type.grid(row=1, column=3, sticky=tk.EW, **g_sm)
        self.combo_check_type.current(0)

        prog_lf = ttk.Labelframe(right_col, text="Программы обучения", padding=SPACING_SM)
        prog_lf.grid(row=2, column=0, sticky=tk.EW, pady=(SPACING_SM, 0))
        for col in range(4):
            prog_lf.columnconfigure(col, weight=1)
        self._program_checkbuttons: list[ttk.Checkbutton] = []
        for pi, (key, _sheet, _fb) in enumerate(PROTOCOL_PROGRAM_DEFS):
            cb = ttk.Checkbutton(
                prog_lf,
                text=PROTOCOL_PROGRAM_CHECKBOX_SHORT.get(
                    key, PROTOCOL_PROGRAM_UI_LABELS.get(key, key)
                ),
                variable=self._prog_vars[key],
            )
            cb.grid(row=0, column=pi, sticky=tk.W, padx=(1, 4), pady=0)
            self._program_checkbuttons.append(cb)
            tip = PROTOCOL_PROGRAM_UI_LABELS.get(key, "").strip()
            if tip:
                _attach_tooltip(cb, tip)
        self.cb_technical_protocol = ttk.Checkbutton(
            prog_lf,
            text="Технич. протокол (лист Tech_V)",
            variable=self.var_technical_protocol,
            command=self._on_technical_protocol_toggle,
        )
        _tech_cb_row = 1
        self.cb_technical_protocol.grid(
            row=_tech_cb_row,
            column=0,
            columnspan=4,
            sticky=tk.W,
            padx=1,
            pady=(2, 0),
        )
        _attach_tooltip(
            self.cb_technical_protocol,
            "Отдельный шаблон Word (например default_protocol_tehnicheskiy.docx): номер в строке с "
            "«техническ…»; выберите программу в списке ниже (все строки листа Tech_V); комиссия — "
            "вкладка «Технич. вопросы» в окне приказа. Реестр Минтруда не используется.",
        )
        self.lbl_tech_v_program = ttk.Label(prog_lf, text="Программа по листу Tech_V:")
        self.cb_tech_v_program = ttk.Combobox(
            prog_lf,
            state="readonly",
            width=36,
        )
        self._tech_v_pick_lbl_grid = dict(
            row=_tech_cb_row + 1,
            column=0,
            sticky=tk.W,
            padx=4,
            pady=(0, 4),
        )
        self._tech_v_pick_cb_grid = dict(
            row=_tech_cb_row + 1,
            column=1,
            columnspan=3,
            sticky=tk.EW,
            padx=4,
            pady=(0, 4),
        )
        self.lbl_tech_v_program.grid(**self._tech_v_pick_lbl_grid)
        self.cb_tech_v_program.grid(**self._tech_v_pick_cb_grid)
        self.lbl_tech_v_program.grid_remove()
        self.cb_tech_v_program.grid_remove()
        _attach_tooltip(
            self.cb_tech_v_program,
            "Список строк листа Tech_V с непустым наименованием программы. "
            "Подставляется в шаблон (маркеры {{ТЕХ_ПРОГРАММА}} и текст после «…по программе:»). F5 — обновить с диска.",
        )
        _tech_tpl_row = _tech_cb_row + 2
        self._tech_tpl_btns = ttk.Frame(prog_lf)
        self._tech_tpl_btns_grid = dict(
            row=_tech_tpl_row,
            column=0,
            columnspan=4,
            sticky=tk.EW,
            padx=4,
            pady=(0, 2),
        )
        ttk.Button(
            self._tech_tpl_btns,
            text="Шаблон Word (тех.)…",
            command=self.pick_technical_template,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            self._tech_tpl_btns,
            text="Сброс тех. шаблона",
            command=self.clear_technical_template,
        ).pack(side=tk.LEFT, padx=(0, 8))
        self.lbl_technical_template_main = ttk.Label(
            prog_lf,
            text="",
            style="Hint.TLabel",
            wraplength=340,
        )
        self._tech_tpl_lbl_main_grid = dict(
            row=_tech_tpl_row + 1,
            column=0,
            columnspan=4,
            sticky=tk.W,
            padx=4,
            pady=(0, 4),
        )
        self._tech_tpl_btns.grid(**self._tech_tpl_btns_grid)
        self.lbl_technical_template_main.grid(**self._tech_tpl_lbl_main_grid)
        self._tech_tpl_btns.grid_remove()
        self.lbl_technical_template_main.grid_remove()

        bottom_lf = ttk.Labelframe(lf, text="Сформировать и сохранить", padding=SPACING_SM)
        bottom_lf.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(SPACING, 0))
        bottom_lf.columnconfigure(0, weight=2)
        bottom_lf.columnconfigure(1, weight=2)
        bottom_lf.columnconfigure(2, weight=1)
        bottom_lf.columnconfigure(3, weight=1)
        bottom_lf.columnconfigure(4, weight=1)

        self.btn_generate = ttk.Button(
            bottom_lf,
            text="Сформировать протокол",
            command=self.generate_protocol,
            style="Accent.TButton",
        )
        self.btn_generate.grid(row=0, column=0, sticky=tk.EW, padx=(0, 6), pady=(0, 4))
        ttk.Button(
            bottom_lf,
            text="По одному на каждого → в папку…",
            command=self.generate_protocol_per_employee_to_folder,
        ).grid(row=0, column=1, sticky=tk.EW, padx=(0, 6), pady=(0, 4))
        self.btn_refresh_mintrud_registry = ttk.Button(
            bottom_lf,
            text="Обновить из реестра Минтруд",
            command=self.regenerate_protocol_with_mintrud_registry,
            style="Small.TButton",
        )
        self.btn_refresh_mintrud_registry.grid(row=0, column=2, sticky=tk.EW, pady=(0, 4))
        _attach_tooltip(
            self.btn_refresh_mintrud_registry,
            "Пересобрать протокол с регистрационными номерами из файла реестра Минтруда "
            "(«Минтруд» → «Файл реестра обученных»). Те же ФИО, программы и дата, что на экране.",
        )

        self.btn_save = ttk.Button(
            bottom_lf,
            text="Сохранить DOCX",
            command=self.save_to_docx,
            state="disabled",
        )
        self.btn_save.grid(row=1, column=0, sticky=tk.EW, padx=(0, 6))

        self.btn_save_pdf = ttk.Button(
            bottom_lf,
            text="Сохранить PDF",
            command=self.save_to_pdf,
            state="disabled",
        )
        self.btn_save_pdf.grid(row=1, column=1, sticky=tk.EW, padx=(0, 6))

        self.btn_preview = ttk.Button(
            bottom_lf,
            text="Предпросмотр…",
            command=self._open_preview_window_manual,
            state="disabled",
        )
        self.btn_preview.grid(row=1, column=2, sticky=tk.EW)
        _attach_tooltip(
            self.btn_preview,
            "Открыть окно с текстом последнего сформированного протокола (то же, что после «Сформировать»).",
        )

        status_fr = ttk.Frame(main)
        status_fr.grid(row=2, column=0, sticky=tk.EW, **g)
        status_fr.columnconfigure(0, weight=1)
        ttk.Separator(status_fr, orient=tk.HORIZONTAL).grid(row=0, column=0, sticky=tk.EW, pady=(0, 6))
        ttk.Label(
            status_fr,
            textvariable=self._status_var,
            style="Status.TLabel",
            anchor=tk.W,
            wraplength=820,
        ).grid(row=1, column=0, sticky=tk.EW)

    def _open_preview_window_manual(self) -> None:
        if not (self._preview_plain_text or "").strip() and self._last_preview_doc is None:
            messagebox.showinfo(
                "Предпросмотр",
                "Сначала сформируйте протокол кнопкой «Сформировать протокол» "
                "или загрузите текст из журнала.",
                parent=self,
            )
            return
        self._show_preview_toplevel()

    def _show_preview_toplevel(self) -> None:
        if not (self._preview_plain_text or "").strip() and self._last_preview_doc is None:
            return
        if self._preview_win is not None:
            try:
                if self._preview_win.winfo_exists():
                    self._preview_win.destroy()
            except tk.TclError:
                pass
            self._preview_win = None

        win = self._themed_toplevel()
        self._preview_win = win
        win.title("Предпросмотр протокола")
        win.transient(self)
        win.minsize(560, 420)
        win.geometry(f"700x520+{self.winfo_rootx() + 40}+{self.winfo_rooty() + 40}")

        body = ttk.Frame(win, padding=8)
        body.pack(fill=tk.BOTH, expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        txt = tk.Text(body, wrap=tk.WORD, font=UI.font_body)
        configure_readonly_text(txt)
        sb = ttk.Scrollbar(body, command=txt.yview)
        txt.grid(row=0, column=0, sticky=tk.NSEW)
        sb.grid(row=0, column=1, sticky=tk.NS)
        txt.configure(yscrollcommand=sb.set)

        if self._last_preview_doc is not None:
            render_document_to_text_widget(txt, self._last_preview_doc)
        else:
            txt.insert("1.0", self._preview_plain_text)

        bf = ttk.Frame(win, padding=(8, 0, 8, 8))
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="Сохранить в DOCX", command=self.save_to_docx).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(bf, text="Сохранить в PDF", command=self.save_to_pdf).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        def _on_close() -> None:
            try:
                self._close_modal_window(win)
            finally:
                self._preview_win = None

        ttk.Button(bf, text="Закрыть", command=_on_close).pack(side=tk.RIGHT)

        win.protocol("WM_DELETE_WINDOW", _on_close)
        self._make_modal(win, parent=self._dialog_parent())
        self._register_clipboard_for_window(win)

    def _set_protocol_preview(
        self,
        doc: Document | None,
        plain: str,
        *,
        show_window: bool = True,
    ) -> None:
        self._preview_plain_text = (plain or "").rstrip()
        self._last_preview_doc = doc
        has = bool(self._preview_plain_text) or (doc is not None)
        if has:
            self.btn_save.state(["!disabled"])
            self.btn_save_pdf.state(["!disabled"])
            self.btn_preview.state(["!disabled"])
        if show_window and (
            (self._preview_plain_text or "").strip() or self._last_preview_doc is not None
        ):
            self.after(80, self._show_preview_toplevel)

    def _setup_admin_window(self) -> None:
        win = self._themed_toplevel()
        win.title("Администрирование — настройки и данные")
        win.withdraw()
        win.transient(self)
        self._admin_win = win
        win.protocol("WM_DELETE_WINDOW", self._close_admin_window)
        self._build_admin_window_content(win)
        self._register_clipboard_for_window(win)

    def _open_admin_window(self) -> None:
        if self._admin_win is None:
            return
        self._refresh_mintrud_registry_label()
        self._refresh_employees_file_label()
        self._refresh_programs_file_label()
        self._refresh_technical_template_labels()
        self._admin_win.deiconify()
        self._fit_window_geometry(self._admin_win, margin_x=32, margin_y=56, min_w=520, min_h=420)
        self._make_modal(self._admin_win)

    def _close_admin_window(self) -> None:
        if self._admin_win is not None:
            self._release_modal(self._admin_win)
            self._admin_win.withdraw()

    def _apply_ui_color_scheme(self) -> None:
        sid = self._ui_scheme_var.get()
        apply_color_scheme(sid, self)
        label = SCHEME_LABELS.get(sid, sid)
        self._status_var.set(f"Цветовая схема: {label} (сохранена)")
        if hasattr(self, "cmb_ui_scheme"):
            ids = getattr(self, "_ui_scheme_ids", [])
            if sid in ids:
                self.cmb_ui_scheme.current(ids.index(sid))
        for win in (self._admin_win, self._commission_win):
            if win is not None and self._visible_toplevel(win):
                apply_theme_to_window(win)

    def _apply_ui_color_scheme_from_admin(self) -> None:
        if not hasattr(self, "cmb_ui_scheme"):
            return
        idx = self.cmb_ui_scheme.current()
        ids = getattr(self, "_ui_scheme_ids", [])
        if idx < 0 or idx >= len(ids):
            return
        self._ui_scheme_var.set(ids[idx])
        self._apply_ui_color_scheme()

    def _open_commission_window(self) -> None:
        if self._commission_win is not None and self._commission_win.winfo_exists():
            self._commission_win.deiconify()
            self._make_modal(self._commission_win)
            self._register_clipboard_for_window(self._commission_win)
            if self._commission_panel is not None:
                refresh_commission_pool_from_excel(
                    self._commission_state,
                    self._employees_file_resolved(),
                    show_errors=False,
                    parent=self._commission_win,
                )
                self._tech_commission_state.pool = self._commission_state.pool
                self._commission_panel.refresh_pool_display()
                self._commission_panel.load_from_db_into_ui()
                if self._tech_commission_panel is not None:
                    self._tech_commission_panel.refresh_pool_display()
                    self._tech_commission_panel.load_from_db_into_ui()
            return

        win = self._themed_toplevel()
        win.title("Приказ и комиссия по проверке знаний")
        win.minsize(560, 640)
        win.transient(self)
        self._commission_win = win

        def _on_close_commission() -> None:
            # Не обнулять ссылки на панели: окно только скрывается (withdraw), при повторном открытии
            # нужны те же CommissionAdminPanel для обновления списка komission и загрузки из БД.
            self._release_modal(win)
            win.withdraw()

        win.protocol("WM_DELETE_WINDOW", _on_close_commission)

        outer = ttk.Frame(win, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        refresh_commission_pool_from_excel(
            self._commission_state,
            self._employees_file_resolved(),
            show_errors=False,
            parent=win,
        )
        self._tech_commission_state.pool = self._commission_state.pool

        nb = ttk.Notebook(outer)
        nb.pack(fill=tk.BOTH, expand=True)
        f_ot = ttk.Frame(nb, padding=4)
        f_tech = ttk.Frame(nb, padding=4)
        nb.add(f_ot, text="Охрана труда")
        nb.add(f_tech, text="Технич. вопросы")

        panel_ot = CommissionAdminPanel(
            f_ot,
            state=self._commission_state,
            get_excel_path=self._employees_file_resolved,
            dialog_parent=win,
            commission_kind=COMMISSION_KIND_OT,
            mirror_pool_state=self._tech_commission_state,
        )
        panel_ot.pack(fill=tk.BOTH, expand=True)
        self._commission_panel = panel_ot

        panel_tech = CommissionAdminPanel(
            f_tech,
            state=self._tech_commission_state,
            get_excel_path=self._employees_file_resolved,
            dialog_parent=win,
            commission_kind=COMMISSION_KIND_TECH,
            mirror_pool_state=self._commission_state,
        )
        panel_tech.pack(fill=tk.BOTH, expand=True)
        self._tech_commission_panel = panel_tech

        bf = ttk.Frame(outer)
        bf.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(bf, text="Закрыть", command=_on_close_commission).pack(side=tk.LEFT)

        win.deiconify()
        self._make_modal(win)
        self._register_clipboard_for_window(win)

    def _export_recovery_templates_bundle(self) -> None:
        d = filedialog.askdirectory(
            title="Папка для шаблонов восстановления (protocols.db, Data_base.xlsx, …)",
            parent=self,
        )
        if not d:
            return
        try:
            done, missing = export_recovery_templates_to_folder(Path(d))
        except EmployeeExcelError as e:
            messagebox.showerror("Выгрузка", str(e), parent=self)
            return
        except OSError as e:
            messagebox.showerror("Выгрузка", f"Не удалось записать файлы:\n{e}", parent=self)
            return
        except Exception as e:
            messagebox.showerror("Выгрузка", f"Ошибка:\n{e}", parent=self)
            return
        msg = f"Папка:\n{d}\n\nСоздано и скопировано файлов: {len(done)}."
        if missing:
            msg += (
                "\n\nВ комплекте программы не найдены (при необходимости добавьте вручную):\n"
                + "\n".join(missing)
            )
        messagebox.showinfo("Шаблоны восстановления", msg, parent=self)

    def _build_admin_window_content(self, win: tk.Toplevel) -> None:
        g = {"padx": 5, "pady": 5}
        outer = ttk.Frame(win, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        nb = ttk.Notebook(outer)
        nb.grid(row=0, column=0, sticky=tk.NSEW)

        tab_data = ttk.Frame(nb, padding=4)
        tab_tpl = ttk.Frame(nb, padding=4)
        tab_more = ttk.Frame(nb, padding=4)
        nb.add(tab_data, text="Файлы данных")
        nb.add(tab_tpl, text="Шаблоны")
        nb.add(tab_more, text="Прочее")
        tab_data.columnconfigure(0, weight=1)
        tab_tpl.columnconfigure(0, weight=1)
        tab_more.columnconfigure(0, weight=1)

        lf_ex = ttk.Labelframe(tab_data, text="Файл данных (сотрудники)", padding=6)
        lf_ex.grid(row=0, column=0, sticky=tk.EW, **g)
        eb = ttk.Frame(lf_ex)
        eb.grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Button(eb, text="Загрузить из Excel", command=self._reload_employees_from_admin).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 6)
        )
        ttk.Button(eb, text="Файл сотрудников…", command=self.pick_employees_excel).grid(
            row=0, column=1, sticky=tk.W, padx=(0, 8)
        )
        self.lbl_employees_file = ttk.Label(lf_ex, text="", wraplength=520)
        self.lbl_employees_file.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))

        lf_prog = ttk.Labelframe(tab_data, text="Справочник программ обучения", padding=6)
        lf_prog.grid(row=1, column=0, sticky=tk.EW, **g)
        pb = ttk.Frame(lf_prog)
        pb.grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Button(pb, text="Файл программ…", command=self.pick_programs_excel).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 6)
        )
        ttk.Button(
            pb,
            text="Вынести листы из Data_base…",
            command=self._export_programs_workbook_from_combined,
        ).grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        self.lbl_programs_file = ttk.Label(lf_prog, text="", wraplength=520)
        self.lbl_programs_file.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))

        lf_tpl = ttk.Labelframe(tab_tpl, text="Шаблон протокола", padding=6)
        lf_tpl.grid(row=0, column=0, sticky=tk.EW, **g)
        tr = ttk.Frame(lf_tpl)
        tr.grid(row=0, column=0, sticky=tk.W)
        ttk.Button(tr, text="Выбрать шаблон", command=self.pick_template).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 6)
        )
        ttk.Button(
            tr,
            text="Переменные шаблона",
            command=self.show_template_variables_help,
        ).grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        self.var_protect_bundle_templates = tk.BooleanVar(
            value=load_protect_bundle_templates_enabled()
        )
        ttk.Checkbutton(
            tr,
            text="Защита стандартных шаблонов в Word",
            variable=self.var_protect_bundle_templates,
            command=self._on_protect_bundle_templates_toggle,
        ).grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        tr_prot = ttk.Frame(lf_tpl)
        tr_prot.grid(row=2, column=0, sticky=tk.W, pady=(6, 0))
        ttk.Button(
            tr_prot,
            text="Включить защиту шаблонов…",
            command=self._protect_bundle_templates_clicked,
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        ttk.Button(
            tr_prot,
            text="Снять защиту (для правки)…",
            command=self._unprotect_bundle_templates_clicked,
        ).grid(row=0, column=1, sticky=tk.W, padx=(0, 6))
        ttk.Button(
            tr_prot,
            text="Папка шаблона…",
            command=self._open_template_folder_clicked,
        ).grid(row=0, column=2, sticky=tk.W)
        self.lbl_template = ttk.Label(lf_tpl, text=self._template_status_text(), wraplength=500)
        self.lbl_template.grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        tr_tech = ttk.Frame(lf_tpl)
        tr_tech.grid(row=3, column=0, sticky=tk.W, pady=(6, 0))
        ttk.Button(
            tr_tech,
            text="Шаблон Word (тех.)…",
            command=self.pick_technical_template,
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        ttk.Button(
            tr_tech,
            text="Сброс тех. шаблона",
            command=self.clear_technical_template,
        ).grid(row=0, column=1, sticky=tk.W)
        self.lbl_technical_template_admin = ttk.Label(
            lf_tpl,
            text="",
            wraplength=500,
            style="Hint.TLabel",
        )
        self.lbl_technical_template_admin.grid(row=4, column=0, sticky=tk.W, pady=(4, 0))

        lf_reg = ttk.Labelframe(
            tab_more,
            text="Реестр обученных (выгрузка с сайта Минтруда)",
            padding=6,
        )
        lf_reg.grid(row=0, column=0, sticky=tk.EW, **g)
        reg_bt = ttk.Frame(lf_reg)
        reg_bt.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        ttk.Button(
            reg_bt,
            text="Выбрать файл…",
            command=lambda: self._pick_mintrud_trained_registry(win),
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Button(
            reg_bt,
            text="Сбросить",
            command=self._clear_mintrud_trained_registry,
        ).grid(row=0, column=1, sticky=tk.W)
        self.lbl_mintrud_registry = ttk.Label(lf_reg, text="", wraplength=480)
        self.lbl_mintrud_registry.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))
        self._refresh_mintrud_registry_label()

        lf_more = ttk.Labelframe(tab_more, text="Дополнительно при заполнении", padding=6)
        lf_more.grid(row=1, column=0, sticky=tk.EW, **g)
        lf_more.columnconfigure(1, weight=1)
        ttk.Label(
            lf_more,
            text="Регистрационный номер (вручную, если нет в файле реестра):",
        ).grid(row=0, column=0, sticky=tk.NW, **g)
        self.entry_registry_no = ttk.Entry(lf_more, width=55, style=FIELD_STYLE)
        self.entry_registry_no.grid(row=0, column=1, sticky=tk.EW, **g)
        ttk.Label(lf_more, text="Доп. тема / строка для .txt:").grid(
            row=1, column=0, sticky=tk.W, **g
        )
        self.entry_theme = ttk.Entry(lf_more, width=55, style=FIELD_STYLE)
        self.entry_theme.grid(row=1, column=1, sticky=tk.EW, **g)

        lf_ui = ttk.Labelframe(tab_more, text="Оформление интерфейса", padding=6)
        lf_ui.grid(row=2, column=0, sticky=tk.EW, **g)
        lf_ui.columnconfigure(1, weight=1)
        ttk.Label(lf_ui, text="Цветовая схема:").grid(row=0, column=0, sticky=tk.W, **g)
        self._ui_scheme_ids = [sid for sid, _lbl in color_scheme_choices()]
        scheme_labels = [lbl for _sid, lbl in color_scheme_choices()]
        self.cmb_ui_scheme = ttk.Combobox(
            lf_ui,
            values=scheme_labels,
            state="readonly",
            width=44,
            style=FIELD_COMBO_STYLE,
        )
        self.cmb_ui_scheme.grid(row=0, column=1, sticky=tk.EW, **g)
        cur = current_color_scheme_id()
        if cur in self._ui_scheme_ids:
            self.cmb_ui_scheme.current(self._ui_scheme_ids.index(cur))
        self.cmb_ui_scheme.bind("<<ComboboxSelected>>", lambda _e: self._apply_ui_color_scheme_from_admin())
        ui_btns = ttk.Frame(lf_ui)
        ui_btns.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        ttk.Button(
            ui_btns,
            text="Применить схему",
            command=self._apply_ui_color_scheme_from_admin,
            style="Accent.TButton",
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Label(
            lf_ui,
            text="Кнопки получают фон и рамку; основные действия — синяя кнопка. "
            "Тот же выбор — меню «Вид» → «Цветовая схема».",
            wraplength=520,
            style="Hint.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))

        jf = ttk.Labelframe(tab_more, text="Журналы протоколов", padding=6)
        jf.grid(row=3, column=0, sticky=tk.W, **g)
        ttk.Button(
            jf,
            text="Журнал (охрана труда)…",
            command=lambda: self.show_protocol_journal(PROTOCOL_JOURNAL_KIND_OT),
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Button(
            jf,
            text="Журнал (тех. протоколы)…",
            command=lambda: self.show_protocol_journal(PROTOCOL_JOURNAL_KIND_TECH),
        ).grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        ttk.Button(jf, text="Очистить все журналы…", command=self.clear_protocol_database).grid(
            row=0, column=2, sticky=tk.W
        )

    def _refresh_mintrud_registry_label(self) -> None:
        if not hasattr(self, "lbl_mintrud_registry"):
            return
        s = load_mintrud_trained_registry_path().strip()
        if not s:
            self.lbl_mintrud_registry.configure(
                text="Файл не выбран — в протокол идут только номера из поля вручную.",
            )
            return
        p = Path(s).expanduser()
        if p.is_file():
            self.lbl_mintrud_registry.configure(text=f"Используется: {p}")
        else:
            self.lbl_mintrud_registry.configure(
                text=f"Путь сохранён, файл не найден: {s}",
            )

    def _should_offer_mintrud_registry_reform(self) -> bool:
        """Есть ли уже сформированный протокол в окне — имеет смысл предложить обновить номера реестра."""
        if self._last_export_persons:
            return True
        return bool((self._preview_plain_text or "").strip())

    def regenerate_protocol_with_mintrud_registry(self) -> None:
        """
        Повторно строит протокол из текущих полей и шаблона, заново читая файл реестра Минтруда
        (после обновления выгрузки с сайта — подставляются регистрационные номера).
        """
        if self.var_technical_protocol.get():
            messagebox.showinfo(
                "Реестр Минтруда",
                "Для протокола по техническим вопросам выгрузка реестра с портала не используется.",
                parent=self,
            )
            return
        if self._trained_registry_path_for_protocol() is None:
            messagebox.showwarning(
                "Реестр Минтруда",
                "Сначала укажите доступный файл выгрузки реестра:\n"
                "меню «Минтруд» → «Файл реестра обученных (.xlsx)…».",
                parent=self,
            )
            return
        self.generate_protocol()

    def _pick_mintrud_trained_registry(self, parent: tk.Misc | None = None) -> None:
        par = parent or self
        path = filedialog.askopenfilename(
            parent=par,
            title="Выгрузка реестра обученных (Excel)",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        save_mintrud_trained_registry_path(path)
        self._refresh_mintrud_registry_label()
        idx = load_trained_registry_index(Path(path))
        if idx is None:
            messagebox.showwarning(
                "Реестр",
                "Файл сохранён, но не удалось разобрать таблицу: на первом листе нужны «Номер в реестре» "
                "и ФИО (одна колонка или «Фамилия», «Имя», «Отчество»), по возможности СНИЛС, программа и номер протокола.",
                parent=par,
            )
            return
        if self._should_offer_mintrud_registry_reform():
            if messagebox.askyesno(
                "Реестр обновлён",
                "Пересформировать протокол сейчас, чтобы подставить регистрационные номера "
                "из выбранного файла реестра?\n\n"
                "Те же сотрудники и программы, что на экране; файл реестра читается заново.\n\n"
                "Для сохранения «по одному в папку» снова нажмите «По одному на каждого → в папку…».",
                parent=par,
            ):
                self.generate_protocol()

    def _clear_mintrud_trained_registry(self) -> None:
        save_mintrud_trained_registry_path("")
        self._refresh_mintrud_registry_label()

    def _trained_registry_path_for_protocol(self) -> Path | None:
        if self.var_technical_protocol.get():
            return None
        s = load_mintrud_trained_registry_path().strip()
        if not s:
            return None
        p = Path(s).expanduser()
        return p if p.is_file() else None

    def _table_fill_warning_text(self) -> str:
        if self.var_technical_protocol.get():
            return (
                "В документе не найдена таблица результатов (две строки шапки с «п/п» и «Фамилия», "
                "строка-образец с «ФИО» во второй колонке) — список сотрудников не подставлен. "
                "Проверьте шаблон .docx."
            )
        return (
            "В документе не найдена таблица с маркерными строками «1.1», «2.1», … "
            "и текстом «ФИО» во второй колонке — список сотрудников в таблицу не подставлен. "
            "Проверьте шаблон .docx."
        )

    def _active_protocol_template_path(self) -> Path:
        """Шаблон для формирования: для тех. режима — отдельный путь или встроенный тех. .docx."""
        tech_ov: Path | None = None
        if self.var_technical_protocol.get():
            tech_ov = self.technical_template_path
        return resolve_protocol_template_path(
            technical_protocol=self.var_technical_protocol.get(),
            user_override=self.template_path,
            technical_user_override=tech_ov,
        )

    def _ensure_bundle_templates_protected(self) -> None:
        if not load_protect_bundle_templates_enabled():
            return
        try:
            protect_standard_protocol_templates(application_exe_dir())
        except OSError:
            pass

    def _on_protect_bundle_templates_toggle(self) -> None:
        save_protect_bundle_templates_enabled(self.var_protect_bundle_templates.get())
        if self.var_protect_bundle_templates.get():
            self._ensure_bundle_templates_protected()

    def _protect_bundle_templates_clicked(self) -> None:
        self.var_protect_bundle_templates.set(True)
        save_protect_bundle_templates_enabled(True)
        try:
            done = protect_standard_protocol_templates(application_exe_dir())
        except OSError as e:
            messagebox.showerror("Защита шаблонов", str(e), parent=self)
            return
        if not done:
            messagebox.showinfo(
                "Защита шаблонов",
                "Стандартные шаблоны (default_protocol*.docx) не найдены в папке программы.",
                parent=self,
            )
            return
        messagebox.showinfo(
            "Защита шаблонов",
            "Включена защита «только чтение» в Word для стандартных шаблонов:\n\n"
            + "\n".join(done[:6])
            + ("\n…" if len(done) > 6 else "")
            + "\n\nСформированные протоколы не затрагиваются. "
            "Программа по-прежнему читает шаблон в память.",
            parent=self,
        )

    def _template_paths_for_protection(self) -> list[Path]:
        extra: list[Path] = [
            protocol_template_path(),
            protocol_technical_template_path(),
            self._active_protocol_template_path(),
        ]
        if self.template_path is not None:
            extra.append(self.template_path)
        if self.technical_template_path is not None:
            extra.append(self.technical_template_path)
        return extra

    def _unprotect_bundle_templates_clicked(self) -> None:
        active = self._active_protocol_template_path()
        if not messagebox.askyesno(
            "Снятие защиты",
            "Снять защиту со стандартных шаблонов в папке программы?\n\n"
            f"Для формирования протоколов используется:\n{active}\n\n"
            "После снятия защиты правьте этот файл в Word (или копию с тем же содержимым). "
            "Не удаляйте маркеры {{ПОДРАЗДЕЛЕНИЕ_ПРОВЕРКИ}}, {{УТВЕРДИЛ_ПРИКАЗ}} и др.",
            parent=self,
        ):
            return
        try:
            ok, err = unprotect_standard_protocol_templates(
                application_exe_dir(),
                extra_paths=self._template_paths_for_protection(),
            )
        except OSError as e:
            messagebox.showerror("Защита шаблонов", str(e), parent=self)
            return
        if not ok and not err:
            messagebox.showinfo(
                "Защита шаблонов",
                "Стандартные шаблоны (default_protocol*.docx) не найдены в папке программы.",
                parent=self,
            )
            return
        lines: list[str] = []
        if ok:
            lines.append("Защита снята:\n" + "\n".join(ok))
        if err:
            lines.append(
                "Не удалось (закройте файл в Word и нажмите снова):\n" + "\n".join(err)
            )
        lines.append(f"\nФайл для формирования протоколов:\n{active}")
        messagebox.showinfo("Защита шаблонов", "\n\n".join(lines), parent=self)

    def _open_template_folder_clicked(self) -> None:
        folder = protocol_template_path().parent
        if not folder.is_dir():
            messagebox.showerror(
                "Шаблон протокола",
                f"Папка не найдена:\n{folder}",
                parent=self,
            )
            return
        try:
            if sys.platform == "win32":
                os.startfile(folder)  # noqa: S606 — открыть проводник
            else:
                import subprocess

                subprocess.run(["xdg-open", str(folder)], check=False)
        except OSError as e:
            messagebox.showerror("Шаблон протокола", str(e), parent=self)

    def _template_status_text(self) -> str:
        """Подпись для шаблона ОТ / общего (не тех.): поле «Выбрать шаблон» в админке."""
        if self.template_path is not None:
            return f"Шаблон ОТ / общий: {self.template_path}"
        bundle_dir = protocol_template_path().parent
        return (
            f"Шаблон ОТ / общий: {PROTOCOL_TEMPLATE_FILENAME} "
            f"(папка с программой: {bundle_dir})"
        )

    def _sync_technical_template_path_from_settings(self) -> None:
        raw = load_tech_protocol_template_docx_path().strip()
        if not raw:
            self.technical_template_path = None
            return
        self.technical_template_path = Path(raw).expanduser()

    def _technical_template_caption(self) -> str:
        bundle_dir = protocol_template_path().parent
        if self.technical_template_path is not None:
            p = self.technical_template_path.expanduser().resolve()
            if p.is_file():
                return f"Тех. шаблон Word: {p}"
            return f"Тех. шаблон (файл не найден): {self.technical_template_path}"
        tech = protocol_technical_template_path()
        if tech.is_file():
            return (
                f"Тех. шаблон по умолчанию: {PROTOCOL_TEMPLATE_TECH_FILENAME} "
                f"(папка: {bundle_dir})"
            )
        return (
            f"Файл {PROTOCOL_TEMPLATE_TECH_FILENAME} не найден — для тех. протокола будет "
            f"{PROTOCOL_TEMPLATE_FILENAME}"
        )

    def _refresh_technical_template_labels(self) -> None:
        cap = self._technical_template_caption()
        if hasattr(self, "lbl_technical_template_main"):
            self.lbl_technical_template_main.configure(text=cap)
        if hasattr(self, "lbl_technical_template_admin"):
            self.lbl_technical_template_admin.configure(text=cap)
        if self.var_technical_protocol.get():
            self.after_idle(self._fit_main_window_to_form)

    def pick_technical_template(self) -> None:
        path = filedialog.askopenfilename(
            parent=self._admin_win or self,
            title="Шаблон Word для технического протокола",
            filetypes=[
                ("Документ Word", "*.docx"),
                ("Все файлы", "*.*"),
            ],
        )
        if not path:
            return
        p = Path(path).expanduser().resolve()
        self.technical_template_path = p
        save_tech_protocol_template_docx_path(str(p))
        self._refresh_technical_template_labels()

    def clear_technical_template(self) -> None:
        self.technical_template_path = None
        save_tech_protocol_template_docx_path("")
        self._refresh_technical_template_labels()

    def show_template_variables_help(self) -> None:
        win = self._themed_toplevel()
        win.title("Переменные и маркеры шаблона протокола")
        win.minsize(560, 420)
        frm = ttk.Frame(win, padding=8)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        box = tk.Text(frm, wrap=tk.WORD)
        configure_readonly_text(box)
        sb = ttk.Scrollbar(frm, command=box.yview)
        box.configure(yscrollcommand=sb.set)
        box.grid(row=0, column=0, sticky=tk.NSEW)
        sb.grid(row=0, column=1, sticky=tk.NS)
        box.insert("1.0", PROTOCOL_TEMPLATE_VARIABLES_DOC)
        box.configure(state=tk.DISABLED)
        ttk.Button(frm, text="Закрыть", command=lambda: self._close_modal_window(win)).grid(
            row=1, column=0, columnspan=2, pady=(10, 0)
        )
        self._make_modal(win, parent=self._dialog_parent())

    def clear_protocol_database(self) -> None:
        """Удаление всех записей журналов protocols (ОТ и тех.) с подтверждением."""
        if not messagebox.askyesno(
            "Очистка журналов",
            "Удалить все записи обоих журналов протоколов из базы (охрана труда и тех. протоколы)?\n\n"
            "Действие нельзя отменить.",
        ):
            return
        try:
            n = clear_protocol_journal()
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e))
            return
        messagebox.showinfo("Журналы", f"Удалено записей: {n}.")

    def show_protocol_journal(self, journal_kind: str = PROTOCOL_JOURNAL_KIND_OT) -> None:
        """Окно со списком сохранённых протоколов из SQLite и текстом записи."""
        kind = (journal_kind or PROTOCOL_JOURNAL_KIND_OT).strip() or PROTOCOL_JOURNAL_KIND_OT
        try:
            rows = get_protocols_journal_display(protocol_kind=kind)
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e))
            return

        win = self._themed_toplevel()
        modal_parent = self._dialog_parent()
        jtitle = (
            "Журнал технических протоколов"
            if kind == PROTOCOL_JOURNAL_KIND_TECH
            else "Журнал протоколов (охрана труда)"
        )
        win.title(f"{jtitle} — {DATABASE_FILENAME}")
        win.minsize(720, 520)
        win.geometry("900x600")

        outer = ttk.Frame(win, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.rowconfigure(3, weight=1)
        outer.columnconfigure(0, weight=1)

        dbp = database_path()
        ttk.Label(
            outer,
            text=f"Файл базы: {dbp}",
            style="Hint.TLabel",
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        ttk.Label(
            outer,
            text=(
                "В базе не хранится полный текст протокола — только ФИО, тема, оценка, дата и "
                "№ протокола (как в бланке: номер-месяц-год). Старые записи без № показываются как «—». "
                "«Выгрузить реестр…» — Excel или CSV; протокол на нескольких человек "
                "выгружается отдельной строкой на каждое ФИО."
            ),
            wraplength=860,
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky=tk.W, pady=(0, 4))

        btn_bar = ttk.Frame(outer)
        btn_bar.grid(row=2, column=0, sticky=tk.W, pady=(0, 6))

        pw = ttk.Panedwindow(outer, orient=tk.VERTICAL)
        pw.grid(row=3, column=0, sticky=tk.NSEW)
        outer.rowconfigure(3, weight=1)

        top_fr = ttk.Frame(pw, padding=2)
        bot_fr = ttk.Frame(pw, padding=2)
        pw.add(top_fr, weight=1)
        pw.add(bot_fr, weight=2)

        top_fr.rowconfigure(0, weight=1)
        top_fr.columnconfigure(0, weight=1)
        sb_list = ttk.Scrollbar(top_fr)
        lb = tk.Listbox(
            top_fr,
            height=8,
            yscrollcommand=sb_list.set,
            exportselection=False,
        )
        configure_listbox(lb, mono=True)
        lb.grid(row=0, column=0, sticky=tk.NSEW)
        sb_list.grid(row=0, column=1, sticky=tk.NS)
        sb_list.configure(command=lb.yview)

        bot_fr.rowconfigure(1, weight=1)
        bot_fr.columnconfigure(0, weight=1)
        ttk.Label(bot_fr, text="Просмотр (если текст не сохранялся в журнал — см. подсказку ниже):").grid(
            row=0, column=0, sticky=tk.W
        )
        sb_txt = ttk.Scrollbar(bot_fr)
        txt = tk.Text(bot_fr, wrap=tk.WORD, height=14, state=tk.DISABLED)
        configure_readonly_text(txt)
        txt.grid(row=1, column=0, sticky=tk.NSEW)
        sb_txt.grid(row=1, column=1, sticky=tk.NS)
        sb_txt.configure(command=txt.yview)
        txt.configure(yscrollcommand=sb_txt.set)

        def refresh_list() -> None:
            nonlocal rows
            try:
                rows = get_protocols_journal_display(protocol_kind=kind)
            except sqlite3.Error as e:
                messagebox.showerror("База данных", str(e))
                return
            lb.delete(0, tk.END)
            for r in rows:
                lb.insert(tk.END, format_journal_list_line(r))
            txt.configure(state=tk.NORMAL)
            txt.delete("1.0", tk.END)
            txt.configure(state=tk.DISABLED)

        def on_select(_evt: object | None = None) -> None:
            sel = lb.curselection()
            if not sel:
                return
            r = rows[int(sel[0])]
            pn = export_meta_protocol_no(r.get("export_meta_json"))
            header = f"№ протокола: {pn or '—'}\nДата: {(r.get('date') or '').strip()}\n\n"
            body = (r.get("content") or "").strip()
            if not body:
                body = JOURNAL_PLACEHOLDER_NO_BODY
            txt.configure(state=tk.NORMAL)
            txt.delete("1.0", tk.END)
            txt.insert("1.0", header + body)
            txt.configure(state=tk.DISABLED)

        def copy_to_main_preview() -> None:
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Журнал", "Выберите запись в списке.")
                return
            r = rows[int(sel[0])]
            body = (r.get("content") or "").strip()
            if not body:
                messagebox.showinfo(
                    "Журнал",
                    "Текст протокола в этой записи не сохраняется в базе.\n"
                    "Сформируйте протокол снова или откройте сохранённый DOCX/PDF.",
                )
                return
            self._set_protocol_preview(None, body, show_window=False)
            messagebox.showinfo(
                "Журнал",
                "После закрытия этого окна откроется «Предпросмотр протокола».\n"
                "Сохранение в DOCX/PDF — из того окна или с главного экрана.",
            )
            self._show_preview_toplevel()

        def on_clear_journal() -> None:
            jname = (
                "технических протоколов"
                if kind == PROTOCOL_JOURNAL_KIND_TECH
                else "протоколов охраны труда"
            )
            if not messagebox.askyesno(
                "Очистка журнала",
                f"Удалить все записи журнала ({jname}) из базы?\n\nДействие нельзя отменить.",
                parent=win,
            ):
                return
            try:
                n = clear_protocol_journal(kind)
            except sqlite3.Error as e:
                messagebox.showerror("База данных", str(e), parent=win)
                return
            refresh_list()
            messagebox.showinfo("Журнал", f"Удалено записей: {n}.", parent=win)

        def on_export_registry() -> None:
            if not rows:
                messagebox.showinfo(
                    "Выгрузка реестра",
                    "В журнале нет записей для выгрузки.",
                    parent=win,
                )
                return
            default_name = default_journal_registry_export_path(journal_kind=kind)
            out = filedialog.asksaveasfilename(
                parent=win,
                title="Сохранить реестр сформированных протоколов",
                initialdir=str(default_name.parent),
                initialfile=default_name.name,
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel", "*.xlsx"),
                    ("CSV (разделитель ;)", "*.csv"),
                    ("Все файлы", "*.*"),
                ],
            )
            if not out:
                return
            try:
                n_lines = export_protocol_journal_registry(Path(out), rows)
            except (OSError, RuntimeError, sqlite3.Error) as e:
                messagebox.showerror("Выгрузка реестра", str(e), parent=win)
                return
            extra = ""
            if n_lines != len(rows):
                extra = (
                    f"\n\nВ журнале записей: {len(rows)}; в файле строк: {n_lines} "
                    "(протоколы на нескольких человек разбиты по ФИО)."
                )
            messagebox.showinfo(
                "Выгрузка реестра",
                f"Сохранено строк: {n_lines}{extra}\n{out}",
                parent=win,
            )

        ttk.Button(btn_bar, text="Обновить список", command=refresh_list).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(btn_bar, text="Выгрузить реестр…", command=on_export_registry).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(btn_bar, text="В окно предпросмотра", command=copy_to_main_preview).grid(
            row=0, column=2, padx=(0, 8)
        )
        ttk.Button(btn_bar, text="Очистить журнал…", command=on_clear_journal).grid(
            row=0, column=3, padx=(0, 8)
        )
        ttk.Button(btn_bar, text="Закрыть", command=lambda: self._close_modal_window(win, parent=modal_parent)).grid(
            row=0, column=4
        )

        lb.bind("<<ListboxSelect>>", on_select)
        for r in rows:
            lb.insert(tk.END, format_journal_list_line(r))
        if rows:
            lb.selection_set(0)
            on_select()
        self._make_modal(win, parent=modal_parent)

    def _open_mintrud_employer_window(self) -> None:
        """Реквизиты работодателя и организации 2 для шаблона выгрузки Минтруда (protocols.db)."""
        win = self._themed_toplevel()
        win.title("Минтруд — реквизиты для шаблона")
        win.minsize(440, 340)
        win.resizable(True, False)

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)

        try:
            inn0, name0, inn2_0, org2_0 = load_mintrud_employer_from_db()
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e), parent=win)
            win.destroy()
            return

        lf = ttk.Labelframe(
            outer,
            text="Работодатель (инн работодателя, название работодателя)",
            padding=8,
        )
        lf.grid(row=0, column=0, sticky=tk.EW, pady=(0, 8))
        lf.columnconfigure(1, weight=1)

        ttk.Label(lf, text="ИНН работодателя:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        ent_inn = ttk.Entry(lf, width=48, style=FIELD_STYLE)
        ent_inn.grid(row=0, column=1, sticky=tk.EW, pady=4)
        ent_inn.insert(0, inn0)

        ttk.Label(lf, text="Название работодателя:").grid(row=1, column=0, sticky=tk.NW, padx=(0, 8), pady=4)
        ent_name = ttk.Entry(lf, width=48, style=FIELD_STYLE)
        ent_name.grid(row=1, column=1, sticky=tk.EW, pady=4)
        ent_name.insert(0, name0)

        lf2 = ttk.Labelframe(
            outer,
            text="Дополнительно: организация 2 (инн организации2, наименование организации2)",
            padding=8,
        )
        lf2.grid(row=1, column=0, sticky=tk.EW, pady=(0, 8))
        lf2.columnconfigure(1, weight=1)

        ttk.Label(lf2, text="ИНН организации 2:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        ent_inn2 = ttk.Entry(lf2, width=48, style=FIELD_STYLE)
        ent_inn2.grid(row=0, column=1, sticky=tk.EW, pady=4)
        ent_inn2.insert(0, inn2_0)

        ttk.Label(lf2, text="Наименование организации 2:").grid(
            row=1, column=0, sticky=tk.NW, padx=(0, 8), pady=4
        )
        ent_org2 = ttk.Entry(lf2, width=48, style=FIELD_STYLE)
        ent_org2.grid(row=1, column=1, sticky=tk.EW, pady=4)
        ent_org2.insert(0, org2_0)

        ttk.Label(
            outer,
            text="Все поля подставляются в каждую строку при выгрузке «Шаблон для загрузки на сайт…».",
            style="Hint.TLabel",
            wraplength=420,
        ).grid(row=2, column=0, sticky=tk.W, pady=(0, 8))

        def do_save() -> None:
            try:
                save_mintrud_employer_to_db(
                    ent_inn.get(),
                    ent_name.get(),
                    ent_inn2.get(),
                    ent_org2.get(),
                )
            except sqlite3.Error as e:
                messagebox.showerror("База данных", str(e), parent=win)
                return
            messagebox.showinfo("Минтруд", "Реквизиты сохранены в базе.", parent=win)

        bf = ttk.Frame(outer)
        bf.grid(row=3, column=0, sticky=tk.E)
        ttk.Button(bf, text="Сохранить", command=do_save).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(bf, text="Закрыть", command=lambda: self._close_modal_window(win)).grid(row=0, column=1)
        self._register_clipboard_for_window(win)
        self._make_modal(win)

    def _open_mintrud_export_window(self) -> None:
        """История журнала: выбор одной или нескольких записей, выгрузка шаблона Excel для реестра Минтруда."""
        try:
            rows = get_protocols_journal_display(protocol_kind=PROTOCOL_JOURNAL_KIND_OT)
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e))
            return

        win = self._themed_toplevel()
        win.title("Шаблон для сайта Минтруда — выбор из журнала протоколов")
        win.minsize(760, 520)
        win.geometry("920x580")

        outer = ttk.Frame(win, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.rowconfigure(2, weight=1)
        outer.columnconfigure(0, weight=1)

        dbp = database_path()
        ttk.Label(
            outer,
            text=(
                "Выберите одну или несколько записей журнала (Ctrl/Shift — множественный выбор). "
                f"База: {dbp}\n\n"
                "Файл для выгрузки строится на основе официального шаблона "
                "«Шаблон_Минтруд_XSD_УМН.xlsx» в папке с программой (лист «Шаблон»). "
                "СНИЛС в журнале не хранится; при выгрузке подставляется только из файла сотрудников (Excel) по ФИО. "
                "Должность — из журнала (метаданные) или из того же файла, если совпало ФИО. "
                "Реквизиты работодателя и организации 2 — «Минтруд» → «Реквизиты работодателя…». "
                "В списке — по одной актуальной записи на протокол (без дублей). "
                "Повторная выгрузка формирует новый файл из выбранных записей, "
                "не дополняет ранее сохранённый Excel."
            ),
            wraplength=880,
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 6))

        btn_bar = ttk.Frame(outer)
        btn_bar.grid(row=1, column=0, sticky=tk.W, pady=(0, 6))

        list_fr = ttk.Frame(outer)
        list_fr.grid(row=2, column=0, sticky=tk.NSEW)
        list_fr.rowconfigure(0, weight=1)
        list_fr.columnconfigure(0, weight=1)
        sb_list = ttk.Scrollbar(list_fr)
        lb = tk.Listbox(
            list_fr,
            height=16,
            yscrollcommand=sb_list.set,
            exportselection=False,
            selectmode=tk.EXTENDED,
        )
        configure_listbox(lb, mono=True)
        lb.grid(row=0, column=0, sticky=tk.NSEW)
        sb_list.grid(row=0, column=1, sticky=tk.NS)
        sb_list.configure(command=lb.yview)

        lbl_empty = ttk.Label(
            list_fr,
            text=(
                "Записей в журнале пока нет. Сформируйте протокол по охране труда — "
                "данные попадут в журнал автоматически. Затем нажмите «Обновить список»."
            ),
            wraplength=860,
            style="Muted.TLabel",
        )

        def refresh_list() -> None:
            nonlocal rows
            try:
                rows = get_protocols_journal_display(protocol_kind=PROTOCOL_JOURNAL_KIND_OT)
            except sqlite3.Error as e:
                messagebox.showerror("База данных", str(e), parent=win)
                return
            lb.delete(0, tk.END)
            if rows:
                lbl_empty.grid_remove()
                for r in rows:
                    lb.insert(tk.END, format_journal_list_line(r))
            else:
                lbl_empty.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))

        def select_all() -> None:
            if lb.size() == 0:
                return
            lb.selection_set(0, tk.END)

        def select_none() -> None:
            lb.selection_clear(0, tk.END)

        def show_mintrud_help() -> None:
            hw = self._themed_toplevel(win)
            self._apply_embedded_window_icon(hw)
            hw.title("Справка: выгрузка для Минтруда")
            hw.minsize(520, 360)
            f = ttk.Frame(hw, padding=10)
            f.pack(fill=tk.BOTH, expand=True)
            t = tk.Text(f, wrap=tk.WORD, height=16)
            configure_readonly_text(t)
            t.pack(fill=tk.BOTH, expand=True)
            t.insert(
                "1.0",
                "Сохраняется копия официального Excel-шаблона XSD "
                "«Шаблон_Минтруд_XSD_УМН.xlsx» (лист «Шаблон») с заполненными строками "
                "и одноимённый файл .xml с теми же данными (проверка, архив, дальнейшая обработка). "
                "Шаблон Excel должен лежать в папке с программой.\n\n"
                "Заполняется из журнала: наименование программы (тема в журнале), фамилия/имя/отчество "
                "(несколько сотрудников из одной записи — отдельные строки при перечислении ФИО "
                "через запятую), дата в удостоверении (дата из формы протокола), номер протокола "
                "(по строке «ПРОТОКОЛ №» в тексте записи), колонка «тест пройден» — да/нет по оценке.\n\n"
                "ИНН и название работодателя, ИНН и наименование организации 2 — меню "
                "«Минтруд» → «Реквизиты работодателя…». СНИЛС только из файла сотрудников "
                "(столбец «СНИЛС») по совпадению ФИО. "
                "Должности и программы — из метаданных протокола (снимок при сохранении в журнал); "
                "для старых записей без должностей — из файла сотрудников и Programs_base.\n\n"
                "«Тест пройден»: 1 — удовлетворительно, 0 — неудовлетворительно.\n\n"
                "Список записей — без дублей (актуальная версия каждого протокола). "
                "Сохранённый Excel не дополняется автоматически — каждый раз формируется "
                "из выбранных строк журнала.\n\n"
                "При необходимости вручную: ID программы в реестре — по инструкции портала "
                "(например https://akot.rosmintrud.ru/ ).",
            )
            t.configure(state=tk.DISABLED)
            ttk.Button(f, text="Закрыть", command=hw.destroy).pack(pady=(8, 0))

        def export_template() -> None:
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo(
                    "Выгрузка",
                    "Выберите в списке одну или несколько записей журнала.",
                    parent=win,
                )
                return
            chosen = [rows[int(i)] for i in sel]
            default_name = f"mintrud_shablon_{date.today().strftime('%Y%m%d')}.xlsx"
            path = filedialog.asksaveasfilename(
                parent=win,
                title="Сохранить шаблон Excel и XML",
                defaultextension=".xlsx",
                initialdir=str(mintrud_export_output_dir()),
                initialfile=default_name,
                filetypes=[
                    ("Книга Excel", "*.xlsx"),
                    ("Все файлы", "*.*"),
                ],
            )
            if not path:
                return
            try:
                inn_e, nm_e, inn2_e, org2_e = load_mintrud_employer_from_db()
            except sqlite3.Error as e:
                messagebox.showerror("База данных", str(e), parent=win)
                return
            try:
                write_mintrud_template_xlsx(
                    Path(path),
                    chosen,
                    inn_employer=inn_e,
                    employer_name=nm_e,
                    inn_org2=inn2_e,
                    org2_name=org2_e,
                    employees_excel_path=self._employees_file_resolved(),
                    programs_excel_path=self._programs_file_resolved(),
                    v_parts_for_employee=v_program_merged_parts_for_raw_employee,
                )
            except FileNotFoundError as e:
                messagebox.showerror("Шаблон Минтруда", str(e), parent=win)
                return
            except ValueError as e:
                messagebox.showerror("Шаблон Минтруда", str(e), parent=win)
                return
            except OSError as e:
                messagebox.showerror("Файл", f"Не удалось сохранить:\n{e}", parent=win)
                return
            except RuntimeError as e:
                messagebox.showerror("Выгрузка", str(e), parent=win)
                return
            except Exception as e:
                messagebox.showerror("Выгрузка", f"Ошибка при формировании файла:\n{e}", parent=win)
                return
            xml_p = str(Path(path).with_suffix(".xml"))
            messagebox.showinfo(
                "Готово",
                f"Сохранено:\n{path}\n{xml_p}\n\n"
                "Дополните при необходимости пустые колонки в Excel и проверьте соответствие инструкции портала.",
                parent=win,
            )

        ttk.Button(btn_bar, text="Обновить список", command=refresh_list).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(btn_bar, text="Выбрать всё", command=select_all).grid(
            row=0, column=1, padx=(0, 6)
        )
        ttk.Button(btn_bar, text="Снять выбор", command=select_none).grid(
            row=0, column=2, padx=(0, 6)
        )
        ttk.Button(btn_bar, text="Сформировать шаблон Excel…", command=export_template).grid(
            row=0, column=3, padx=(0, 6)
        )
        ttk.Button(btn_bar, text="Справка…", command=show_mintrud_help).grid(
            row=0, column=4, padx=(0, 6)
        )
        ttk.Button(btn_bar, text="Закрыть", command=lambda: self._close_modal_window(win)).grid(row=0, column=5)

        if rows:
            for r in rows:
                lb.insert(tk.END, format_journal_list_line(r))
        else:
            lbl_empty.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))
        self._make_modal(win)

    def pick_template(self) -> None:
        path = filedialog.askopenfilename(
            parent=self._dialog_parent(),
            title="Выберите шаблон протокола",
            filetypes=[
                ("Текст UTF-8", "*.txt"),
                ("Документ Word", "*.docx"),
                ("Все файлы", "*.*"),
            ],
        )
        if not path:
            return
        self.template_path = Path(path).expanduser().resolve()
        self.lbl_template.configure(text=self._template_status_text())

    def _refresh_employees_file_label(self) -> None:
        p = self._employees_file_resolved()
        src = (
            "выбранный файл"
            if self.employees_excel_path
            else f"по умолчанию ({_default_user_data_hint()})"
        )
        extra = ""
        if not p.is_file():
            extra = (
                f" — файл не найден (положите {EMPLOYEES_EXCEL_FILENAME} в {_default_user_data_hint()} "
                "или укажите «Файл сотрудников…»)"
            )
        elif self._employee_records:
            extra = f" — загружено: {len(self._employee_records)}"
        else:
            extra = (
                " — список пуст или файл не прочитан (кнопка «Загрузить из Excel» покажет ошибку; "
                "проверьте лист и первую строку заголовков)"
            )
        self.lbl_employees_file.configure(text=f"{src}: {p.name}{extra}")

    def pick_employees_excel(self) -> None:
        par = self._dialog_parent()
        path = filedialog.askopenfilename(
            parent=par,
            title="Файл с базой сотрудников",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("Excel с макросами", "*.xlsm"),
                ("Все файлы", "*.*"),
            ],
        )
        if not path:
            return
        self.employees_excel_path = Path(path).expanduser().resolve()
        self._refresh_employees_file_label()
        self.reload_employees(show_errors=True, notify_success=True, parent=par)

    def _reload_employees_from_admin(self) -> None:
        self.reload_employees(
            show_errors=True,
            notify_success=True,
            parent=self._admin_win or self,
        )

    def _refresh_programs_file_label(self) -> None:
        if not hasattr(self, "lbl_programs_file"):
            return
        p = self._programs_file_resolved()
        same_as_emp = p.resolve() == self._employees_file_resolved().resolve()
        if self.programs_excel_path is not None:
            src = "выбранный файл"
        elif same_as_emp:
            src = f"как сотрудники ({EMPLOYEES_EXCEL_FILENAME})"
        else:
            src = f"по умолчанию {PROGRAMS_EXCEL_FILENAME}"
        extra = ""
        if not p.is_file():
            extra = f" — не найден (положите в {_default_user_data_hint()} или укажите файл)"
        self.lbl_programs_file.configure(text=f"{src}: {p.name}{extra}")

    def pick_programs_excel(self) -> None:
        path = filedialog.askopenfilename(
            parent=self._admin_win or self,
            title="Файл справочника программ (B, V_PROF, PP, SIZ, V)",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("Excel с макросами", "*.xlsm"),
                ("Все файлы", "*.*"),
            ],
        )
        if not path:
            return
        self.programs_excel_path = Path(path).expanduser().resolve()
        invalidate_program_catalog_cache_for_path(self._programs_file_resolved())
        self._refresh_programs_file_label()

    def _export_programs_workbook_from_combined(self) -> None:
        par = self._admin_win or self
        src = filedialog.askopenfilename(
            parent=par,
            title="Исходный объединённый Excel (с листами программ)",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if not src:
            return
        dst = filedialog.asksaveasfilename(
            parent=par,
            title="Сохранить справочник программ как",
            defaultextension=".xlsx",
            initialfile=PROGRAMS_EXCEL_FILENAME,
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if not dst:
            return
        try:
            copy_program_sheets_from_workbook(Path(src), Path(dst))
        except EmployeeExcelError as e:
            messagebox.showerror("Листы программ", str(e), parent=par)
            return
        messagebox.showinfo(
            "Готово",
            f"Создан файл программ:\n{dst}\n\nПри необходимости укажите его кнопкой «Файл программ…».",
            parent=par,
        )

    def _refresh_data_bases_clicked(self) -> None:
        emp = self._employees_file_resolved()
        prog = self._programs_file_resolved()
        try:
            invalidate_employees_cache_for_path(emp)
        except OSError:
            pass
        try:
            invalidate_program_catalog_cache_for_path(prog)
        except OSError:
            pass
        refresh_commission_pool_from_excel(
            self._commission_state,
            emp,
            show_errors=False,
            parent=self,
        )
        self._sync_and_refresh_commission_pool_listboxes()
        self.reload_employees(show_errors=False)
        self._refresh_employees_file_label()
        self._refresh_programs_file_label()
        self._refresh_tech_v_program_combo(silent=True)
        messagebox.showinfo(
            "Базы обновлены",
            "Кэш сброшен, сотрудники и справочник программ перечитаны с диска "
            "(Data_base / Programs_base).",
        )

    def reload_employees(
        self,
        *,
        show_errors: bool = False,
        notify_success: bool = False,
        parent: tk.Misc | None = None,
    ) -> None:
        par = parent if parent is not None else self._dialog_parent()
        path = self._employees_file_resolved()
        try:
            self._employee_records = load_employees_from_excel(path)
        except EmployeeExcelError as e:
            self._employee_records = []
            self._employee_list_slot_gi = []
            self._employee_search_blobs = []
            self.list_employees.delete(0, tk.END)
            self._refresh_employees_file_label()
            if show_errors:
                messagebox.showerror("Сотрудники Excel", str(e), parent=par)
            self._status_var.set(
                f"Ошибка загрузки сотрудников ({path.name}). Откройте «Настройки» и проверьте файл."
            )
            return
        sort_employees_by_subdivision_then_fio(self._employee_records)
        self._rebuild_employee_search_blobs()
        self._refilter_employee_list()
        self._refresh_employees_file_label()
        refresh_commission_pool_from_excel(
            self._commission_state,
            self._employees_file_resolved(),
            show_errors=False,
            parent=par,
        )
        self._sync_and_refresh_commission_pool_listboxes()
        save_employees_cache(path, self._employee_records)
        if notify_success:
            n = len(self._employee_records)
            if n:
                messagebox.showinfo(
                    "Сотрудники Excel",
                    f"Загружено записей: {n}\n{path}",
                    parent=par,
                )
            else:
                messagebox.showwarning(
                    "Сотрудники Excel",
                    f"Файл прочитан, но записей не найдено.\n{path}\n\n"
                    "Проверьте лист сотрудников и строку заголовков (ФИО, должность).",
                    parent=par,
                )

    def _rebuild_employee_search_blobs(self) -> None:
        """Один раз на загрузку: нижний регистр для быстрого поиска по списку."""
        self._employee_search_blobs = [
            (
                f"{rec.fio} {rec.profession} {rec.subdivision} "
                f"{rec.profession2} {rec.snils}"
            ).lower()
            for rec in self._employee_records
        ]

    def _schedule_refilter_employee_list(self) -> None:
        if self._after_refilter_id is not None:
            try:
                self.after_cancel(self._after_refilter_id)
            except (ValueError, tk.TclError):
                pass
        self._after_refilter_id = self.after(140, self._refilter_employee_list_run)

    def _refilter_employee_list_run(self) -> None:
        self._after_refilter_id = None
        self._refilter_employee_list()

    def _build_employee_list_slots(self, global_indices: list[int]) -> None:
        """Список сотрудников с заголовками подразделений (слот None — заголовок)."""
        if not global_indices:
            self._employee_list_slot_gi = []
            self._list_index_to_sub_header = {}
            self.list_employees.delete(0, tk.END)
            return
        records = self._employee_records
        groups: dict[str, list[int]] = {}
        group_order: list[str] = []
        for gi in global_indices:
            key = subdivision_group_key(records[gi].subdivision)
            if key not in groups:
                groups[key] = []
                group_order.append(key)
            groups[key].append(gi)

        slots: list[int | None] = []
        labels: list[str] = []
        header_map: dict[int, str] = {}
        for sub_key in group_order:
            gis = groups[sub_key]
            sub_display = (records[gis[0]].subdivision or "").strip()
            collapsed = sub_key in self._emp_collapsed_subdivisions
            header_idx = len(slots)
            slots.append(None)
            labels.append(
                listbox_subdivision_header(
                    sub_display, len(gis), collapsed=collapsed
                )
            )
            header_map[header_idx] = sub_key
            if not collapsed:
                for gi in gis:
                    slots.append(gi)
                    labels.append(
                        listbox_label_for_employee(
                            records[gi], grouped_by_subdivision=True
                        )
                    )
        self._employee_list_slot_gi = slots
        self._list_index_to_sub_header = header_map
        self.list_employees.delete(0, tk.END)
        for text in labels:
            self.list_employees.insert(tk.END, text)

    def _refilter_employee_list(self) -> None:
        if not hasattr(self, "list_employees"):
            return
        records = self._employee_records
        if len(self._employee_search_blobs) != len(records):
            self._rebuild_employee_search_blobs()
        q = self.var_emp_search.get().strip().lower()
        prev_sel_global: list[int] = []
        for li in self.list_employees.curselection():
            li = int(li)
            if 0 <= li < len(self._employee_list_slot_gi):
                gi = self._employee_list_slot_gi[li]
                if gi is not None:
                    prev_sel_global.append(gi)
        if not q:
            filtered = list(range(len(records)))
        else:
            filtered = []
            blobs = self._employee_search_blobs
            for i, _rec in enumerate(records):
                blob = blobs[i] if i < len(blobs) else ""
                if q in blob:
                    filtered.append(i)
        self._build_employee_list_slots(filtered)
        for gi in prev_sel_global:
            for pos, slot in enumerate(self._employee_list_slot_gi):
                if slot == gi:
                    self.list_employees.selection_set(pos)
        self._sync_status_bar()

    def _try_autoload_employees(self) -> None:
        path = self._employees_file_resolved()
        if not path.is_file():
            return
        cached = try_load_employees_from_cache(path)
        if cached is not None:
            self._employee_records = cached
            sort_employees_by_subdivision_then_fio(self._employee_records)
            self._rebuild_employee_search_blobs()
            self._refilter_employee_list()
            self._refresh_employees_file_label()
            refresh_commission_pool_from_excel(
                self._commission_state,
                path,
                show_errors=False,
                parent=self,
            )
            self._sync_and_refresh_commission_pool_listboxes()
            return
        self.reload_employees(show_errors=False)

    def _on_employee_list_click(self, event: tk.Event) -> str | None:
        """Свернуть/развернуть группу по клику на заголовок подразделения."""
        lb = self.list_employees
        idx = int(lb.nearest(event.y))
        if idx < 0 or idx >= len(self._employee_list_slot_gi):
            return None
        sub_key = self._list_index_to_sub_header.get(idx)
        if sub_key is None:
            return None
        if sub_key in self._emp_collapsed_subdivisions:
            self._emp_collapsed_subdivisions.discard(sub_key)
        else:
            self._emp_collapsed_subdivisions.add(sub_key)
        prev_sel_global: list[int] = []
        for li in lb.curselection():
            li = int(li)
            if 0 <= li < len(self._employee_list_slot_gi):
                gi = self._employee_list_slot_gi[li]
                if gi is not None:
                    prev_sel_global.append(gi)
        q = self.var_emp_search.get().strip().lower()
        records = self._employee_records
        if not q:
            filtered = list(range(len(records)))
        else:
            filtered = []
            blobs = self._employee_search_blobs
            for i, _rec in enumerate(records):
                blob = blobs[i] if i < len(blobs) else ""
                if q in blob:
                    filtered.append(i)
        self._build_employee_list_slots(filtered)
        for gi in prev_sel_global:
            for pos, slot in enumerate(self._employee_list_slot_gi):
                if slot == gi:
                    lb.selection_set(pos)
        return "break"

    def _on_employee_list_select(self, _event: object | None = None) -> None:
        sel = [int(i) for i in self.list_employees.curselection()]
        for li in sel:
            if li in self._list_index_to_sub_header:
                self.list_employees.selection_clear(li)
        sel = [int(i) for i in self.list_employees.curselection()]
        self._invalidate_v_prof_combination_choice()
        if len(sel) != 1:
            if len(sel) > 1:
                self._refresh_v_prof_profession_hint()
            return
        li = sel[0]
        if li < 0 or li >= len(self._employee_list_slot_gi):
            return
        gi = self._employee_list_slot_gi[li]
        if gi is None:
            return
        rec = self._employee_records[gi]
        self.entry_position.delete(0, tk.END)
        self.entry_position.insert(0, rec.profession)
        self.entry_subdivision.delete(0, tk.END)
        self.entry_subdivision.insert(0, rec.subdivision)
        self._refresh_v_prof_profession_hint(rec.profession)

    def _on_position_focus_out(self, _event: object | None = None) -> None:
        self._refresh_v_prof_profession_hint()

    def _hide_v_prof_suggest_combo(self) -> None:
        if not hasattr(self, "_v_prof_suggest_fr"):
            return
        self._v_prof_suggest_professions.clear()
        self.cmb_v_prof_suggest.set("")
        self.cmb_v_prof_suggest.configure(values=())
        self._v_prof_suggest_fr.grid_remove()

    def _v_prof_match_label_text(self, profession: str, v_count: int, *, warn: bool) -> str:
        """Одна строка без переноса — не сжимает список сотрудников."""
        short = profession.strip()
        if len(short) > 78:
            short = short[:75] + "…"
        extra = " — проверьте формулировку" if warn else ""
        return f"{V_PROF_SHEET_NAME}: «{short}», программ «В»: {v_count}{extra}"

    def _show_v_prof_suggest_combo(self, chips: list[VProfProfessionCandidate]) -> None:
        displays: list[str] = []
        self._v_prof_suggest_professions.clear()
        for c in chips:
            mark = "★ " if c.score >= 3 else ""
            disp = f"{mark}{c.profession} ({c.v_program_count} «В»)"
            displays.append(disp)
            self._v_prof_suggest_professions[disp] = c.profession
        max_len = max((len(d) for d in displays), default=MAIN_FORM_ENTRY_CHARS)
        cb_width = min(96, max(MAIN_FORM_ENTRY_CHARS, max_len + 2))
        self.cmb_v_prof_suggest.configure(values=displays, width=cb_width)
        if displays:
            self.cmb_v_prof_suggest.current(0)
        self._v_prof_suggest_fr.grid()
        self._ensure_main_window_width_for_text(*displays, *self._v_prof_suggest_professions.values())

    def _apply_v_prof_profession_from_combo(self) -> None:
        disp = self.cmb_v_prof_suggest.get().strip()
        if not disp:
            return
        profession = self._v_prof_suggest_professions.get(disp)
        if profession:
            self._apply_v_prof_profession_choice(profession)

    def _apply_v_prof_profession_choice(self, profession: str) -> None:
        self.entry_position.delete(0, tk.END)
        self.entry_position.insert(0, profession)
        self._ensure_main_window_width_for_text(profession)
        self._refresh_v_prof_profession_hint(profession)

    def _refresh_v_prof_profession_hint(self, profession: str | None = None) -> None:
        """Подсказка V_PROF и выпадающий список похожих профессий."""
        if not hasattr(self, "lbl_v_prof_match"):
            return
        self._hide_v_prof_suggest_combo()
        pr = (profession if profession is not None else self.entry_position.get()).strip()
        path = self._programs_file_resolved()
        if not pr:
            self.lbl_v_prof_match.configure(text="", foreground=Colors.text_hint)
            self._v_prof_match_tooltip._text = ""
            return
        if not path.is_file():
            self.lbl_v_prof_match.configure(
                text=f"{V_PROF_SHEET_NAME}: файл программ не найден",
                foreground=Colors.error,
            )
            return
        hint_profs = self._professions_for_v_prof_hint()
        if len(hint_profs) > 1:
            candidates = v_prof_candidates_for_profession_list(
                path, hint_profs, limit_per=4, total_limit=12
            )
        else:
            candidates = similar_professions_in_v_prof(path, pr, limit=8)
        if not candidates:
            prefix_disp = v_prof_search_prefix_display(pr)
            if prefix_disp:
                msg = (
                    f"{V_PROF_SHEET_NAME}: нет такой профессии для «{pr}» "
                    f"(по началу фразы: {prefix_disp})"
                )
            else:
                msg = f"{V_PROF_SHEET_NAME}: нет такой профессии для «{pr}»"
            self.lbl_v_prof_match.configure(text=msg, foreground=Colors.error)
            self._v_prof_match_tooltip._text = msg
            return
        best = candidates[0]
        warn = best.score < 2
        extra = " — проверьте формулировку" if warn else ""
        match_line = self._v_prof_match_label_text(
            best.profession, best.v_program_count, warn=warn
        )
        if len(hint_profs) > 1:
            match_line += f" · учтено должностей: {len(hint_profs)}"
        self.lbl_v_prof_match.configure(
            text=match_line,
            foreground=Colors.error if warn else Colors.success,
        )
        self._v_prof_match_tooltip._text = (
            f"{V_PROF_SHEET_NAME}: «{best.profession}», программ «В»: {best.v_program_count}{extra}"
        )
        current_key = pr.strip().lower()
        chips = [
            c
            for c in candidates
            if not (c.profession.strip().lower() == current_key and c.score >= 3)
        ]
        if not chips and (len(candidates) > 1 or best.score < 3):
            chips = candidates
        if not chips:
            return
        self._show_v_prof_suggest_combo(chips)

    def _collect_program_keys_and_titles(
        self, persons_raw: list[EmployeeRecord]
    ) -> tuple[list[str], list[str]]:
        """
        Порядок: Б → PP → СИЗ → В.
        Таблица протокола: «Б»/ПП/СИЗ — названия с листов B, PP, SIZ; «В» — по V_PROF и листу V в docx.
        Шапка — V_PROF (Б/ПП/СИЗ) и лист V ст. C для «В»; таблица «В» — лист V ст. B.
        """
        path = self._programs_file_resolved()
        keys: list[str] = []
        titles: list[str] = []
        for key, _, fallback in PROTOCOL_PROGRAM_DEFS:
            if not self._prog_vars[key].get():
                continue
            keys.append(key)
            if key == "V":
                v_parts_g = v_program_ordered_unique_parts_global(
                    path,
                    persons_raw,
                    face_sheet_profession=self._face_sheet_profession(),
                    persons_row_source=self._collect_table_persons(),
                    **self._v_prof_combo_kwargs(),
                )
                titles.append(format_v_program_table_block_title(v_parts_g, fallback))
            elif key == "B":
                t = get_cached_b_program_title(path).strip()
                titles.append(t if t else fallback)
            elif key == "PP":
                t = get_cached_pp_table_title(path).strip()
                titles.append(t if t else fallback)
            elif key == "SIZ":
                t = get_cached_siz_table_title(path).strip()
                titles.append(t if t else fallback)
            else:
                titles.append(fallback)
        return keys, titles

    def _collect_table_persons(self) -> list[EmployeeRecord]:
        sel = self.list_employees.curselection()
        if sel:
            order = sorted(int(i) for i in sel)
            out: list[EmployeeRecord] = []
            for i in order:
                if 0 <= i < len(self._employee_list_slot_gi):
                    gi = self._employee_list_slot_gi[i]
                    if gi is not None:
                        out.append(self._employee_records[gi])
            return out
        fio = self.entry_fio.get().strip()
        if fio:
            return [
                EmployeeRecord(
                    fio=fio,
                    profession=self.entry_position.get().strip(),
                    subdivision=self.entry_subdivision.get().strip(),
                )
            ]
        return []

    def _collect_table_persons_merged_by_fio(self) -> list[EmployeeRecord]:
        """
        Выбранные в списке строки с одним и тем же ФИО объединяются в одну запись
        (должности в profession/profession2, подразделение и СНИЛС — по той же логике, что таблица в .docx).
        """
        return _table_employees_dedupe_by_fio(self._collect_table_persons())

    def _on_technical_protocol_toggle(self) -> None:
        if self.var_technical_protocol.get():
            for v in self._prog_vars.values():
                v.set(False)
            for cb in self._program_checkbuttons:
                cb.state(["disabled"])
            self.lbl_tech_v_program.grid(**self._tech_v_pick_lbl_grid)
            self.cb_tech_v_program.grid(**self._tech_v_pick_cb_grid)
            self._tech_tpl_btns.grid(**self._tech_tpl_btns_grid)
            self.lbl_technical_template_main.grid(**self._tech_tpl_lbl_main_grid)
            self.list_employees.configure(height=EMPLOYEE_LIST_HEIGHT_TECH)
            self._refresh_tech_v_program_combo(silent=False)
        else:
            for cb in self._program_checkbuttons:
                cb.state(["!disabled"])
            self.lbl_tech_v_program.grid_remove()
            self.cb_tech_v_program.grid_remove()
            self._tech_tpl_btns.grid_remove()
            self.lbl_technical_template_main.grid_remove()
            self._tech_v_programs_list = []
            self.list_employees.configure(height=EMPLOYEE_LIST_HEIGHT_NORMAL)
        if hasattr(self, "lbl_template"):
            self.lbl_template.configure(text=self._template_status_text())
        self._refresh_technical_template_labels()
        self.after_idle(self._fit_main_window_to_form)

    def _refresh_tech_v_program_combo(self, *, silent: bool = False) -> None:
        if not self.var_technical_protocol.get():
            return
        prev_key = ""
        cur = self.cb_tech_v_program.current()
        if 0 <= cur < len(self._tech_v_programs_list):
            r0 = self._tech_v_programs_list[cur]
            prev_key = f"{r0.program_name.strip()}\t{r0.approver.strip()}\t{r0.approval_date_raw.strip()}"
        path = self._programs_file_resolved()
        try:
            rows = load_all_tech_v_programs_from_excel(path)
        except EmployeeExcelError as e:
            self._tech_v_programs_list = []
            self.cb_tech_v_program.configure(values=[])
            self.cb_tech_v_program.set("")
            if not silent:
                messagebox.showerror("Лист Tech_V", str(e), parent=self)
            return
        self._tech_v_programs_list = rows
        labels: list[str] = []
        for i, r in enumerate(rows):
            base = f"{i + 1}. {r.program_name.strip()}"
            extra = r.approver.strip()
            if extra:
                base = f"{base} — {extra}"
            labels.append(base)
        self.cb_tech_v_program.configure(values=labels)
        pick = 0
        if prev_key:
            for i, r in enumerate(rows):
                k = f"{r.program_name.strip()}\t{r.approver.strip()}\t{r.approval_date_raw.strip()}"
                if k == prev_key:
                    pick = i
                    break
        if labels:
            self.cb_tech_v_program.current(pick)
            self.cb_tech_v_program.set(labels[pick])

    def _get_selected_tech_v_program(self) -> TechVProgramInfo | None:
        if not self.var_technical_protocol.get() or not self._tech_v_programs_list:
            return None
        idx = self.cb_tech_v_program.current()
        if idx < 0 or idx >= len(self._tech_v_programs_list):
            return self._tech_v_programs_list[0]
        return self._tech_v_programs_list[idx]

    def _resolve_program_keys_and_tech_extra(
        self, persons_raw: list[EmployeeRecord]
    ) -> tuple[list[str], list[str], dict[str, Any]]:
        programs_path = self._programs_file_resolved()
        if self.var_technical_protocol.get():
            tinfo = self._get_selected_tech_v_program()
            if tinfo is None:
                raise EmployeeExcelError(
                    "Включён протокол по техническим вопросам: выберите программу в списке "
                    "«Программа по листу Tech_V» или обновите файл программ (F5) и проверьте лист Tech_V."
                )
            return ["TECH"], [tinfo.program_name.strip()], {
                "technical_protocol": True,
                "tech_approver": tinfo.approver,
                "tech_program_name": tinfo.program_name.strip(),
                "tech_approval_date_raw": tinfo.approval_date_raw,
            }
        keys, titles = self._collect_program_keys_and_titles(persons_raw)
        return keys, titles, {
            "technical_protocol": False,
            "tech_approver": "",
            "tech_program_name": "",
            "tech_approval_date_raw": "",
        }

    def _sync_and_refresh_commission_pool_listboxes(self) -> None:
        self._tech_commission_state.pool = self._commission_state.pool
        for pan in (self._commission_panel, self._tech_commission_panel):
            if pan is None:
                continue
            try:
                if pan.winfo_exists():
                    pan.refresh_pool_display()
            except tk.TclError:
                pass

    def generate_protocol(self) -> None:
        theme = self.entry_theme.get().strip()
        date_str = self.entry_date.get().strip()
        protocol_no = self.entry_protocol_no.get().strip()
        registry_no = self.entry_registry_no.get().strip()
        grade = self.combo_grade.get().strip()
        check_type = self.combo_check_type.get().strip() or "плановая"

        persons_raw = self._collect_table_persons_merged_by_fio()
        if not persons_raw:
            messagebox.showwarning(
                "Проверка",
                "Выберите одного или нескольких сотрудников в списке "
                "или введите ФИО вручную (без выбора в списке).",
            )
            return
        persons_b_src = self._collect_table_persons()
        if not self._configure_v_prof_combinations(persons_b_src):
            return

        programs_path = self._programs_file_resolved()
        try:
            program_keys, program_titles, tech_kw = self._resolve_program_keys_and_tech_extra(
                persons_raw
            )
        except EmployeeExcelError as e:
            messagebox.showerror("Лист Tech_V", str(e), parent=self)
            return
        tpl = self._active_protocol_template_path()
        is_docx = is_word_protocol_template(tpl)
        if is_docx and not program_titles and not self.var_technical_protocol.get():
            messagebox.showwarning(
                "Программы",
                "Для шаблона Word отметьте хотя бы одну программу обучения "
                f"(данные — листы {B_PROGRAM_SHEET_NAME} и {V_PROF_SHEET_NAME}).",
            )
            return

        if not date_str:
            date_str = date.today().strftime("%d.%m.%Y")
            self.entry_date.delete(0, tk.END)
            self.entry_date.insert(0, date_str)

        try:
            if is_docx:
                doc, table_excess = build_filled_protocol_document(
                    tpl,
                    protocol_no=protocol_no,
                    date_str=date_str,
                    theme=theme,
                    table_persons=persons_raw,
                    program_titles=program_titles,
                    program_keys=program_keys,
                    excel_path=programs_path,
                    persons_v_raw=persons_raw,
                    persons_b_row_source=persons_b_src,
                    grade=grade,
                    registry_no=registry_no,
                    check_type=check_type,
                    trained_registry_path=self._trained_registry_path_for_protocol(),
                    face_sheet_profession=self._face_sheet_profession(),

                    **self._v_prof_combo_kwargs(),
                    **tech_kw,
                )
                if table_excess > 0:
                    messagebox.showwarning(
                        "Таблица протокола",
                        self._table_fill_warning_text(),
                    )
                text = document_to_plain_text(doc)
                self._set_protocol_preview(doc, text, show_window=True)
            else:
                txt_theme = "; ".join(program_titles) if program_titles else theme
                if not txt_theme.strip():
                    messagebox.showwarning(
                        "Тема",
                        "Отметьте программы обучения или введите доп. тему для текстового шаблона.",
                    )
                    return
                text = build_protocol_text(
                    txt_theme,
                    date_str,
                    protocol_no=protocol_no,
                    template_path=self._active_protocol_template_path(),
                )
                self._set_protocol_preview(None, text, show_window=True)
        except ProtocolTemplateError as e:
            messagebox.showerror("Шаблон протокола", str(e))
            return
        except ValueError as e:
            messagebox.showerror("Шаблон протокола", str(e))
            return

        self._last_export_persons = list(persons_raw)
        self._persist_protocol_no_field()

        try:
            fio_summary = ", ".join(p.fio for p in persons_raw)
            topic_db = "; ".join(program_titles) if program_titles else theme
            pn_fmt = format_protocol_number_for_template(protocol_no, date_str)
            if not (pn_fmt or "").strip():
                pn_fmt = (protocol_no or "").strip()
            v_combo = self._v_prof_combo_kwargs()
            meta_json = build_protocol_export_meta_json(
                program_keys,
                program_titles,
                protocol_no_formatted=pn_fmt,
                persons_raw=persons_raw,
                persons_row_source=persons_b_src,
                face_sheet_profession=self._face_sheet_profession(),
                v_prof_enabled_by_fio=v_combo.get("v_prof_enabled_by_fio"),  # type: ignore[arg-type]
                v_prof_main_by_fio=v_combo.get("v_prof_main_by_fio"),  # type: ignore[arg-type]
            )
            save_protocol(
                fio_summary,
                topic_db,
                date_str,
                grade,
                "",
                export_meta_json=meta_json,
                protocol_kind=(
                    PROTOCOL_JOURNAL_KIND_TECH
                    if self.var_technical_protocol.get()
                    else PROTOCOL_JOURNAL_KIND_OT
                ),
            )
        except sqlite3.Error as e:
            messagebox.showwarning(
                "База данных",
                f"Протокол сформирован, но не удалось сохранить в базу:\n{e}",
            )

    def generate_protocol_per_employee_to_folder(self) -> None:
        theme = self.entry_theme.get().strip()
        date_str = self.entry_date.get().strip()
        registry_no = self.entry_registry_no.get().strip()
        grade = self.combo_grade.get().strip()
        check_type = self.combo_check_type.get().strip() or "плановая"

        persons_raw = self._collect_table_persons_merged_by_fio()
        if not persons_raw:
            messagebox.showwarning(
                "Проверка",
                "Выберите одного или нескольких сотрудников в списке "
                "или введите ФИО вручную (без выбора в списке).",
            )
            return
        persons_b_src = self._collect_table_persons()
        if not self._configure_v_prof_combinations(persons_b_src):
            return

        programs_path = self._programs_file_resolved()
        try:
            program_keys, program_titles, tech_kw = self._resolve_program_keys_and_tech_extra(
                persons_raw
            )
        except EmployeeExcelError as e:
            messagebox.showerror("Лист Tech_V", str(e), parent=self)
            return
        tpl = self._active_protocol_template_path()
        if not is_word_protocol_template(tpl):
            messagebox.showwarning(
                "Шаблон",
                "Отдельные файлы по сотрудникам доступны только для шаблона Word (.docx).",
            )
            return
        if not program_titles and not self.var_technical_protocol.get():
            messagebox.showwarning(
                "Программы",
                "Для шаблона Word отметьте хотя бы одну программу обучения "
                f"(данные — листы {B_PROGRAM_SHEET_NAME} и {V_PROF_SHEET_NAME}).",
            )
            return

        if not date_str:
            date_str = date.today().strftime("%d.%m.%Y")
            self.entry_date.delete(0, tk.END)
            self.entry_date.insert(0, date_str)

        base_no = protocol_sequence_start_int(self.entry_protocol_no.get())
        batch_pairs: list[tuple[str, str]] = []
        for i, emp in enumerate(persons_raw):
            n = base_no + i
            protocol_no_i = str(n)
            pn_fmt_i = format_protocol_number_for_template(protocol_no_i, date_str)
            if not (pn_fmt_i or "").strip():
                pn_fmt_i = protocol_no_i
            batch_pairs.append((emp.fio or "", pn_fmt_i))

        j_kind = (
            PROTOCOL_JOURNAL_KIND_TECH
            if self.var_technical_protocol.get()
            else PROTOCOL_JOURNAL_KIND_OT
        )
        journal_ids_replace, jerr, journal_fio_notes = journal_ids_and_error_for_per_employee_batch(
            date_str, batch_pairs, journal_kind=j_kind
        )
        if jerr:
            messagebox.showerror("Журнал протоколов", jerr, parent=self)
            return

        out_dir = filedialog.askdirectory(
            title="Папка для сохранения протоколов (по одному файлу на сотрудника)",
            initialdir=str(protocols_output_dir()),
            parent=self,
        )
        if not out_dir:
            return

        trained = self._trained_registry_path_for_protocol()
        out_path = Path(out_dir)
        existing_docx = existing_per_employee_docx_in_folder(
            out_path, persons_raw, base_no, date_str
        )

        allow_overwrite = False
        if journal_ids_replace or existing_docx:
            jn = len(journal_ids_replace)
            fn = len(existing_docx)
            parts: list[str] = []
            if jn:
                parts.append(
                    f"В журнале уже есть записи за эту дату с теми же номерами протокола ({jn} шт.) — "
                    "они будут обновлены."
                )
            if journal_fio_notes:
                lim = 6
                tail = journal_fio_notes[:lim]
                more = len(journal_fio_notes) - lim
                parts.append("Смена привязки номера к ФИО:\n" + "\n".join(tail))
                if more > 0:
                    parts.append(f"… и ещё {more}.")
            if fn:
                parts.append(f"В выбранной папке уже есть одноимённые файлы DOCX ({fn} шт.).")
            msg = "\n".join(parts) + (
                "\n\nПерезаписать файлы и обновить журнал?\n\n"
                "«Нет» — отмена (смените номер в поле «№», дату, папку или удалите записи вручную)."
            )
            if not messagebox.askyesno("Перезапись партии", msg, parent=self):
                return
            allow_overwrite = True

        n_ok = 0
        last_doc: Document | None = None
        last_no_str = ""
        last_emp = persons_raw[-1]
        table_warn = False
        raw_selection_b = self._collect_table_persons()

        for i, emp in enumerate(persons_raw):
            n = base_no + i
            protocol_no = str(n)
            last_no_str = protocol_no
            pn_fmt = format_protocol_number_for_template(protocol_no, date_str)
            if not (pn_fmt or "").strip():
                pn_fmt = protocol_no
            try:
                b_rows_src = raw_employee_rows_same_fio_as(raw_selection_b, emp)
                doc, table_excess = build_filled_protocol_document(
                    tpl,
                    protocol_no=protocol_no,
                    date_str=date_str,
                    theme=theme,
                    table_persons=[emp],
                    program_titles=program_titles,
                    program_keys=program_keys,
                    excel_path=programs_path,
                    persons_v_raw=[emp],
                    persons_b_row_source=b_rows_src,
                    grade=grade,
                    registry_no=registry_no,
                    check_type=check_type,
                    trained_registry_path=trained,
                    face_sheet_profession=self._face_sheet_profession(),

                    **self._v_prof_combo_kwargs(),
                    **tech_kw,
                )
                if table_excess > 0:
                    table_warn = True
                fname = default_protocol_save_filename(
                    protocol_no,
                    date_str,
                    ".docx",
                    person_suffix=format_fio_filename_surname_initials(emp.fio),
                )
                fpath = Path(out_dir) / fname
                if fpath.is_file() and not allow_overwrite:
                    stem = fpath.stem
                    for k in range(2, 50):
                        alt = Path(out_dir) / f"{stem}_{k}{fpath.suffix}"
                        if not alt.is_file():
                            fpath = alt
                            break
                save_formed_protocol_docx(doc, fpath)
                n_ok += 1
                last_doc = doc

                try:
                    emp_persons = raw_employee_rows_same_fio_as(raw_selection_b, emp)
                    keys_i, titles_i = self._collect_program_keys_and_titles(emp_persons)
                    topic_db = "; ".join(titles_i) if titles_i else theme
                    v_combo = self._v_prof_combo_kwargs()
                    meta_json = build_protocol_export_meta_json(
                        keys_i,
                        titles_i,
                        protocol_no_formatted=pn_fmt,
                        persons_raw=emp_persons,
                        persons_row_source=raw_selection_b,
                        face_sheet_profession=self._face_sheet_profession(),
                        v_prof_enabled_by_fio=v_combo.get("v_prof_enabled_by_fio"),  # type: ignore[arg-type]
                        v_prof_main_by_fio=v_combo.get("v_prof_main_by_fio"),  # type: ignore[arg-type]
                    )
                    save_protocol(
                        (emp.fio or "").strip(),
                        topic_db,
                        date_str,
                        grade,
                        "",
                        export_meta_json=meta_json,
                        protocol_kind=j_kind,
                    )
                except sqlite3.Error:
                    pass
            except (ProtocolTemplateError, ValueError) as e:
                messagebox.showerror(
                    "Ошибка",
                    f"Сотрудник: {emp.fio}\nНомер: {protocol_no}\n{e}",
                    parent=self,
                )
                return
            except OSError as e:
                messagebox.showerror(
                    "Ошибка",
                    f"Сотрудник: {emp.fio}\nНе удалось сохранить файл:\n{e}",
                    parent=self,
                )
                return
            except Exception as e:
                messagebox.showerror(
                    "Ошибка",
                    f"Сотрудник: {emp.fio}\n{e}",
                    parent=self,
                )
                return

        self._last_export_persons = [last_emp]
        if last_doc is not None:
            self._set_protocol_preview(
                last_doc,
                document_to_plain_text(last_doc),
                show_window=False,
            )
        else:
            self._preview_plain_text = ""
            self._last_preview_doc = None
            self.btn_save.state(["disabled"])
            self.btn_save_pdf.state(["disabled"])
            self.btn_preview.state(["disabled"])

        if table_warn:
            messagebox.showwarning(
                "Таблица протокола",
                self._table_fill_warning_text(),
                parent=self,
            )
        messagebox.showinfo(
            "Готово",
            f"Сохранено файлов DOCX: {n_ok}.\nПапка:\n{out_dir}\n\n"
            f"В файлах использованы номера с {base_no} по {last_no_str}.\n"
            f"Поле «№ протокола» в окне не менялось — следующий номер укажите вручную.",
            parent=self,
        )
        if last_doc is not None:
            self.after(100, self._show_preview_toplevel)

    def save_to_pdf(self) -> None:
        content = (self._preview_plain_text or "").rstrip()
        if not content:
            messagebox.showwarning("Сохранение", "Нет текста протокола для сохранения.")
            return

        protocol_no = self.entry_protocol_no.get().strip()
        date_str = self.entry_date.get().strip()
        path = filedialog.asksaveasfilename(
            title="Сохранить протокол в PDF",
            defaultextension=".pdf",
            initialdir=str(protocols_output_dir()),
            initialfile=self._default_protocol_initialfile(protocol_no, date_str, ".pdf"),
            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        tpl = self._active_protocol_template_path()
        is_docx_tpl = is_word_protocol_template(tpl)
        theme = self.entry_theme.get().strip()
        registry_no = self.entry_registry_no.get().strip()
        grade = self.combo_grade.get().strip()
        check_type = self.combo_check_type.get().strip() or "плановая"
        persons_raw = self._collect_table_persons_merged_by_fio()
        if not persons_raw and is_docx_tpl and self._last_export_persons:
            persons_raw = list(self._last_export_persons)
        programs_path = self._programs_file_resolved()
        tech_kw: dict[str, Any] = {
            "technical_protocol": False,
            "tech_approver": "",
            "tech_program_name": "",
            "tech_approval_date_raw": "",
        }
        program_keys: list[str] = []
        program_titles: list[str] = []
        if persons_raw:
            try:
                program_keys, program_titles, tech_kw = self._resolve_program_keys_and_tech_extra(
                    persons_raw
                )
            except EmployeeExcelError as e:
                messagebox.showerror("Лист Tech_V", str(e), parent=self)
                return
        can_word_pdf = is_docx_tpl and bool(persons_raw) and bool(program_titles)

        if is_docx_tpl and not persons_raw:
            messagebox.showwarning(
                "Сохранение",
                "Для PDF как в Word нужен список сотрудников (выберите в списке / ФИО) "
                "или снова нажмите «Сформировать протокол». Сейчас будет сохранён упрощённый PDF из текста предпросмотра.",
            )

        pdf_from_preview = False
        try:
            if can_word_pdf:
                try:
                    write_protocol_pdf_from_docx_template(
                        tpl,
                        path,
                        protocol_no=protocol_no,
                        date_str=date_str,
                        theme=theme,
                        table_persons=persons_raw,
                        program_titles=program_titles,
                        program_keys=program_keys,
                        excel_path=programs_path,
                        persons_v_raw=persons_raw,
                        persons_b_row_source=self._collect_table_persons(),
                        grade=grade,
                        registry_no=registry_no,
                        check_type=check_type,
                        trained_registry_path=self._trained_registry_path_for_protocol(),
                        face_sheet_profession=self._face_sheet_profession(),
    
                    **self._v_prof_combo_kwargs(),
                        **tech_kw,
                    )
                except Exception as e_word:
                    write_protocol_pdf(path, content)
                    pdf_from_preview = True
                    messagebox.showwarning(
                        "PDF из предпросмотра",
                        "Не удалось сохранить PDF с оформлением DOCX (нужны LibreOffice "
                        "или Microsoft Word + docx2pdf, либо сбой конвертации).\n\n"
                        f"{e_word}\n\n"
                        "Сохранён упрощённый PDF из текста предпросмотра (без таблиц и оформления Word).",
                    )
            else:
                write_protocol_pdf(path, content)
        except ValueError as e:
            messagebox.showerror("Шаблон", str(e))
            return
        except OSError as e:
            messagebox.showerror("Ошибка", f"Не удалось записать PDF-файл:\n{e}")
            return
        except RuntimeError as e:
            messagebox.showerror("Ошибка PDF", str(e))
            return
        except Exception as e:
            messagebox.showerror("Ошибка PDF", f"Не удалось сформировать PDF:\n{e}")
            return

        done_msg = f"Протокол сохранён в PDF:\n{path}"
        if pdf_from_preview:
            done_msg += "\n\n(Вариант из предпросмотра — см. предупреждение выше.)"
        messagebox.showinfo("Готово", done_msg)

    def save_to_docx(self) -> None:
        content = (self._preview_plain_text or "").rstrip()
        if not content:
            messagebox.showwarning("Сохранение", "Нет текста протокола для сохранения.")
            return

        protocol_no = self.entry_protocol_no.get().strip()
        date_str = self.entry_date.get().strip()
        path = filedialog.asksaveasfilename(
            title="Сохранить протокол в Word",
            defaultextension=".docx",
            initialdir=str(protocols_output_dir()),
            initialfile=self._default_protocol_initialfile(protocol_no, date_str, ".docx"),
            filetypes=[("Документ Word", "*.docx"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        tpl = self._active_protocol_template_path()
        is_docx_tpl = is_word_protocol_template(tpl)
        theme = self.entry_theme.get().strip()
        registry_no = self.entry_registry_no.get().strip()
        grade = self.combo_grade.get().strip()
        check_type = self.combo_check_type.get().strip() or "плановая"
        persons_raw = self._collect_table_persons_merged_by_fio()
        if not persons_raw and is_docx_tpl and self._last_export_persons:
            persons_raw = list(self._last_export_persons)
        programs_path = self._programs_file_resolved()
        tech_kw2: dict[str, Any] = {
            "technical_protocol": False,
            "tech_approver": "",
            "tech_program_name": "",
            "tech_approval_date_raw": "",
        }
        program_keys2: list[str] = []
        program_titles2: list[str] = []
        if persons_raw:
            try:
                program_keys2, program_titles2, tech_kw2 = self._resolve_program_keys_and_tech_extra(
                    persons_raw
                )
            except EmployeeExcelError as e:
                messagebox.showerror("Лист Tech_V", str(e), parent=self)
                return

        try:
            if is_docx_tpl:
                if not persons_raw:
                    messagebox.showwarning(
                        "Сохранение",
                        "Для сохранения в Word по шаблону выберите сотрудников или снова сформируйте протокол. "
                        "Сохраняю простой DOCX из текста предпросмотра.",
                    )
                    write_protocol_docx(path, content)
                    messagebox.showinfo("Готово", f"Протокол сохранён в DOCX (текст предпросмотра):\n{path}")
                    return
                if not program_titles2 and not self.var_technical_protocol.get():
                    messagebox.showwarning(
                        "Сохранение",
                        "Программы не отмечены — сохраняю простой DOCX из текста предпросмотра.",
                    )
                    write_protocol_docx(path, content)
                    messagebox.showinfo("Готово", f"Протокол сохранён в DOCX (текст предпросмотра):\n{path}")
                    return
                save_protocol_docx_from_template(
                    tpl,
                    path,
                    protocol_no=protocol_no,
                    date_str=date_str,
                    theme=theme,
                    table_persons=persons_raw,
                    program_titles=program_titles2,
                    program_keys=program_keys2,
                    excel_path=programs_path,
                    persons_v_raw=persons_raw,
                    persons_b_row_source=self._collect_table_persons(),
                    grade=grade,
                    registry_no=registry_no,
                    check_type=check_type,
                    trained_registry_path=self._trained_registry_path_for_protocol(),
                    face_sheet_profession=self._face_sheet_profession(),

                    **self._v_prof_combo_kwargs(),
                    **tech_kw2,
                )
            else:
                write_protocol_docx(path, content)
        except ValueError as e:
            messagebox.showerror("Шаблон", str(e))
            return
        except OSError as e:
            messagebox.showerror("Ошибка", f"Не удалось записать DOCX:\n{e}")
            return
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сформировать DOCX:\n{e}")
            return

        messagebox.showinfo("Готово", f"Протокол сохранён в DOCX:\n{path}")

