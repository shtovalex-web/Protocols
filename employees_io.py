# -*- coding: utf-8 -*-
"""Чтение сотрудников и кандидатов в комиссию из Excel (Data_base.xlsx)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class EmployeeExcelError(Exception):
    """Ошибка чтения списка сотрудников из Excel."""


@dataclass
class EmployeeRecord:
    """Строка сотрудника из листа rabotnik (ФИО, должность, подразделение, совмещаемая профессия)."""

    fio: str
    profession: str = ""
    subdivision: str = ""
    profession2: str = ""


EMPLOYEES_EXCEL_FILENAME = "Data_base.xlsx"
EMPLOYEES_SHEET_NAME = "rabotnik"
EMPLOYEES_SHEET_ALIASES: tuple[str, ...] = (
    "rabotnik",
    "работник",
    "работники",
    "сотрудники",
    "сотрудник",
    "список сотрудников",
    "кадры",
)

COMMISSION_SHEET_NAME = "komission"
COMMISSION_SHEET_ALIASES: tuple[str, ...] = (
    "komission",
    "комиссия",
    "комиссия по проверке",
    "commission",
)
# Лист komission: данные с 3-й строки Excel. A — ФИО председателя, B — должность справа от A;
# D — ФИО члена, E — должность справа от D.
COMMISSION_FIRST_DATA_ROW = 3
COMMISSION_COL_CHAIR_FIO = 1  # A
COMMISSION_COL_CHAIR_POSITION = 2  # B
COMMISSION_COL_MEMBER_FIO = 4  # D
COMMISSION_COL_MEMBER_POSITION = 5  # E
COMMISSION_MAX_COL = max(COMMISSION_COL_CHAIR_POSITION, COMMISSION_COL_MEMBER_POSITION)
COMMISSION_MAX_SCAN_ROWS = 500


def _normalize_excel_header(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", s)


def _header_column_role(header: str) -> str | None:
    if not header:
        return None
    if any(
        x in header
        for x in (
            "фио",
            "ф.и.о",
            "работник",
            "сотрудник",
            "фамилия",
            "name",
            "full name",
            "employee",
        )
    ):
        return "fio"
    if any(
        x in header
        for x in (
            "профессия",
            "должность",
            "специальность",
            "position",
            "job title",
            "title",
            "должн",
        )
    ):
        return "profession"
    if any(
        x in header
        for x in (
            "подразделение",
            "цех",
            "участок",
            "отдел",
            "department",
            "unit",
            "division",
        )
    ):
        return "subdivision"
    return None


def _detect_employee_columns(header_row: tuple[Any, ...]) -> dict[str, int]:
    roles: dict[str, int] = {}
    prof_cols: list[int] = []
    for j, cell in enumerate(header_row):
        hn = _normalize_excel_header(cell)
        if ("совмещ" in hn or "вторая" in hn) and (
            "проф" in hn or "должн" in hn or "спец" in hn
        ):
            if "profession2" not in roles:
                roles["profession2"] = j
            continue
        role = _header_column_role(hn)
        if role == "profession":
            prof_cols.append(j)
        elif role and role not in roles:
            roles[role] = j
    if prof_cols:
        roles["profession"] = prof_cols[0]
        if len(prof_cols) > 1 and "profession2" not in roles:
            roles["profession2"] = prof_cols[1]
    return roles


def _excel_cell_str(row: tuple[Any, ...], index: int) -> str:
    if index >= len(row) or row[index] is None:
        return ""
    return str(row[index]).strip()


def _pick_employee_worksheet(wb: Any, preferred_sheet: str) -> Any:
    """Лист сотрудников: сначала точное имя (без регистра), затем алиасы."""
    try:
        raw_names = wb.sheetnames
    except AttributeError as e:
        raise EmployeeExcelError("Не удалось прочитать список листов книги Excel.") from e
    names_lower: dict[str, str] = {}
    for n in raw_names:
        names_lower[n.lower().strip()] = n
    hints: list[str] = []
    for h in (preferred_sheet,) + EMPLOYEES_SHEET_ALIASES:
        t = h.strip()
        if t and t.lower() not in [x.lower() for x in hints]:
            hints.append(t)
    for h in hints:
        real = names_lower.get(h.lower().strip())
        if real is not None:
            return wb[real]
    avail = ", ".join(raw_names) if raw_names else "(листов нет)"
    raise EmployeeExcelError(
        f"Не найден лист сотрудников. Проверьте имя листа (ожидается «{EMPLOYEES_SHEET_NAME}» или похожее).\n"
        f"Доступные листы: {avail}"
    )


def load_employees_from_excel(path: Path, *, sheet_name: str = EMPLOYEES_SHEET_NAME) -> list[EmployeeRecord]:
    """Читает сотрудников с листа Excel; первая строка — заголовки с ФИО / должностью / подразделением."""
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните в папке проекта:\n"
            "  pip install openpyxl"
        ) from e

    if not path.is_file():
        raise EmployeeExcelError(f"Файл не найден:\n{path}")

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise EmployeeExcelError(
            f"Не удалось открыть файл (нужен формат .xlsx или .xlsm, не старый .xls):\n{path}\n\n"
            f"{type(e).__name__}: {e}"
        ) from e
    try:
        ws = _pick_employee_worksheet(wb, sheet_name)
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            return []
        cols = _detect_employee_columns(tuple(header))
        if "fio" not in cols:
            col_fio, col_prof, col_sub = 0, 1, 2
            col_prof2 = -1
        else:
            col_fio = cols["fio"]
            col_prof = cols.get("profession", col_fio + 1)
            col_sub = cols.get("subdivision", col_fio + 2)
            col_prof2 = cols.get("profession2", -1)

        out: list[EmployeeRecord] = []
        for row in rows:
            if not row:
                continue
            tup = tuple(row)
            fio = _excel_cell_str(tup, col_fio)
            if not fio:
                continue
            p2 = _excel_cell_str(tup, col_prof2) if col_prof2 >= 0 else ""
            out.append(
                EmployeeRecord(
                    fio=fio,
                    profession=_excel_cell_str(tup, col_prof),
                    subdivision=_excel_cell_str(tup, col_sub),
                    profession2=p2,
                )
            )
        return out
    except EmployeeExcelError:
        raise
    except Exception as e:
        raise EmployeeExcelError(
            f"Ошибка при чтении листа сотрудников в файле:\n{path}\n\n{type(e).__name__}: {e}"
        ) from e
    finally:
        wb.close()


def _pick_commission_worksheet(wb: Any) -> Any:
    """Лист комиссии: сначала «komission», затем алиасы (без учёта регистра)."""
    try:
        raw_names = wb.sheetnames
    except AttributeError as e:
        raise EmployeeExcelError("Не удалось прочитать список листов книги Excel.") from e
    names_lower: dict[str, str] = {}
    for n in raw_names:
        names_lower[n.lower().strip()] = n
    hints: list[str] = []
    for h in (COMMISSION_SHEET_NAME,) + COMMISSION_SHEET_ALIASES:
        t = h.strip()
        if t and t.lower() not in [x.lower() for x in hints]:
            hints.append(t)
    for h in hints:
        real = names_lower.get(h.lower().strip())
        if real is not None:
            return wb[real]
    avail = ", ".join(raw_names) if raw_names else "(листов нет)"
    raise EmployeeExcelError(
        f"Не найден лист комиссии. Ожидается «{COMMISSION_SHEET_NAME}» или похожее имя.\n"
        f"Доступные листы: {avail}"
    )


def _row_value_str(row: tuple[Any, ...], col_one_based: int) -> str:
    idx = col_one_based - 1
    if idx < 0 or idx >= len(row) or row[idx] is None:
        return ""
    return str(row[idx]).replace("\r\n", " ").replace("\n", " ").strip()


def load_commission_from_excel(path: Path) -> list[EmployeeRecord]:
    """
    Читает кандидатов в комиссию с листа komission.
    Со строки COMMISSION_FIRST_DATA_ROW: A+должность в B, D+должность в E;
    в списке — уникальные пары ФИО+должность (по строкам: сначала блок A/B, затем D/E).
    """
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните в папке проекта:\n"
            "  pip install openpyxl"
        ) from e

    if not path.is_file():
        raise EmployeeExcelError(f"Файл не найден:\n{path}")

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise EmployeeExcelError(
            f"Не удалось открыть файл (нужен формат .xlsx или .xlsm):\n{path}\n\n"
            f"{type(e).__name__}: {e}"
        ) from e
    try:
        ws = _pick_commission_worksheet(wb)
        max_r = ws.max_row or 0
        last = min(max_r, COMMISSION_FIRST_DATA_ROW + COMMISSION_MAX_SCAN_ROWS - 1)
        if last < COMMISSION_FIRST_DATA_ROW:
            return []

        seen_keys: set[str] = set()
        out: list[EmployeeRecord] = []

        def _try_add(fio_col: int, pos_col: int, tup: tuple[Any, ...]) -> None:
            fio = _row_value_str(tup, fio_col)
            if not fio:
                return
            rec = EmployeeRecord(
                fio=fio,
                profession=_row_value_str(tup, pos_col),
            )
            key = employee_unique_key(rec)
            if key in seen_keys:
                return
            seen_keys.add(key)
            out.append(rec)

        for row in ws.iter_rows(
            min_row=COMMISSION_FIRST_DATA_ROW,
            max_row=last,
            min_col=1,
            max_col=COMMISSION_MAX_COL,
            values_only=True,
        ):
            tup = tuple(row) if row is not None else ()
            _try_add(COMMISSION_COL_CHAIR_FIO, COMMISSION_COL_CHAIR_POSITION, tup)
            _try_add(COMMISSION_COL_MEMBER_FIO, COMMISSION_COL_MEMBER_POSITION, tup)
        return out
    except EmployeeExcelError:
        raise
    except Exception as e:
        raise EmployeeExcelError(
            f"Ошибка при чтении листа комиссии в файле:\n{path}\n\n{type(e).__name__}: {e}"
        ) from e
    finally:
        wb.close()


def format_fio_iof(fio: str) -> str:
    """
    ФИО для подписи: «И.О. Фамилия» (инициалы — имя и отчество, затем фамилия).
    Допускает вход «Фамилия Имя Отчество», «Фамилия И.О.», «Фамилия И. О.».
    """
    fio = (fio or "").strip()
    if not fio:
        return ""
    parts = fio.split()
    if len(parts) == 1:
        return parts[0]

    # Фамилия И.О. (без пробела между инициалами)
    m = re.match(
        r"^(.+?)\s+([А-ЯЁA-Z])\.([А-ЯЁA-Z])\.(\s*)$",
        fio,
        re.IGNORECASE,
    )
    if m:
        fam = m.group(1).strip()
        return f"{m.group(2).upper()}.{m.group(3).upper()}. {fam}"

    # Фамилия И. О.
    m = re.match(
        r"^(.+?)\s+([А-ЯЁA-Z])\.(\s+)([А-ЯЁA-Z])\.(\s*)$",
        fio,
        re.IGNORECASE,
    )
    if m:
        fam = m.group(1).strip()
        return f"{m.group(2).upper()}.{m.group(4).upper()}. {fam}"

    # Фамилия И.
    m = re.match(r"^(.+?)\s+([А-ЯЁA-Z])\.(\s*)$", fio, re.IGNORECASE)
    if m:
        fam = m.group(1).strip()
        return f"{m.group(2).upper()}. {fam}"

    # Фамилия Имя Отчество (без точек в частях)
    if len(parts) == 3 and all("." not in p for p in parts):
        fam, name, pat = parts
        if name and pat:
            return f"{name[0].upper()}.{pat[0].upper()}. {fam}"

    if len(parts) == 2 and "." not in parts[1]:
        fam, name = parts
        if name:
            return f"{name[0].upper()}. {fam}"

    return fio


def format_person_iof_line(fio: str, profession: str = "") -> str:
    """Строка для блока подписей: И.О. Фамилия, должность (именительный падеж как в базе)."""
    s = format_fio_iof(fio)
    p = (profession or "").strip()
    if s and p:
        return f"{s}, {p}"
    return s or p


def listbox_label_for_employee(rec: EmployeeRecord) -> str:
    extra = ""
    if rec.profession2:
        extra = f" + {rec.profession2}"
    if rec.profession:
        return f"{rec.fio} — {rec.profession}{extra}"
    return rec.fio


def employee_unique_key(rec: EmployeeRecord) -> str:
    return "|".join(
        (
            (rec.fio or "").strip().lower(),
            (rec.profession or "").strip().lower(),
            (rec.subdivision or "").strip().lower(),
            (rec.profession2 or "").strip().lower(),
        )
    )
