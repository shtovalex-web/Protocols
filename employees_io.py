# -*- coding: utf-8 -*-
"""Чтение сотрудников и кандидатов в комиссию из Excel (Data_base.xlsx)."""

from __future__ import annotations

import logging
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

_logger = logging.getLogger(__name__)


class EmployeeExcelError(Exception):
    """Ошибка чтения списка сотрудников из Excel."""


def _workbook_path_for_openpyxl(path: Path) -> Path:
    from bundle_integration import BundleOfficeConvertError, resolve_openpyxl_workbook_path

    try:
        return resolve_openpyxl_workbook_path(path)
    except BundleOfficeConvertError as e:
        raise EmployeeExcelError(str(e)) from e


@dataclass
class EmployeeRecord:
    """Строка сотрудника из листа rabotnik (ФИО, должность, подразделение, совмещаемая профессия, СНИЛС)."""

    fio: str
    profession: str = ""
    subdivision: str = ""
    profession2: str = ""
    snils: str = ""


def _norm_employee_sort_str(value: str) -> str:
    return (value or "").strip().lower().replace("ё", "е")


def sort_employees_by_fio_alphabet(records: list[EmployeeRecord]) -> None:
    """Сортирует список на месте по ФИО (лексикографически, без учёта регистра; буква ё рядом с е)."""

    records.sort(key=lambda rec: _norm_employee_sort_str(rec.fio))


def sort_employees_by_subdivision_then_fio(records: list[EmployeeRecord]) -> None:
    """Сортирует на месте: по подразделению, затем по ФИО; пустое подразделение — в конце списка."""

    def _key(rec: EmployeeRecord) -> tuple[int, str, str]:
        sub = _norm_employee_sort_str(rec.subdivision)
        fio = _norm_employee_sort_str(rec.fio)
        return (1 if not sub else 0, sub, fio)

    records.sort(key=_key)


@dataclass
class TechVProgramInfo:
    """Строка листа Tech_V: кто утвердил, наименование программы, дата (протокол по техническим вопросам)."""

    approver: str
    program_name: str
    approval_date_raw: str


EMPLOYEES_EXCEL_FILENAME = "Data_base.xlsx"
# Справочник программ (B, V_PROF, PP, SIZ, V). Если файла нет — программы читаются из Data_base.xlsx.
PROGRAMS_EXCEL_FILENAME = "Programs_base.xlsx"
PROGRAM_WORKBOOK_CANONICAL_SHEETS: tuple[str, ...] = (
    "B",
    "V_PROF",
    "PP",
    "SIZ",
    "V",
)
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

# Лист Tech_V в файле программ: кто утвердил программу, наименование, дата утверждения (протокол по техническим вопросам).
TECH_V_SHEET_NAME = "Tech_V"
TECH_V_SHEET_ALIASES: tuple[str, ...] = (
    "tech_v",
    "тех_v",
    "TECH_V",
    "Тех_V",
)


def _normalize_excel_header(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", s)


def _header_column_role(header: str) -> str | None:
    if not header:
        return None
    # Одна колонка «Фамилия, Имя, Отчество» или ФИО (актуальная форма Data_base.xlsx).
    if (
        "фио" in header
        or "ф.и.о" in header
        or ("фамилия" in header and "имя" in header)
        or any(x in header for x in ("работник", "сотрудник", "full name", "employee"))
    ):
        return "fio"
    # Раздельные колонки ФИО (если есть в файле).
    if header in ("фамилия",) or (
        header.startswith("фамилия ") and "," not in header and "имя" not in header
    ):
        return "surname"
    if header == "имя" or header.startswith("имя "):
        return "name"
    if "отчество" in header:
        return "patronymic"
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
    ) and "п/п" not in header:
        return "profession"
    if any(
        x in header
        for x in (
            "подразделение",
            "наименование подразделения",
            "цех",
            "участок",
            "отдел",
            "department",
            "unit",
            "division",
        )
    ):
        return "subdivision"
    if any(
        x in header
        for x in (
            "снилс",
            "страховой номер",
            "индивидуальный номер",
            "snils",
            "страхов",
            "свидетельств",
        )
    ):
        return "snils"
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


_FIO_HEADER_PLACEHOLDERS = frozenset(
    {
        "фио",
        "фамилия, имя, отчество",
        "фамилия имя отчество",
        "фамилия",
        "имя",
        "отчество",
    }
)


def _is_employee_data_fio(fio: str) -> bool:
    n = _normalize_excel_header(fio)
    return bool(n) and n not in _FIO_HEADER_PLACEHOLDERS


def _find_employee_header_row_index(rows: list[tuple[Any, ...] | None]) -> int:
    """Первая строка с распознанными заголовками сотрудников (до 25 строк листа)."""
    for i, row in enumerate(rows[:25]):
        if not row:
            continue
        cols = _detect_employee_columns(tuple(row))
        if "fio" in cols or ("surname" in cols and "name" in cols):
            return i
    return 0


def _employee_fio_from_row(row: tuple[Any, ...], cols: dict[str, int]) -> str:
    if "fio" in cols:
        return _excel_cell_str(row, cols["fio"])
    parts: list[str] = []
    for key in ("surname", "name", "patronymic"):
        if key in cols:
            part = _excel_cell_str(row, cols[key])
            if part:
                parts.append(part)
    return " ".join(parts)


def _detect_commission_first_data_row(ws: Any) -> int:
    """
    Строка с подписями «ФИО» в колонках A и D → данные со следующей строки.
    Иначе — COMMISSION_FIRST_DATA_ROW (как в старых шаблонах).
    """
    max_r = min(int(ws.max_row or 0), COMMISSION_FIRST_DATA_ROW + 12)
    for r in range(1, max_r + 1):
        row = next(
            ws.iter_rows(
                min_row=r,
                max_row=r,
                min_col=1,
                max_col=COMMISSION_MAX_COL,
                values_only=True,
            ),
            None,
        )
        tup = tuple(row) if row is not None else ()
        a = _normalize_excel_header(_row_value_str(tup, COMMISSION_COL_CHAIR_FIO))
        d = _normalize_excel_header(_row_value_str(tup, COMMISSION_COL_MEMBER_FIO))
        if a == "фио" and d == "фио":
            return r + 1
    return COMMISSION_FIRST_DATA_ROW


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

    path = _workbook_path_for_openpyxl(path)

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise EmployeeExcelError(
            f"Не удалось открыть файл (нужен формат .xlsx или .xlsm, не старый .xls):\n{path}\n\n"
            f"{type(e).__name__}: {e}"
        ) from e
    try:
        ws = _pick_employee_worksheet(wb, sheet_name)
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            return []
        header_idx = _find_employee_header_row_index(all_rows)
        header = tuple(all_rows[header_idx])
        cols = _detect_employee_columns(header)
        has_fio_cols = "fio" in cols or ("surname" in cols and "name" in cols)
        if not has_fio_cols:
            col_fio, col_prof, col_sub = 0, 1, 2
            col_prof2 = -1
            col_snils = -1
            use_legacy_cols = True
        else:
            col_prof = cols.get("profession", -1)
            col_sub = cols.get("subdivision", -1)
            col_prof2 = cols.get("profession2", -1)
            col_snils = cols.get("snils", -1)
            use_legacy_cols = False

        out: list[EmployeeRecord] = []
        for row in all_rows[header_idx + 1 :]:
            if not row:
                continue
            tup = tuple(row)
            if use_legacy_cols:
                fio = _excel_cell_str(tup, col_fio)
            else:
                fio = _employee_fio_from_row(tup, cols)
            if not _is_employee_data_fio(fio):
                continue
            p2 = _excel_cell_str(tup, col_prof2) if col_prof2 >= 0 else ""
            sn = _excel_cell_str(tup, col_snils) if col_snils >= 0 else ""
            prof = (
                _excel_cell_str(tup, col_prof)
                if col_prof >= 0
                else (_excel_cell_str(tup, 1) if use_legacy_cols else "")
            )
            sub = (
                _excel_cell_str(tup, col_sub)
                if col_sub >= 0
                else (_excel_cell_str(tup, 2) if use_legacy_cols else "")
            )
            out.append(
                EmployeeRecord(
                    fio=fio,
                    profession=prof,
                    subdivision=sub,
                    profession2=p2,
                    snils=sn,
                )
            )
        return out
    except EmployeeExcelError:
        raise
    except Exception as e:
        _logger.exception("Ошибка чтения листа сотрудников: %s", path)
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

    path = _workbook_path_for_openpyxl(path)

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        _logger.exception("Не удалось открыть Excel (комиссия): %s", path)
        raise EmployeeExcelError(
            f"Не удалось открыть файл (нужен формат .xlsx или .xlsm):\n{path}\n\n"
            f"{type(e).__name__}: {e}"
        ) from e
    try:
        ws = _pick_commission_worksheet(wb)
        first_row = _detect_commission_first_data_row(ws)
        max_r = ws.max_row or 0
        last = min(max_r, first_row + COMMISSION_MAX_SCAN_ROWS - 1)
        if last < first_row:
            return []

        seen_keys: set[str] = set()
        out: list[EmployeeRecord] = []

        def _try_add(fio_col: int, pos_col: int, tup: tuple[Any, ...]) -> None:
            fio = _row_value_str(tup, fio_col)
            if not fio or _normalize_excel_header(fio) == "фио":
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
            min_row=first_row,
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
        _logger.exception("Ошибка чтения листа комиссии: %s", path)
        raise EmployeeExcelError(
            f"Ошибка при чтении листа комиссии в файле:\n{path}\n\n{type(e).__name__}: {e}"
        ) from e
    finally:
        wb.close()


def _pick_tech_v_worksheet(wb: Any) -> Any:
    try:
        raw_names = wb.sheetnames
    except AttributeError as e:
        raise EmployeeExcelError("Не удалось прочитать список листов книги Excel.") from e
    names_lower: dict[str, str] = {}
    for n in raw_names:
        names_lower[n.lower().strip()] = n
    hints: list[str] = []
    for h in (TECH_V_SHEET_NAME,) + TECH_V_SHEET_ALIASES:
        t = h.strip()
        if t and t.lower() not in [x.lower() for x in hints]:
            hints.append(t)
    for h in hints:
        real = names_lower.get(h.lower().strip())
        if real is not None:
            return wb[real]
    avail = ", ".join(raw_names) if raw_names else "(листов нет)"
    raise EmployeeExcelError(
        f"Не найден лист Tech_V (программы по техническим вопросам). "
        f"Добавьте лист «{TECH_V_SHEET_NAME}» в файл программ.\n"
        f"Доступные листы: {avail}"
    )


def _tech_v_column_map(header_row: tuple[Any, ...]) -> dict[str, int] | None:
    """По строке заголовков: столбцы approver / program / date; None — фиксированные A,B,C."""
    roles: dict[str, int] = {}
    for j, cell in enumerate(header_row):
        hn = _normalize_excel_header(cell)
        if not hn:
            continue
        if "утверд" in hn:
            roles["approver"] = j
        elif "програм" in hn or "наимен" in hn or "тем" in hn or "назван" in hn:
            roles["program"] = j
        elif hn.startswith("дата") or " дата" in f" {hn} ":
            roles["date"] = j
    if "program" in roles and ("approver" in roles or "date" in roles):
        return roles
    if "program" in roles:
        return roles
    return None


def load_all_tech_v_programs_from_excel(path: Path) -> list[TechVProgramInfo]:
    """
    Все строки листа Tech_V с непустым наименованием программы (сверху вниз).
    Заголовки и столбцы — как у load_tech_v_program_from_excel.
    """
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните в папке проекта:\n  pip install openpyxl"
        ) from e

    if not path.is_file():
        raise EmployeeExcelError(f"Файл программ не найден:\n{path}")

    path = _workbook_path_for_openpyxl(path)

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise EmployeeExcelError(
            f"Не удалось открыть файл программ:\n{path}\n\n{type(e).__name__}: {e}"
        ) from e
    try:
        ws = _pick_tech_v_worksheet(wb)
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            raise EmployeeExcelError(f"Лист «{TECH_V_SHEET_NAME}» пуст в файле:\n{path}")

        cmap = _tech_v_column_map(rows[0])
        start = 1 if cmap else 0
        if not cmap:
            cmap = {"approver": 0, "program": 1, "date": 2}

        def cell(row: tuple[Any, ...], role: str) -> str:
            j = cmap.get(role, -1)
            if j < 0:
                return ""
            return _excel_cell_str(row, j)

        out: list[TechVProgramInfo] = []
        for i in range(start, len(rows)):
            tup = rows[i]
            prog = cell(tup, "program").strip()
            if not prog:
                continue
            out.append(
                TechVProgramInfo(
                    approver=cell(tup, "approver").strip(),
                    program_name=prog,
                    approval_date_raw=cell(tup, "date").strip(),
                )
            )

        if not out:
            raise EmployeeExcelError(
                f"На листе «{TECH_V_SHEET_NAME}» нет строки с наименованием программы (столбец программы пуст).\n{path}"
            )
        return out
    except EmployeeExcelError:
        raise
    except Exception as e:
        _logger.exception("Ошибка чтения Tech_V: %s", path)
        raise EmployeeExcelError(
            f"Ошибка при чтении листа Tech_V:\n{path}\n\n{type(e).__name__}: {e}"
        ) from e
    finally:
        wb.close()


def load_tech_v_program_from_excel(path: Path) -> TechVProgramInfo:
    """
    Первая строка с программой на листе Tech_V (для обратной совместимости).
    """
    rows = load_all_tech_v_programs_from_excel(path)
    return rows[0]


def format_fio_filename_surname_initials(fio: str) -> str:
    """
    Краткая подпись для имени файла: «Фамилия И.О.» (например, Иванов И.П.).
    Вход: «Фамилия Имя Отчество», «Фамилия И. О.», одно слово — как есть.
    """
    fio = (fio or "").strip()
    if not fio:
        return ""
    parts = fio.split()
    if len(parts) == 1:
        return parts[0]

    m = re.match(
        r"^(.+?)\s+([А-ЯЁA-Z])\.([А-ЯЁA-Z])\.(\s*)$",
        fio,
        re.IGNORECASE,
    )
    if m:
        return f"{m.group(1).strip()} {m.group(2).upper()}.{m.group(3).upper()}."

    m = re.match(
        r"^(.+?)\s+([А-ЯЁA-Z])\.(\s+)([А-ЯЁA-Z])\.(\s*)$",
        fio,
        re.IGNORECASE,
    )
    if m:
        return f"{m.group(1).strip()} {m.group(2).upper()}.{m.group(4).upper()}."

    m = re.match(r"^(.+?)\s+([А-ЯЁA-Z])\.(\s*)$", fio, re.IGNORECASE)
    if m:
        return f"{m.group(1).strip()} {m.group(2).upper()}."

    fam = parts[0]
    if len(parts) >= 3 and all("." not in p for p in parts):
        name, pat = parts[1], parts[2]
        if name and pat:
            return f"{fam} {name[0].upper()}.{pat[0].upper()}."
    if len(parts) == 2 and "." not in parts[1]:
        name = parts[1]
        if name:
            return f"{fam} {name[0].upper()}."
    return fam


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


def listbox_label_for_employee(
    rec: EmployeeRecord, *, grouped_by_subdivision: bool = False
) -> str:
    extra = ""
    if rec.profession2:
        extra = f" + {rec.profession2}"
    if grouped_by_subdivision:
        prof = (rec.profession or "").strip() or "—"
        return f"  {rec.fio} — {prof}{extra}"
    if rec.profession:
        return f"{rec.fio} — {rec.profession}{extra}"
    return rec.fio


def subdivision_group_key(subdivision: str) -> str:
    """Ключ группы подразделения для сворачивания списка в интерфейсе."""
    s = (subdivision or "").strip().lower().replace("ё", "е")
    return s if s else "__no_sub__"


def listbox_subdivision_header(
    subdivision: str,
    employee_count: int,
    *,
    collapsed: bool = False,
) -> str:
    """Заголовок группы в списке сотрудников (клик — свернуть/развернуть)."""
    sub = (subdivision or "").strip() or "(без подразделения)"
    n = max(0, int(employee_count))
    mark = "▸" if collapsed else "▾"
    hint = " — свернуто, щёлкните чтобы развернуть" if collapsed else ""
    return f"{mark} {sub}  ({n}){hint}"


def employee_unique_key(rec: EmployeeRecord) -> str:
    return "|".join(
        (
            (rec.fio or "").strip().lower(),
            (rec.profession or "").strip().lower(),
            (rec.subdivision or "").strip().lower(),
            (rec.profession2 or "").strip().lower(),
        )
    )


def write_template_data_base_workbook(path: Path) -> None:
    """
    Пустая книга Data_base.xlsx: лист сотрудников (заголовки) и лист комиссии (пояснение + данные с 3-й строки).
    """
    try:
        from openpyxl import Workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните: pip install openpyxl"
        ) from e

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = EMPLOYEES_SHEET_NAME
    ws.append(
        [
            "№ п/п",
            "Таб.№",
            "Фамилия, Имя, Отчество",
            "Подразделение",
            "Должность",
            "№ страхового свидетельства",
        ],
    )
    ws2 = wb.create_sheet(COMMISSION_SHEET_NAME)
    ws2["A1"] = "Председатель"
    ws2["D1"] = "Члены  комиссии"
    ws2.append(["ФИО", "Должность", None, "ФИО", "Должность"])
    wb.save(path)


def _workbook_sheet_lookup(wb: Any) -> dict[str, str]:
    return {n.lower(): n for n in wb.sheetnames}


def _source_sheet_name_for_program(wb: Any, canonical: str) -> str | None:
    m = _workbook_sheet_lookup(wb)
    if canonical.lower() in m:
        return m[canonical.lower()]
    if canonical == "V":
        for alias in ("v", "в"):
            if alias in m:
                return m[alias]
    return None


def copy_program_sheets_from_workbook(source: Path, dest: Path) -> list[str]:
    """
    Копирует листы программ из объединённого Data_base (или другого файла) в отдельную книгу Programs_base.
    Возвращает список скопированных канонических имён листов.
    """
    try:
        from openpyxl import Workbook, load_workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните: pip install openpyxl"
        ) from e

    source = Path(source)
    dest = Path(dest)
    if not source.is_file():
        raise EmployeeExcelError(f"Файл не найден:\n{source}")

    source = _workbook_path_for_openpyxl(source)

    swb = load_workbook(source, data_only=True, read_only=True)
    try:
        dwb = Workbook()
        if dwb.active is not None:
            dwb.remove(dwb.active)
        copied: list[str] = []
        for canon in PROGRAM_WORKBOOK_CANONICAL_SHEETS:
            src_name = _source_sheet_name_for_program(swb, canon)
            if not src_name:
                continue
            ws_src = swb[src_name]
            ws_dst = dwb.create_sheet(canon)
            for row in ws_src.iter_rows(values_only=True):
                ws_dst.append(list(row) if row is not None else [])
            copied.append(canon)
    finally:
        swb.close()

    if not copied:
        raise EmployeeExcelError(
            "В выбранном файле нет листов программ (ожидаются: B, V_PROF, PP, SIZ, V)."
        )

    dest.parent.mkdir(parents=True, exist_ok=True)
    dwb.save(dest)
    return copied


def split_combined_employees_workbook(
    employees_path: Path,
    programs_path: Path | None = None,
    *,
    backup_employees: bool = True,
) -> tuple[list[str], Path]:
    """
    Вынести листы программ из объединённого Data_base в Programs_base.xlsx и удалить их из книги сотрудников.

    Сначала (при backup_employees) сохраняется копия «имя_before_split.xlsx» с полным содержимым до разбиения.
    Значения на листах программ копируются как в copy_program_sheets_from_workbook (без формул).
    """
    employees_path = Path(employees_path).expanduser().resolve()
    programs_path = (
        Path(programs_path).expanduser().resolve()
        if programs_path is not None
        else employees_path.parent / PROGRAMS_EXCEL_FILENAME
    )
    if not employees_path.is_file():
        raise EmployeeExcelError(f"Файл не найден:\n{employees_path}")

    if backup_employees:
        bak = employees_path.with_name(f"{employees_path.stem}_before_split{employees_path.suffix}")
        shutil.copy2(employees_path, bak)

    copied = copy_program_sheets_from_workbook(employees_path, programs_path)

    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните: pip install openpyxl"
        ) from e

    wb = load_workbook(employees_path, data_only=False)
    to_remove: list[str] = []
    for canon in PROGRAM_WORKBOOK_CANONICAL_SHEETS:
        sn = _source_sheet_name_for_program(wb, canon)
        if sn and sn not in to_remove:
            to_remove.append(sn)
    for sn in to_remove:
        wb.remove(wb[sn])
    if not wb.sheetnames:
        raise EmployeeExcelError(
            "После удаления листов программ в книге не осталось листов — файл не сохранён (восстановите из копии _before_split)."
        )
    wb.save(employees_path)
    return copied, programs_path


def write_template_programs_workbook(path: Path) -> None:
    """Пустая книга Programs_base.xlsx: листы B, V_PROF, PP, SIZ, V с поясняющими заголовками."""
    try:
        from openpyxl import Workbook
    except ImportError as e:
        raise EmployeeExcelError(
            "Не установлен пакет openpyxl. Выполните: pip install openpyxl"
        ) from e

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "B"
    ws.append(
        [
            "(опц.) примечание",
            "Наименование программы «Б» (столбец 2)",
            "Объём, ч (столбец 3, опц.)",
        ]
    )
    vp = wb.create_sheet("V_PROF")
    hdr = ["Должность (A)", "Якорь «Б» (2)", "Якорь ПП (3)", "Якорь СИЗ (4)"]
    hdr.extend([f"Фрагм. «В» ({i})" for i in range(5, 23)])
    vp.append(hdr)
    pp = wb.create_sheet("PP")
    pp.append(["…", "Наименование ПП для таблицы (столбец 2)", "Объём, ч (ст. 3, опц.)"])
    siz = wb.create_sheet("SIZ")
    siz.append(["…", "Наименование СИЗ (столбец 2)", "Объём, ч (ст. 3, опц.)"])
    v = wb.create_sheet("V")
    v.append(
        [
            "ID в гос. реестре (A)",
            "Сопоставление с V_PROF (B)",
            "Наименование после проверки (C)",
            "Объём, ч (столбец D, опц.)",
        ]
    )
    wb.save(path)
