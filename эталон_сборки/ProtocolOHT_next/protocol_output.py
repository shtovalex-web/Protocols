# -*- coding: utf-8 -*-
"""Текстовый протокол (по строкам шаблона), простой PDF и PDF через Word из собранного DOCX."""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from typing import Any

from docx import Document

from docx_template_protection import save_formed_protocol_docx
from employees_io import EmployeeRecord
from protocol_docx import _fill_protocol_form, build_filled_protocol_document, load_protocol_form_lines

_FPDF_TYPES: tuple[type[Any], Any, Any] | None = None


def _ensure_fpdf() -> tuple[type[Any], Any, Any]:
    """fpdf2 при import тянет sign → unittest.mock → asyncio → _overlapped; откладываем до PDF."""
    global _FPDF_TYPES
    if _FPDF_TYPES is None:
        from fpdf import FPDF

        try:
            from fpdf.enums import XPos, YPos
        except ImportError:
            XPos = None
            YPos = None
        _FPDF_TYPES = (FPDF, XPos, YPos)
    return _FPDF_TYPES


def build_protocol_text(
    theme: str,
    date_str: str,
    protocol_no: str = "",
    template_path: Path | None = None,
) -> str:
    form = load_protocol_form_lines(template_path)
    filled = _fill_protocol_form(
        form,
        protocol_no=protocol_no,
        date_str=date_str,
        theme=theme,
    )
    return "\n".join(filled)


def _fpdf_try_add_ttf(pdf: Any, family: str, ttf_path: Path) -> None:
    """Совместимость PyFPDF 1.x (uni=True) и fpdf2 (аргумент uni удалён)."""
    p = str(ttf_path.resolve())
    try:
        pdf.add_font(family, "", p, uni=True)
    except TypeError:
        pdf.add_font(family, "", p)


def _fpdf_output_file(pdf: Any, path: str) -> None:
    try:
        pdf.output(path, "F")
    except TypeError:
        pdf.output(path)


def _windows_cyrillic_ttf_candidates() -> list[Path]:
    fonts = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    names = ("arial.ttf", "calibri.ttf", "segoeui.ttf", "arialuni.ttf")
    return [fonts / n for n in names if (fonts / n).is_file()]


def _configure_fpdf_font(pdf: Any, content: str) -> None:
    """Latin-1 — Helvetica; кириллица — первый подходящий TTF из папки Fonts Windows."""
    try:
        content.encode("latin-1")
        pdf.set_font("Helvetica", "", 12)
    except UnicodeEncodeError:
        last_err: Exception | None = None
        for i, ttf in enumerate(_windows_cyrillic_ttf_candidates()):
            fam = f"AppCyr{i}"
            try:
                _fpdf_try_add_ttf(pdf, fam, ttf)
                pdf.set_font(fam, "", 12)
                return
            except Exception as e:
                last_err = e
                continue
        msg = (
            "Для кириллицы в PDF нужен файл шрифта (.ttf) в папке Fonts Windows "
            "(например arial.ttf или calibri.ttf). Стандартные шрифты PDF кириллицу не рисуют."
        )
        if last_err is not None:
            raise RuntimeError(f"{msg}\nПодробнее: {last_err}") from last_err
        raise RuntimeError(msg) from None


def write_protocol_pdf(path: str, content: str) -> None:
    FPDF, XPos, YPos = _ensure_fpdf()
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    _configure_fpdf_font(pdf, content)
    normalized = content.replace("\r\n", "\n").replace("\r", "\n")
    w = getattr(pdf, "epw", pdf.w - pdf.l_margin - pdf.r_margin)
    w = max(w, 20.0)
    if XPos is not None and YPos is not None:
        for line in normalized.split("\n"):
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(
                w,
                7,
                line,
                align="L",
                new_x=XPos.LMARGIN,
                new_y=YPos.NEXT,
            )
    else:
        for line in normalized.split("\n"):
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(w, 7, line)
    _fpdf_output_file(pdf, path)


def _docx_to_pdf_via_word_com(docx_path: Path, pdf_path: Path) -> None:
    """
    Конвертация DOCX → PDF через Microsoft Word (COM).
    Не использует docx2pdf: там tqdm пишет в stderr, а в tkinter / pythonw / PyInstaller --windowed
    часто sys.stderr is None → 'NoneType' object has no attribute 'write'.
    """
    try:
        import win32com.client  # type: ignore[import-untyped]
    except ImportError as e:
        raise RuntimeError(
            "Для PDF через Word нужны Microsoft Word и пакет pywin32:\n"
            "py -3 -m pip install pywin32"
        ) from e

    docx_path = docx_path.resolve()
    pdf_path = pdf_path.resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    wd_format_pdf = 17
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        word.DisplayAlerts = 0
    except Exception:
        pass
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        doc.SaveAs(str(pdf_path), FileFormat=wd_format_pdf)
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def write_protocol_pdf_from_docx_template(
    template_path: Path,
    pdf_path: str,
    *,
    protocol_no: str,
    date_str: str,
    theme: str,
    table_persons: list[EmployeeRecord],
    program_titles: list[str] | None = None,
    program_keys: list[str] | None = None,
    excel_path: Path | None = None,
    persons_v_raw: list[EmployeeRecord] | None = None,
    persons_b_row_source: list[EmployeeRecord] | None = None,
    grade: str = "",
    registry_no: str = "",
    check_type: str = "плановая",
    trained_registry_path: Path | None = None,
    technical_protocol: bool = False,
    tech_approver: str = "",
    tech_program_name: str = "",
    tech_approval_date_raw: str = "",
    face_sheet_profession: str | None = None,
    v_prof_enabled_norm_keys: frozenset[str] | None = None,
    v_prof_enabled_by_fio: dict[str, frozenset[str]] | None = None,
    v_prof_main_by_fio: dict[str, str] | None = None,
) -> None:
    """Собирает DOCX из шаблона и конвертирует в PDF через Word (сохраняет оформление)."""
    fd, tmp_docx = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        doc, _ = build_filled_protocol_document(
            template_path,
            protocol_no=protocol_no,
            date_str=date_str,
            theme=theme,
            table_persons=table_persons,
            program_titles=program_titles,
            program_keys=program_keys,
            excel_path=excel_path,
            persons_v_raw=persons_v_raw,
            persons_b_row_source=persons_b_row_source,
            grade=grade,
            registry_no=registry_no,
            check_type=check_type,
            trained_registry_path=trained_registry_path,
            technical_protocol=technical_protocol,
            tech_approver=tech_approver,
            tech_program_name=tech_program_name,
            tech_approval_date_raw=tech_approval_date_raw,
            face_sheet_profession=face_sheet_profession,
            v_prof_enabled_norm_keys=v_prof_enabled_norm_keys,
            v_prof_enabled_by_fio=v_prof_enabled_by_fio,
            v_prof_main_by_fio=v_prof_main_by_fio,
        )
        save_formed_protocol_docx(doc, tmp_docx)
        tmp_p = Path(tmp_docx).resolve()
        out_p = Path(pdf_path).resolve()
        if sys.platform == "win32":
            _docx_to_pdf_via_word_com(tmp_p, out_p)
        else:
            try:
                from docx2pdf import convert
            except ImportError as e:
                raise RuntimeError(
                    "Для PDF с форматированием Word установите: pip install docx2pdf\n"
                    "и приложение Microsoft Word."
                ) from e
            convert(str(tmp_p), str(out_p))
    finally:
        try:
            os.unlink(tmp_docx)
        except OSError:
            pass


def write_protocol_docx(path: str, content: str) -> None:
    doc = Document()
    normalized = content.replace("\r\n", "\n").replace("\r", "\n")
    for line in normalized.split("\n"):
        doc.add_paragraph(line)
    doc.save(path)
