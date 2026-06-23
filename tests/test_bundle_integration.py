# -*- coding: utf-8 -*-
"""Интеграция с файлами bundle/ (.odt, .ods)."""

from __future__ import annotations

import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from bundle_integration import find_bundle_asset, resolve_docx_template_path, resolve_openpyxl_workbook_path


def _workbook_path_for_openpyxl(path: Path) -> Path:
    from bundle_integration import BundleOfficeConvertError

    try:
        return resolve_openpyxl_workbook_path(path)
    except BundleOfficeConvertError as e:
        raise RuntimeError(str(e)) from e


class TestBundleIntegration(unittest.TestCase):
    def test_find_bundle_odt_and_ods(self) -> None:
        root = Path(__file__).resolve().parents[1]
        bundle = root / "bundle"
        if not bundle.is_dir():
            self.skipTest("нет каталога bundle/")
        odt_only = find_bundle_asset("default_protocol", ".odt")
        if (bundle / "default_protocol.odt").is_file():
            self.assertIsNotNone(odt_only)
            self.assertEqual(odt_only.suffix.lower(), ".odt")
        ods = find_bundle_asset("Data_base", ".xlsx", ".ods")
        if (bundle / "Data_base.ods").is_file():
            self.assertIsNotNone(ods)
            self.assertIn(ods.suffix.lower(), (".xlsx", ".ods"))

    def test_bundle_spreadsheet_resolves_ods(self) -> None:
        root = Path(__file__).resolve().parents[1]
        ods = root / "bundle" / "Programs_base.ods"
        if not ods.is_file():
            self.skipTest("нет Programs_base.ods")
        try:
            resolved = _workbook_path_for_openpyxl(ods)
        except RuntimeError as e:
            if "LibreOffice" in str(e):
                self.skipTest(str(e))
            raise
        self.assertTrue(resolved.is_file())
        self.assertEqual(resolved.suffix.lower(), ".xlsx")

    def test_resolve_docx_keeps_docx(self) -> None:
        root = Path(__file__).resolve().parents[1]
        docx = root / "bundle" / "default_protocol.docx"
        if not docx.is_file():
            self.skipTest("нет default_protocol.docx")
        self.assertEqual(resolve_docx_template_path(docx), docx.resolve())


if __name__ == "__main__":
    unittest.main()
