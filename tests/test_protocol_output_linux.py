# -*- coding: utf-8 -*-
"""Тесты Linux-адаптаций protocol_output (оверлей linux_port)."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from _bootstrap import setup_linux_port_app_paths


class TestLinuxProtocolOutput(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        try:
            setup_linux_port_app_paths()
        except RuntimeError as e:
            raise unittest.SkipTest(str(e)) from e

    def test_cyrillic_ttf_candidates_linux_includes_bundle_fonts_dir(self) -> None:
        from protocol_output import _linux_font_dirs, cyrillic_ttf_candidates

        with mock.patch("protocol_output.sys.platform", "linux"):
            with mock.patch("protocol_output._linux_cyrillic_ttf_candidates") as mocked:
                mocked.return_value = [Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")]
                result = cyrillic_ttf_candidates()
        self.assertEqual(len(result), 1)
        dirs = _linux_font_dirs()
        self.assertTrue(any("fonts" in str(d) for d in dirs))

    def test_libreoffice_executable_none_when_missing(self) -> None:
        from protocol_output import _libreoffice_executable

        with mock.patch("protocol_output.shutil.which", return_value=None):
            self.assertIsNone(_libreoffice_executable())

    def test_libreoffice_executable_prefers_libreoffice(self) -> None:
        from protocol_output import _libreoffice_executable

        with mock.patch(
            "protocol_output.shutil.which",
            side_effect=lambda name: "/usr/bin/libreoffice" if name == "libreoffice" else None,
        ):
            self.assertEqual(_libreoffice_executable(), "/usr/bin/libreoffice")

    def test_docx_to_pdf_non_windows_tries_libreoffice_first(self) -> None:
        from protocol_output import _docx_to_pdf_non_windows

        with tempfile.TemporaryDirectory() as tmp:
            docx = Path(tmp) / "test.docx"
            pdf = Path(tmp) / "test.pdf"
            docx.write_bytes(b"docx")
            calls: list[str] = []

            def fake_lo(d: Path, p: Path) -> None:
                calls.append("lo")
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_bytes(b"%PDF")

            with mock.patch("protocol_output._docx_to_pdf_via_libreoffice", side_effect=fake_lo):
                with mock.patch("protocol_output._docx_to_pdf_via_docx2pdf") as d2p:
                    _docx_to_pdf_non_windows(docx, pdf)
                    d2p.assert_not_called()
            self.assertEqual(calls, ["lo"])

    def test_docx_to_pdf_non_windows_falls_back_to_docx2pdf(self) -> None:
        from protocol_output import _docx_to_pdf_non_windows

        with tempfile.TemporaryDirectory() as tmp:
            docx = Path(tmp) / "test2.docx"
            pdf = Path(tmp) / "test2.pdf"
            docx.write_bytes(b"docx")

            def fake_d2p(d: Path, p: Path) -> None:
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_bytes(b"%PDF")

            with mock.patch(
                "protocol_output._docx_to_pdf_via_libreoffice",
                side_effect=RuntimeError("no lo"),
            ):
                with mock.patch("protocol_output._docx_to_pdf_via_docx2pdf", side_effect=fake_d2p):
                    _docx_to_pdf_non_windows(docx, pdf)
            self.assertTrue(pdf.is_file())


if __name__ == "__main__":
    unittest.main()
