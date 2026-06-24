# -*- coding: utf-8
"""Установка каталога data/ при обновлении."""

from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from update_data_installer import (  # noqa: E402
    apply_data_updates,
    data_file_destination,
)
from update_manifest import DataFilePayload, UpdateManifest, WindowsUpdatePayload  # noqa: E402


class TestUpdateDataInstaller(unittest.TestCase):
    def test_apply_data_updates_replaces_files_in_data_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            share = root / "share"
            version_dir = share / "windows" / "1.5.2"
            data_share = version_dir / "data"
            data_share.mkdir(parents=True)
            payload = b"new-template"
            (data_share / "default_protocol.docx").write_bytes(payload)
            digest = hashlib.sha256(payload).hexdigest()

            install = root / "install"
            install.mkdir()
            exe = install / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            data_local = install / "data"
            data_local.mkdir()
            (data_local / "default_protocol.docx").write_bytes(b"old")
            (install / "protocols.db").write_bytes(b"db")

            manifest_path = share / "manifest.json"
            manifest = UpdateManifest(
                latest_version="1.5.2",
                windows=WindowsUpdatePayload(
                    relative_path="windows/1.5.2/ProtocolOOT.exe",
                    sha256="00" * 32,
                    size=3,
                ),
                data_files=[
                    DataFilePayload(
                        relative_path="windows/1.5.2/data/default_protocol.docx",
                        sha256=digest,
                        size=len(payload),
                        policy="replace",
                    )
                ],
            )

            apply_data_updates(manifest_path, manifest, exe)

            self.assertEqual((data_local / "default_protocol.docx").read_bytes(), payload)
            self.assertEqual((install / "protocols.db").read_bytes(), b"db")
            self.assertFalse((install / "Data_base.xlsx").exists())

    def test_apply_data_updates_restores_backup_on_checksum_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            share = root / "share"
            version_dir = share / "windows" / "1.5.2" / "data"
            version_dir.mkdir(parents=True)
            (version_dir / "default_protocol.docx").write_bytes(b"bad")

            install = root / "install"
            install.mkdir()
            exe = install / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            data_local = install / "data"
            data_local.mkdir()
            (data_local / "default_protocol.docx").write_bytes(b"keep")

            manifest_path = share / "manifest.json"
            manifest = UpdateManifest(
                latest_version="1.5.2",
                windows=WindowsUpdatePayload(
                    relative_path="windows/1.5.2/ProtocolOOT.exe",
                    sha256="00" * 32,
                    size=3,
                ),
                data_files=[
                    DataFilePayload(
                        relative_path="windows/1.5.2/data/default_protocol.docx",
                        sha256="ab" * 32,
                        size=3,
                        policy="replace",
                    )
                ],
            )

            from update_installer import UpdateInstallerError

            with self.assertRaises(UpdateInstallerError):
                apply_data_updates(manifest_path, manifest, exe)
            self.assertEqual((data_local / "default_protocol.docx").read_bytes(), b"keep")

    def test_data_file_destination_under_data_subdir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            exe = Path(tmp) / "ProtocolOOT.exe"
            exe.write_bytes(b"x")
            entry = DataFilePayload(
                relative_path="windows/1.5.2/data/FAQ.txt",
                sha256="00" * 32,
                size=1,
            )
            dest = data_file_destination(exe, entry)
            self.assertEqual(dest.resolve(), (Path(tmp) / "data" / "FAQ.txt").resolve())


if __name__ == "__main__":
    unittest.main()
