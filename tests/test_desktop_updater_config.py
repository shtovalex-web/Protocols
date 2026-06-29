# -*- coding: utf-8
"""UpdaterConfig и configure()."""

from __future__ import annotations

import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from desktop_updater import UpdaterConfig, configure, get_config  # noqa: E402
from desktop_updater.registry import try_get_config  # noqa: E402
import protocol_updater_config  # noqa: E402, F401


class TestDesktopUpdaterConfig(unittest.TestCase):
    def test_protocoloot_configured(self) -> None:
        cfg = get_config()
        self.assertEqual(cfg.app_name, "ProtocolOOT")
        self.assertEqual(cfg.exe_name, "ProtocolOOT.exe")
        self.assertEqual(cfg.env_manifest, "PROTOCOLOOT_UPDATE_MANIFEST")
        self.assertEqual(cfg.resolved_app_bundle_zip_name, "ProtocolOOT_app.zip")

    def test_configure_custom_app(self) -> None:
        previous = try_get_config()
        try:
            custom = UpdaterConfig(
                app_name="DemoApp",
                exe_name="DemoApp.exe",
                app_version="2.0.0",
                default_share_root=Path(r"\\SERVER\SOFT\DemoApp"),
                env_prefix="DEMOAPP",
                data_replace_filenames=("readme.txt",),
            )
            configure(custom)
            cfg = get_config()
            self.assertEqual(cfg.exe_name, "DemoApp.exe")
            self.assertEqual(cfg.env_force_check, "DEMOAPP_UPDATE_CHECK")
            self.assertEqual(cfg.data_replace_filenames, ("readme.txt",))
        finally:
            if previous is not None:
                configure(previous)


if __name__ == "__main__":
    unittest.main()
