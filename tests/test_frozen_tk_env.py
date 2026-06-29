# -*- coding: utf-8
"""Tcl/Tk окружение для frozen onefile (runtime hook)."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from frozen_tk_env import configure_frozen_tk_environment  # noqa: E402


class TestFrozenTkEnv(unittest.TestCase):
    def test_clears_inherited_and_sets_meipass_tcl(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            meipass = Path(tmp)
            tcl_dir = meipass / "_tcl_data"
            tk_dir = meipass / "_tk_data"
            tcl_dir.mkdir()
            tk_dir.mkdir()
            (tcl_dir / "init.tcl").write_text("# stub\n", encoding="ascii")

            env = {
                "TCL_LIBRARY": r"C:\Users\Old\_MEI111\_tcl_data",
                "TK_LIBRARY": r"C:\Users\Old\_MEI111\_tk_data",
            }
            with patch.dict(os.environ, env, clear=False):
                configure_frozen_tk_environment(frozen=True, meipass=str(meipass))
                self.assertEqual(
                    os.path.normcase(os.environ["TCL_LIBRARY"]),
                    os.path.normcase(str(tcl_dir)),
                )
                self.assertEqual(
                    os.path.normcase(os.environ["TK_LIBRARY"]),
                    os.path.normcase(str(tk_dir)),
                )
                self.assertNotIn("Old", os.environ["TCL_LIBRARY"])

    def test_noop_when_not_frozen(self) -> None:
        with patch.dict(os.environ, {"TCL_LIBRARY": "keep"}, clear=False):
            configure_frozen_tk_environment(frozen=False)
            self.assertEqual(os.environ.get("TCL_LIBRARY"), "keep")


if __name__ == "__main__":
    unittest.main()
