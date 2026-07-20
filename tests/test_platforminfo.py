"""Tests for the platform boundary (platforminfo.py).

These lock down the per-OS paths and the Resolve-env bootstrap so a change
to one platform can't silently break another. OS is monkeypatched, so the
Windows/Linux branches are exercised on the Mac test runner.
"""
import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import platforminfo as pi  # noqa: E402


class TestOsFlagsExclusive(unittest.TestCase):
    def test_exactly_one_flag_is_true(self):
        self.assertEqual(sum([pi.IS_MAC, pi.IS_WIN, pi.IS_LINUX]), 1)


class TestScriptsDir(unittest.TestCase):
    def test_mac_dir(self):
        with mock.patch.multiple(pi, IS_MAC=True, IS_WIN=False, IS_LINUX=False):
            self.assertTrue(pi.resolve_scripts_dir().endswith(
                "Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Edit"))

    def test_windows_dir_uses_appdata(self):
        with mock.patch.multiple(pi, IS_MAC=False, IS_WIN=True, IS_LINUX=False), \
             mock.patch.dict(os.environ, {"APPDATA": r"C:\Users\e\AppData\Roaming"}):
            d = pi.resolve_scripts_dir()
            self.assertIn("AppData", d)
            self.assertTrue(d.endswith(os.path.join(
                "DaVinci Resolve", "Support", "Fusion", "Scripts", "Edit")))

    def test_linux_dir(self):
        with mock.patch.multiple(pi, IS_MAC=False, IS_WIN=False, IS_LINUX=True):
            self.assertIn("DaVinciResolve", pi.resolve_scripts_dir())


class TestScriptApiAndLib(unittest.TestCase):
    def test_windows_lib_is_dll(self):
        with mock.patch.multiple(pi, IS_MAC=False, IS_WIN=True, IS_LINUX=False), \
             mock.patch.dict(os.environ, {"ProgramFiles": r"C:\Program Files",
                                          "PROGRAMDATA": r"C:\ProgramData"}):
            self.assertTrue(pi.resolve_script_lib().endswith("fusionscript.dll"))
            self.assertIn("Support", pi.resolve_script_api())

    def test_mac_lib_is_so(self):
        with mock.patch.multiple(pi, IS_MAC=True, IS_WIN=False, IS_LINUX=False):
            self.assertTrue(pi.resolve_script_lib().endswith("fusionscript.so"))


class TestFindTool(unittest.TestCase):
    def test_prefers_path(self):
        with mock.patch("shutil.which", return_value="/usr/bin/ffmpeg"):
            self.assertEqual(pi.find_tool("ffmpeg"), "/usr/bin/ffmpeg")

    def test_falls_back_to_known_location(self):
        with mock.patch("shutil.which", return_value=None), \
             mock.patch.multiple(pi, IS_MAC=True, IS_WIN=False, IS_LINUX=False), \
             mock.patch("os.path.exists",
                        side_effect=lambda p: p == "/opt/homebrew/bin/ffmpeg"):
            self.assertEqual(pi.find_tool("ffmpeg"), "/opt/homebrew/bin/ffmpeg")

    def test_returns_none_when_missing(self):
        with mock.patch("shutil.which", return_value=None), \
             mock.patch("os.path.exists", return_value=False):
            self.assertIsNone(pi.find_tool("ffmpeg"))

    def test_windows_candidates_have_exe(self):
        with mock.patch.multiple(pi, IS_MAC=False, IS_WIN=True, IS_LINUX=False):
            cands = pi._tool_candidates("ffmpeg")
            self.assertTrue(all(c.endswith("ffmpeg.exe") for c in cands))


class TestAddCudaDllDir(unittest.TestCase):
    def test_noop_off_windows(self):
        # On non-Windows this must do nothing and never raise, even if torch
        # is importable. add_dll_directory only exists on Windows anyway.
        with mock.patch.multiple(pi, IS_WIN=False):
            pi.add_cuda_dll_dir()  # should simply return

    def test_windows_missing_torch_is_safe(self):
        with mock.patch.multiple(pi, IS_WIN=True), \
             mock.patch.dict("sys.modules", {"torch": None}):
            # importing torch raises -> caught, no exception propagates
            pi.add_cuda_dll_dir()


class TestBootstrapResolveEnv(unittest.TestCase):
    def test_sets_env_when_absent(self):
        with mock.patch.dict(os.environ, {}, clear=True), \
             mock.patch.object(pi, "resolve_script_api", return_value="/api"), \
             mock.patch.object(pi, "resolve_script_lib", return_value="/lib.so"), \
             mock.patch("os.path.isdir", return_value=False):
            pi.bootstrap_resolve_env()
            self.assertEqual(os.environ["RESOLVE_SCRIPT_API"], "/api")
            self.assertEqual(os.environ["RESOLVE_SCRIPT_LIB"], "/lib.so")

    def test_honors_preset_env(self):
        with mock.patch.dict(os.environ,
                             {"RESOLVE_SCRIPT_API": "/custom/api"}, clear=True), \
             mock.patch.object(pi, "resolve_script_lib", return_value="/lib.so"), \
             mock.patch("os.path.isdir", return_value=False):
            pi.bootstrap_resolve_env()
            self.assertEqual(os.environ["RESOLVE_SCRIPT_API"], "/custom/api")

    def test_adds_modules_to_syspath(self):
        saved = list(sys.path)
        try:
            with mock.patch.dict(os.environ, {}, clear=True), \
                 mock.patch.object(pi, "resolve_script_api", return_value="/api"), \
                 mock.patch.object(pi, "resolve_script_lib", return_value="/lib.so"), \
                 mock.patch("os.path.isdir", return_value=True):
                pi.bootstrap_resolve_env()
                self.assertIn(os.path.join("/api", "Modules"), sys.path)
        finally:
            sys.path[:] = saved


if __name__ == "__main__":
    unittest.main()
