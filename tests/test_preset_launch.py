"""The preset -> launcher handoff, and the auto-place guard.

Both bugs here were reported from a clean, fully-verified install where
`--check` passed every row and every menu entry still did nothing. They are
silent no-ops, which is the failure mode worth the most test surface: nothing
crashes, nothing logs, the user just concludes the tool is broken.
"""
import os
import re
import subprocess
import sys
import unittest

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRESET_DIR = os.path.join(APP_DIR, "presets")
LAUNCHER = os.path.join(APP_DIR, "preset_launch.lua")

FUSCRIPT = ("/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/"
            "Libraries/Fusion/fuscript")


def preset_files():
    return [os.path.join(PRESET_DIR, f)
            for f in sorted(os.listdir(PRESET_DIR)) if f.endswith(".lua")]


class TestLauncherContract(unittest.TestCase):
    """The launcher must hand its behaviour over as a RETURN VALUE.

    Resolve's Lua sandbox does not expose the caller's globals to a dofile()'d
    chunk, so a launcher that reads LAB37_* globals sees nil for all of them
    and silently launches nothing.
    """

    def setUp(self):
        self.src = open(LAUNCHER, encoding="utf-8").read()

    def test_launcher_returns_a_function(self):
        self.assertRegex(self.src, r"return function\s*\(",
                         "launcher must return a function, not run at load time")

    def test_launcher_does_not_read_globals(self):
        for name in ("LAB37_TOOL", "LAB37_ARGS", "LAB37_APP_DIR"):
            self.assertNotIn(
                name, self.src,
                f"{name} is a global; it does not cross the dofile() boundary")

    def test_every_preset_calls_the_returned_function(self):
        for path in preset_files():
            src = open(path, encoding="utf-8").read()
            name = os.path.basename(path)
            self.assertIn("dofile(launcher)", src, f"{name}: no launcher load")
            self.assertRegex(
                src, r"local\s+launch\s*=\s*dofile\(launcher\)",
                f"{name}: must capture the launcher's return value")
            self.assertRegex(
                src, r"launch\(\s*TOOL\s*,\s*ARGS\s*,\s*app_dir\s*\)",
                f"{name}: must call the launcher with explicit arguments")

    def test_every_preset_guards_against_a_stale_launcher(self):
        # Pulling without re-running setup leaves the installed presets and
        # the app-dir launcher out of step; that must be a message, not
        # another silent no-op.
        for path in preset_files():
            src = open(path, encoding="utf-8").read()
            self.assertIn('type(launch) ~= "function"', src,
                          f"{os.path.basename(path)}: no stale-launcher guard")

    def test_presets_do_not_assign_lab37_globals(self):
        for path in preset_files():
            src = open(path, encoding="utf-8").read()
            # re.M matters: without it "^" only matches the start of the
            # whole file, so this assertion could never fail.
            self.assertIsNone(
                re.search(r"^\s*LAB37_\w+\s*=", src, re.M),
                f"{os.path.basename(path)}: sets a global that cannot cross")


@unittest.skipUnless(os.path.exists(FUSCRIPT), "Resolve's fuscript not installed")
class TestUnderResolveInterpreter(unittest.TestCase):
    """Run the real thing under Resolve's own Lua, which is where the
    sandbox behaviour that caused this bug actually lives."""

    def _run(self, lua_path):
        return subprocess.run([FUSCRIPT, "-l", "lua", lua_path],
                              capture_output=True, text=True, timeout=60)

    def test_globals_really_do_not_cross_dofile(self):
        # Pins the platform behaviour the fix is built on. If a future Resolve
        # changes this, the fix is still correct but the reasoning is stale.
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            helper = os.path.join(d, "helper.lua")
            caller = os.path.join(d, "caller.lua")
            with open(helper, "w") as fh:
                fh.write('print("SAW=" .. tostring(LAB37_PROBE))\n')
            with open(caller, "w") as fh:
                fh.write(f'LAB37_PROBE = "set"\ndofile("{helper}")\n')
            out = self._run(caller)
            self.assertIn("SAW=nil", out.stdout,
                          "globals now cross dofile(); revisit preset_launch")

    def test_preset_reaches_the_launcher_with_its_arguments(self):
        # The launcher aborts before spawning anything if the venv python is
        # missing, so point the preset at a directory that has no venv: it
        # proves the handoff happened without starting a transcription.
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            with open(os.path.join(d, "resolve_whisper_path.txt"), "w") as fh:
                fh.write(d)
            # The launcher must be found in that fake app dir.
            import shutil
            shutil.copy(LAUNCHER, os.path.join(d, "preset_launch.lua"))
            preset = os.path.join(d, "LAB37 Check.lua")
            shutil.copy(os.path.join(PRESET_DIR, "LAB37 Check.lua"), preset)

            out = self._run(preset)
            combined = out.stdout + out.stderr
            # "[LAB37 Check]" proves TOOL crossed; the venv error proves
            # app_dir crossed. The old bug printed "[LAB37 Whisper] ERROR:
            # app dir not resolved" instead.
            self.assertIn("[LAB37 Check]", combined, combined)
            self.assertIn("Python venv not found", combined, combined)
            self.assertNotIn("app dir not resolved", combined, combined)


class TestAutoPlaceGuard(unittest.TestCase):
    """The guard exists to avoid stacking onto hand-edited captions. It must
    key off subtitle ITEMS, not tracks: Resolve leaves an empty 'Subtitle 1'
    behind for anyone who has ever touched subtitles, and those users were
    being sent to a Media Pool fallback that does not reliably work."""

    def setUp(self):
        self.src = open(os.path.join(APP_DIR, "caption.py"), encoding="utf-8").read()

    def test_guard_counts_items_not_track_count(self):
        self.assertIn("existing_items = _subtitle_item_count()", self.src)
        # The OUTER decision is on items. AddTrack is still gated on track
        # count inside it, so match on ordering rather than forbidding a
        # substring the fix legitimately contains.
        block = self.src.split("if imported_to_pool and same_timeline:", 1)[1][:1200]
        items_at = block.find("if existing_items == 0:")
        tracks_at = block.find("if existing_sub_tracks == 0:")
        self.assertGreater(items_at, -1, "auto-place is not gated on item count")
        self.assertGreater(tracks_at, items_at,
                           "track count is still the outer guard; an empty "
                           "Subtitle 1 would block auto-place")

    def test_track_is_only_added_when_none_exists(self):
        # AppendToTimeline targets subtitle track 1, so adding a second track
        # when an empty one exists would place onto a track nobody is viewing.
        self.assertRegex(
            self.src,
            r"if existing_sub_tracks == 0:\s*\n\s*_safe\(timeline\.AddTrack, \"subtitle\"\)")

    def test_item_count_spans_every_track(self):
        # Slice by size, not by blank line: the function has a docstring.
        fn = self.src.split("def _subtitle_item_count():", 1)[1][:800]
        self.assertIn("for idx in range(1, tracks + 1):", fn,
                      "must count all subtitle tracks, not just track 1")


class TestVersionStamp(unittest.TestCase):
    """A bug report from another machine is only actionable with a build id."""

    def test_version_module_reports_something(self):
        sys.path.insert(0, APP_DIR)
        import version
        v = version.version_string()
        self.assertTrue(v and isinstance(v, str))
        # Either a git build id, or an explicit statement that it is unknown.
        self.assertTrue(re.match(r"^[0-9a-f]{7,}\s\(\d{4}-\d{2}-\d{2}\)", v)
                        or v.startswith("unknown"), v)

    def test_run_log_is_stamped(self):
        src = open(os.path.join(APP_DIR, "caption.py"), encoding="utf-8").read()
        self.assertIn("resolve-whisper build", src,
                      "every run log must carry the build id")


if __name__ == "__main__":
    unittest.main()
