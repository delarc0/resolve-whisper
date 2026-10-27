"""The preset-repair script, and progress copy.

create_audio_only_preset.py is what the health check tells users to run when
the render preset looks wrong, so it is a remedy: it has to work, and it has
to leave the user's Deliver page exactly as it found it.
"""
import os
from pathlib import Path
import sys
import unittest

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, APP_DIR)

from resolve_api import _audio_codec_warning  # noqa: E402


def _src(name):
    return (Path(APP_DIR) / name).read_text(encoding="utf-8")


PRESET_SCRIPT = _src("create_audio_only_preset.py")
CAPTION = _src("caption.py")
CHECK = _src("check.py")
RESOLVE_API = _src("resolve_api.py")
UPDATE_SH = _src("update.sh")


class TestPresetScriptIsRunnable(unittest.TestCase):
    def test_bootstraps_the_resolve_module_path(self):
        # Without this it exits immediately with "not on PYTHONPATH", which
        # made the remedy the health check prints impossible to follow.
        self.assertIn("platforminfo.bootstrap_resolve_env()", PRESET_SCRIPT)
        boot = PRESET_SCRIPT.find("bootstrap_resolve_env()")
        imp = PRESET_SCRIPT.find("import DaVinciResolveScript")
        self.assertLess(boot, imp, "bootstrap must run before the import")

    def test_explains_a_factory_preset_instead_of_a_falsy_return(self):
        # Resolve 21 ships 'Audio Only' as a factory preset: it cannot be
        # deleted, and audio formats were removed from
        # SetCurrentRenderFormatAndCodec, so --force cannot succeed there.
        self.assertIn("factory preset on Resolve", PRESET_SCRIPT)


class TestPresetScriptRestoresDeliver(unittest.TestCase):
    def test_snapshots_the_users_output_path(self):
        # _validate() overwrites TargetDir with a temp dir to make
        # AddRenderJob work on Resolve 21. Without a snapshot, that temp path
        # is what the user finds in their Deliver page afterwards.
        self.assertIn("saved_target_dir", PRESET_SCRIPT)
        self.assertIn("saved_custom_name", PRESET_SCRIPT)
        self.assertIn("AddRenderJob", PRESET_SCRIPT)

    def test_passes_the_snapshot_to_the_restore(self):
        self.assertIn("saved_target_dir, saved_custom_name,\n"
                      "                               snapshot_failed and output_mutated)",
                      PRESET_SCRIPT)

    def test_a_failed_probe_does_not_leave_our_path_behind(self):
        # If AddRenderJob returns no id we never learn the old TargetDir, but
        # _validate() still overwrites it. Leaving ours there would redirect
        # the user's next export.
        self.assertIn("snapshot_failed = True", PRESET_SCRIPT)
        self.assertIn("snapshot_failed = False", PRESET_SCRIPT)
        # ...and clearing is gated on having actually changed the fields.
        self.assertIn("output_mutated = True", PRESET_SCRIPT)

    def test_restores_the_page_too(self):
        self.assertIn("_restore_page(resolve, saved_page)", PRESET_SCRIPT)

    def test_probe_job_is_deleted_in_a_finally(self):
        tail = PRESET_SCRIPT.split("_probe_id = _safe(project.AddRenderJob)", 1)[1][:600]
        self.assertIn("finally:", tail)
        self.assertIn("DeleteRenderJob", tail)


class FakeResolve:
    """Just enough Resolve for the version gate."""

    def __init__(self, version):
        self._version = version

    def GetVersionString(self):
        return self._version


class TestCodecAdviceIsActionable(unittest.TestCase):
    """A warning telling someone to run a command that cannot work on their
    version is worse than no warning."""

    def test_resolve_21_is_told_there_is_nothing_to_fix(self):
        note = _audio_codec_warning({"AudioCodec": "aac"}, FakeResolve("21.1.1"))
        self.assertIn("nothing to fix", note)
        self.assertNotIn("create_audio_only_preset", note)

    def test_older_resolve_is_told_wav_is_faster(self):
        note = _audio_codec_warning({"AudioCodec": "aac"}, FakeResolve("18.6.6"))
        self.assertIn("WAV/PCM preset transcribes faster", note)
        self.assertNotIn("nothing to fix", note)

    def test_pcm_says_nothing_at_all(self):
        self.assertEqual("", _audio_codec_warning({"AudioCodec": "LinearPCM"},
                                                  FakeResolve("21.1.1")))

    def test_render_path_passes_the_resolve_handle(self):
        # Without it the 21 branch above is unreachable on the render path,
        # and Resolve 21 users were told to run a repair that cannot work.
        self.assertIn("_audio_codec_warning(job_settings, _safe(get_resolve))",
                      RESOLVE_API)

    def test_recreate_advice_is_gated_below_21(self):
        self.assertIn("if _resolve_major(resolve) < 21:", CHECK)
        idx = CHECK.find("if _resolve_major(resolve) < 21:")
        following = CHECK[idx:idx + 400]
        self.assertIn("create_audio_only_preset.py --force", following,
                      "the recreate advice must sit inside the version gate")


class TestProgressCopyDoesNotRepeatItself(unittest.TestCase):
    """The window shows a stage title and a detail line. A detail that just
    restates the title reads as a stuck window and wastes the only line
    available for real information."""

    def test_transcribing_detail_does_not_restate_the_title(self):
        self.assertNotIn('_write_status("transcribing", "Transcribing audio..."',
                         CAPTION)

    def test_rendering_detail_does_not_restate_the_title(self):
        self.assertNotIn('f"Rendering... {pct_int}%"', RESOLVE_API)
        self.assertIn('status("rendering_audio", f"{pct_int}% complete"',
                      RESOLVE_API)


class TestStashNoticeSurvivesSetupOutput(unittest.TestCase):
    def test_notice_is_repeated_after_setup(self):
        # Setup prints ~40 lines, so a notice printed before it has scrolled
        # off by the time anyone stops reading.
        self.assertIn('if [ -n "$STASHED" ]; then', UPDATE_SH)
        setup_at = UPDATE_SH.find('"$APP_DIR/setup.sh"')
        notice_at = UPDATE_SH.find('if [ -n "$STASHED" ]; then')
        self.assertGreater(notice_at, setup_at,
                           "the reminder must come after setup's output")

    def test_a_failed_stash_stops_the_update(self):
        self.assertIn("could not set your local changes aside", UPDATE_SH)


if __name__ == "__main__":
    unittest.main()
