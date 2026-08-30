"""The Deliver page must survive a run untouched.

The tool drives the render queue to get timeline audio, which mutates project
state the user owns: the render preset, format/codec, output path, the queue
itself, and the page they are looking at. Every one of those has to be handed
back. These are source-shape tests because the behaviour lives in calls to a
Resolve API that isn't importable without Resolve running; the live check is
`caption.py --check`.
"""
import os
import re
import unittest

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = open(os.path.join(APP_DIR, "caption.py"), encoding="utf-8").read()


def function_body(name):
    lines = SRC.split("\n")
    start = next((i for i, l in enumerate(lines)
                  if l.startswith(f"def {name}(")), None)
    if start is None:
        raise AssertionError(
            f"caption.py no longer defines {name}(); this test guards its "
            "behaviour, so a rename has to be noticed, not swallowed")
    end = next((i for i, l in enumerate(lines[start + 1:], start + 1)
                if l and not l[0].isspace()), len(lines))
    return "\n".join(lines[start:end])


class TestRenderQueueHygiene(unittest.TestCase):
    def test_only_our_own_jobs_are_deleted(self):
        body = function_body("_delete_job_if_ours")
        self.assertIn("if job_id in pre_existing_ids:", body,
                      "must refuse to delete jobs the user queued")
        self.assertIn("return", body)

    def test_every_added_job_is_deleted(self):
        # Each AddRenderJob has to have a matching cleanup. The probe job
        # (used to read back the user's output path) is the easy one to
        # forget, and it would accumulate one queue entry per run.
        for fn in ("render_audio", "run_check_mode"):
            body = function_body(fn)
            # Count real calls only: the body also mentions AddRenderJob in
            # a comment and in a log message.
            adds = len(re.findall(r"project\.AddRenderJob\b", body))
            deletes = len(re.findall(r"_delete_job_if_ours\(", body))
            self.assertGreaterEqual(
                deletes, adds,
                f"{fn}: {adds} AddRenderJob calls but only {deletes} cleanups")

    def test_render_path_cleans_up_in_a_finally(self):
        body = function_body("render_audio")
        tail = body.split("finally:", 1)
        self.assertEqual(len(tail), 2, "render_audio must have a finally block")
        self.assertIn("_delete_job_if_ours", tail[1])
        self.assertIn("_restore_deliver_state", tail[1])


class TestDeliverSettingsRestored(unittest.TestCase):
    def test_state_is_snapshotted_before_mutation(self):
        body = function_body("render_audio")
        # Search actual calls: the docstring names LoadRenderPreset too.
        snap = body.find("_safe(project.GetCurrentRenderFormatAndCodec")
        load = body.find("_safe(project.LoadRenderPreset")
        self.assertGreater(load, snap,
                           "the preset is loaded before the snapshot is taken")

    def test_user_output_path_is_never_blanked_blindly(self):
        body = function_body("_restore_deliver_state")
        self.assertIn("if saved_target_dir is not None or saved_custom_name is not None:",
                      body,
                      "restoring must not clear the user's own output path "
                      "when we never captured one")

    def test_restore_call_sites_pass_both_conditions(self):
        # Asserting the identifier exists would pass even if the flag were
        # never handed to the restore, so match the actual argument.
        for fn, expr in (("render_audio", "snapshot_failed and output_mutated"),
                         ("run_check_mode", "_snapshot_failed and _output_mutated")):
            self.assertIn(expr, function_body(fn),
                          f"{fn} does not gate clearing on both conditions")

    def test_check_mode_restores_too(self):
        # A health check that edits the user's project is worse than no check.
        body = function_body("run_check_mode")
        self.assertIn("_restore_deliver_state", body)
        self.assertIn("finally:", body)


class TestUserIsNotStrandedOnDeliver(unittest.TestCase):
    """Resolve switches itself to Deliver during AddRenderJob/StartRendering.
    Anything that returns early has to put the user back."""

    def test_page_restore_helper_exists(self):
        body = function_body("_restore_page")
        self.assertIn("GetCurrentPage", body)
        self.assertIn("OpenPage", body)

    def test_helper_is_a_noop_without_a_snapshot(self):
        body = function_body("_restore_page")
        self.assertIn("if not resolve or not saved_page:", body,
                      "must not navigate when it never learned the origin page")

    def test_resolve_mode_restores_page_in_finally(self):
        body = function_body("run_resolve_mode")
        self.assertIn("saved_page = _safe(resolve.GetCurrentPage)", body)
        tail = body.split("finally:", 1)
        self.assertEqual(len(tail), 2)
        self.assertIn("_restore_page(resolve, saved_page)", tail[1],
                      "failure paths leave the user on the Deliver page")

    def test_success_path_does_not_fight_the_restore(self):
        # The happy path deliberately lands on Edit (captions are there); the
        # finally must not then drag the user back to wherever they started.
        body = function_body("run_resolve_mode")
        self.assertIn('saved_page = "edit"', body)
        self.assertLess(body.find('saved_page = "edit"'),
                        body.find('resolve.OpenPage("edit")'))

    def test_check_mode_restores_page(self):
        body = function_body("run_check_mode")
        self.assertIn("_restore_page(resolve, saved_page)", body)


if __name__ == "__main__":
    unittest.main()
