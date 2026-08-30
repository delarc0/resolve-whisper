"""Standalone Tk progress window for Resolve Whisper.

Polls a JSON status file every 200ms; updates the title, message, and progress
bar; auto-closes on stage='done' or stage='error'. Designed to be spawned
as a subprocess so it doesn't block the transcription pipeline.

Styled per apps/DESIGN.md (LAB37 charcoal shell) via ui_theme. The progress
bar is Frame-based: native ttk.Progressbar renders Aqua blue and ignores
styling on macOS.

Usage (the path is always supplied by caption.py; it is per-run):
    python progress_ui.py /tmp/resolve_whisper_status.<pid>.json
"""
import json
import os
import signal
import sys
import time
import tempfile

try:
    import tkinter as tk
except ImportError:
    sys.stderr.write("tkinter not available; UI disabled.\n")
    sys.exit(0)

import ui_theme as th


# Status files are per-run (resolve_whisper_status.<pid>.json) and the path
# is always passed as argv by caption.py. There is no meaningful default any
# more: polling a fixed name would read a stale file no run will update.
DEFAULT_STATUS_FILE = None

# Map internal stage names to user-friendly titles.
_STAGE_TITLES = {
    "starting": "Starting...",
    "rendering_audio": "Rendering timeline audio",
    "loading_model": "Loading AI model",
    "transcribing": "Transcribing audio",
    "writing_srt": "Writing captions",
    "done": "Done",
    "error": "Error",
}


class _Bar:
    """Flat determinate/indeterminate progress bar built from Frames."""

    SWEEP_W = 0.28          # indeterminate segment width (fraction)
    SWEEP_STEP = 0.018      # per tick
    TICK_MS = 30

    def __init__(self, parent, height=6):
        self.track = tk.Frame(parent, bg=th.INPUT, height=height)
        self.track.pack(fill="x")
        self.fill = tk.Frame(self.track, bg=th.FG)
        self.fill.place(x=0, y=0, relheight=1.0, relwidth=0.0)
        self._anim = None
        self._pos = -self.SWEEP_W

    def set(self, pct, color=None):
        self._stop_anim()
        self.fill.config(bg=color or th.FG)
        self.fill.place_configure(relx=0.0,
                                  relwidth=max(0.0, min(pct, 100)) / 100.0)

    def indeterminate(self):
        if self._anim is None:
            self.fill.config(bg=th.FG)
            self._tick()

    def _tick(self):
        self._pos += self.SWEEP_STEP
        if self._pos > 1.0:
            self._pos = -self.SWEEP_W
        relx = max(0.0, self._pos)
        relw = min(self._pos + self.SWEEP_W, 1.0) - relx
        self.fill.place_configure(relx=relx, relwidth=max(relw, 0.0))
        self._anim = self.track.after(self.TICK_MS, self._tick)

    def _stop_anim(self):
        if self._anim is not None:
            self.track.after_cancel(self._anim)
            self._anim = None


class ProgressUI:
    POLL_INTERVAL_MS = 200
    STALE_AFTER_S = 90  # auto-close if status file goes silent for 90s
    DONE_LINGER_MS = 1500
    ERROR_LINGER_MS = 6000

    def __init__(self, status_file: str):
        self.status_file = status_file
        self.last_mtime = 0.0
        self.last_seen = time.time()
        self._closing = False
        self._cancelling = False
        # Latched on the first status read. A lingering window must never
        # act on a pid that isn't the run it was opened for: cancel ->
        # relaunch is a common sequence, and the old shared status file let
        # the dead window's Cancel/X kill the NEW run.
        self._pid = None

        self.root = tk.Tk()
        self.root.title("LAB37 TOOLS: Whisper")
        self.root.configure(bg=th.BG)
        self.root.resizable(False, False)
        try:
            self.root.attributes("-topmost", True)
        except tk.TclError:
            pass
        th.edge_light(self.root)

        frame = tk.Frame(self.root, bg=th.BG, padx=28, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=th.BRAND, font=(th.FONT_MONO, 9),
                 fg=th.MUTED, bg=th.BG).pack(anchor="w")

        self.title_label = tk.Label(
            frame, text="Connecting to Resolve...",
            font=(th.FONT_UI, 13, "bold"), fg=th.FG, bg=th.BG)
        self.title_label.pack(anchor="w", pady=(6, 0))

        self.detail_label = tk.Label(
            frame, text="", font=(th.FONT_UI, 10), fg=th.MUTED, bg=th.BG,
            wraplength=384, justify="left")
        self.detail_label.pack(anchor="w", pady=(2, 14))

        bar_holder = tk.Frame(frame, bg=th.BG, width=384)
        bar_holder.pack(fill="x")
        self.bar = _Bar(bar_holder)
        self.bar.indeterminate()
        self._mode = "indeterminate"

        # Update notice: created empty and only packed once there is
        # something to say, so it costs no vertical space on a normal run.
        self.notice_label = tk.Label(
            frame, text="", font=(th.FONT_UI, 10), fg=th.MUTED, bg=th.BG,
            wraplength=384, justify="left")
        self._notice_shown = False

        button_row = tk.Frame(frame, bg=th.BG)
        button_row.pack(fill="x", pady=(16, 0))
        self._button_row = button_row
        self.cancel_button = th.make_button(
            button_row, "Cancel", "secondary", self._on_cancel)
        self.cancel_button.pack(side="right")

        # Closing the window via the red X also cancels.
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self.root.update_idletasks()
        self.root.minsize(440, self.root.winfo_reqheight())
        th.place_window(self.root)

        # Schedule first poll
        self.root.after(self.POLL_INTERVAL_MS, self._tick)

    def _tick(self):
        if self._closing:
            return
        try:
            try:
                mtime = os.path.getmtime(self.status_file)
            except (FileNotFoundError, OSError):
                mtime = 0

            status = None
            if mtime and mtime > self.last_mtime:
                self.last_mtime = mtime
                self.last_seen = time.time()
                try:
                    with open(self.status_file, "r") as f:
                        status = json.load(f)
                except (json.JSONDecodeError, FileNotFoundError, OSError, ValueError):
                    status = None

            if isinstance(status, dict):
                pid = status.get("pid")
                if self._pid is None:
                    self._pid = pid
                elif pid and pid != self._pid:
                    # A different run owns this file now; this window belongs
                    # to a run that is gone. Leave without touching anything.
                    self._close()
                    return
                self._apply_status(status)

            # Silent too long: only close if the run really is gone. A long
            # blocking Resolve call (or the machine sleeping mid-render) used
            # to make the window vanish while the pipeline kept running,
            # taking the Cancel button with it.
            if time.time() - self.last_seen > self.STALE_AFTER_S:
                if self._pid is None:
                    # Never saw a single status: nothing to wait for.
                    self._close()
                    return
                if self._process_alive():
                    self.last_seen = time.time()  # still working; keep waiting
                else:
                    self.title_label.config(text="Lost contact", fg=th.RED)
                    self.detail_label.config(
                        text="The caption process stopped unexpectedly. "
                             "Check the log for details.")
                    self.bar.set(100, color=th.RED)
                    self.cancel_button.set_enabled(False)
                    self.root.after(self.ERROR_LINGER_MS, self._close)
                    return
        except Exception:
            # A malformed payload must never stop the poll loop: the window
            # would sit on screen forever with no updates and no auto-close.
            pass

        self.root.after(self.POLL_INTERVAL_MS, self._tick)

    def _process_alive(self) -> bool:
        """Is the caption process still running? Unknown counts as alive."""
        if not self._pid:
            return True
        try:
            pid = int(self._pid)
        except (TypeError, ValueError):
            return True
        if sys.platform == "win32":
            # os.kill(pid, 0) TERMINATES on Windows; never probe that way.
            return True
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except (PermissionError, OSError):
            return True

    def _apply_status(self, status: dict):
        # Coerce: a foreign or older leftover file could carry a non-string
        # stage, and the title lookup below would raise mid-callback.
        stage = status.get("stage", "")
        stage = stage if isinstance(stage, str) else ""
        progress = status.get("progress", -1)
        message = status.get("message", "")
        message = message if isinstance(message, str) else ""

        # While cancelling, ignore non-terminal updates: the estimator
        # thread keeps writing "transcribing" until the SIGTERM lands, which
        # would overwrite our "Cancelling..." and make Cancel look ignored.
        if self._cancelling and stage not in ("done", "error"):
            return

        title = _STAGE_TITLES.get(stage, stage.replace("_", " ").title() or "Working...")
        self.title_label.config(text=title)
        self.detail_label.config(text=message)

        notice = status.get("notice", "")
        if notice and not self._notice_shown and isinstance(notice, str):
            self.notice_label.config(text=notice)
            # `before` keeps it above the buttons: packing order, not call
            # order, decides where it lands.
            self.notice_label.pack(anchor="w", pady=(12, 0),
                                   before=self._button_row)
            self._notice_shown = True
            self.root.update_idletasks()
            self.root.minsize(440, self.root.winfo_reqheight())

        if stage == "done":
            self.title_label.config(fg=th.SUCCESS)
            self.bar.set(100, color=th.SUCCESS)
            self._mode = "determinate"
            self.cancel_button.set_enabled(False)
            self.root.after(self.DONE_LINGER_MS, self._close)
            return
        if stage == "error":
            self.title_label.config(fg=th.RED)
            self.bar.set(100, color=th.RED)
            self._mode = "determinate"
            self.cancel_button.set_enabled(False)
            self.root.after(self.ERROR_LINGER_MS, self._close)
            return

        try:
            pct = int(progress)
        except (ValueError, TypeError):
            pct = -1

        if pct >= 0:
            self.bar.set(pct)
            self._mode = "determinate"
        else:
            if self._mode != "indeterminate":
                self.bar.indeterminate()
                self._mode = "indeterminate"

    def _on_cancel(self):
        """Send SIGTERM to the caption.py process.

        The caption script has a SIGTERM handler that converts it to a
        KeyboardInterrupt, so its existing finally blocks run and the project
        state is restored cleanly. We don't kill the UI here; we let the
        caption script's 'error' status drive the auto-close so the user
        sees the cancellation acknowledgement.
        """
        # Idempotent: an impatient second click (the red X is not gated by
        # the disabled button) delivered a second SIGTERM, which landed
        # inside the cleanup finally-blocks and left the render job queued
        # and the Deliver page pointing at a temp dir.
        if self._cancelling:
            return

        # Only ever signal the pid this window latched onto at startup.
        pid = self._pid
        if pid is None:
            try:
                with open(self.status_file, "r", encoding="utf-8") as f:
                    pid = json.load(f).get("pid")
            except (OSError, json.JSONDecodeError, ValueError):
                pid = None

        self._cancelling = True
        self.title_label.config(text="Cancelling...", fg=th.FG)
        self.detail_label.config(text="Stopping render and cleaning up.")
        self.cancel_button.set_enabled(False)

        if pid:
            try:
                os.kill(int(pid), signal.SIGTERM)
            except (ProcessLookupError, PermissionError, ValueError, OSError):
                # Process already gone; close ourselves after a beat.
                self.root.after(500, self._close)
        else:
            # No PID known yet -- caption script hasn't written status yet.
            # Best we can do is close ourselves; caption will write 'done' and
            # we'd see it but the user wants out NOW.
            self.root.after(500, self._close)

    def _close(self):
        self._closing = True
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def run(self):
        self.root.mainloop()


def main():
    if len(sys.argv) > 1:
        status_file = sys.argv[1]
    else:
        sys.stderr.write(
            "usage: progress_ui.py <status-file>\n"
            "(caption.py passes the per-run status file path)\n")
        return 2
    ProgressUI(status_file).run()
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
