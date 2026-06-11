"""Standalone Tk progress window for Resolve Whisper.

Polls a JSON status file every 200ms; updates the title, message, and progress
bar; auto-closes on stage='done' or stage='error'. Designed to be spawned
as a subprocess so it doesn't block the transcription pipeline.

Usage:
    python progress_ui.py /tmp/resolve_whisper_status.json
"""
import json
import os
import signal
import sys
import time
import tempfile

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:
    sys.stderr.write("tkinter not available; UI disabled.\n")
    sys.exit(0)


DEFAULT_STATUS_FILE = os.path.join(tempfile.gettempdir(), "resolve_whisper_status.json")

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

        self.root = tk.Tk()
        self.root.title("LAB37 Resolve Whisper")
        self.root.geometry("440x170")
        self.root.resizable(False, False)
        try:
            self.root.attributes("-topmost", True)
        except tk.TclError:
            pass

        # Center on screen
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = (sw - 440) // 2
        y = (sh - 170) // 3  # upper third feels less intrusive than dead-center
        self.root.geometry(f"440x170+{x}+{y}")

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        self.title_label = ttk.Label(
            frame, text="Connecting to Resolve...", font=("Helvetica", 14, "bold")
        )
        self.title_label.pack(anchor="w")

        self.detail_label = ttk.Label(frame, text="", font=("Helvetica", 11))
        self.detail_label.pack(anchor="w", pady=(2, 12))

        self.progress = ttk.Progressbar(frame, length=400, mode="indeterminate")
        self.progress.pack(fill="x")
        self.progress.start(12)
        self._mode = "indeterminate"

        button_row = ttk.Frame(frame)
        button_row.pack(fill="x", pady=(12, 0))
        self.cancel_button = ttk.Button(button_row, text="Cancel", command=self._on_cancel)
        self.cancel_button.pack(side="right")

        # Closing the window via the red X also cancels.
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # Schedule first poll
        self.root.after(self.POLL_INTERVAL_MS, self._tick)

    def _tick(self):
        if self._closing:
            return
        try:
            mtime = os.path.getmtime(self.status_file)
        except FileNotFoundError:
            mtime = 0

        if mtime and mtime > self.last_mtime:
            self.last_mtime = mtime
            self.last_seen = time.time()
            try:
                with open(self.status_file, "r") as f:
                    status = json.load(f)
                self._apply_status(status)
            except (json.JSONDecodeError, FileNotFoundError, OSError):
                pass

        # Auto-close if silent too long
        if time.time() - self.last_seen > self.STALE_AFTER_S:
            self._close()
            return

        self.root.after(self.POLL_INTERVAL_MS, self._tick)

    def _apply_status(self, status: dict):
        stage = status.get("stage", "")
        progress = status.get("progress", -1)
        message = status.get("message", "")

        # While cancelling, ignore non-terminal updates: the estimator
        # thread keeps writing "transcribing" until the SIGTERM lands, which
        # would overwrite our "Cancelling..." and make Cancel look ignored.
        if self._cancelling and stage not in ("done", "error"):
            return

        title = _STAGE_TITLES.get(stage, stage.replace("_", " ").title() or "Working...")
        self.title_label.config(text=title)
        self.detail_label.config(text=message)

        if stage == "done":
            self.progress.stop()
            self.progress.config(mode="determinate", value=100)
            self._mode = "determinate"
            self.root.after(self.DONE_LINGER_MS, self._close)
            return
        if stage == "error":
            self.progress.stop()
            self.progress.config(mode="determinate", value=0)
            self._mode = "determinate"
            self.root.after(self.ERROR_LINGER_MS, self._close)
            return

        try:
            pct = int(progress)
        except (ValueError, TypeError):
            pct = -1

        if pct >= 0:
            if self._mode != "determinate":
                self.progress.stop()
                self.progress.config(mode="determinate")
                self._mode = "determinate"
            self.progress.config(value=max(0, min(pct, 100)))
        else:
            if self._mode != "indeterminate":
                self.progress.config(mode="indeterminate")
                self.progress.start(12)
                self._mode = "indeterminate"

    def _on_cancel(self):
        """Send SIGTERM to the caption.py process.

        The caption script has a SIGTERM handler that converts it to a
        KeyboardInterrupt, so its existing finally blocks run and the project
        state is restored cleanly. We don't kill the UI here; we let the
        caption script's 'error' status drive the auto-close so the user
        sees the cancellation acknowledgement.
        """
        try:
            with open(self.status_file, "r", encoding="utf-8") as f:
                pid = json.load(f).get("pid")
        except (OSError, json.JSONDecodeError):
            pid = None

        self._cancelling = True
        self.title_label.config(text="Cancelling...")
        self.detail_label.config(text="Stopping render and cleaning up.")
        self.cancel_button.config(state="disabled")

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
    status_file = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_STATUS_FILE
    ProgressUI(status_file).run()


if __name__ == "__main__":
    main()
