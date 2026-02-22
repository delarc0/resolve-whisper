"""
LAB37 TOOLS: Resolve Whisper
AI-powered caption generation for DaVinci Resolve Studio.
Runs from Workspace > Scripts > LAB37 Resolve Whisper
"""
import os
import sys
import traceback

# --- Debug log (Resolve swallows errors) ---
_LOG_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "resolve_whisper_debug.log")

def _log(msg):
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{msg}\n")
    except Exception:
        pass

def _log_clear():
    try:
        with open(_LOG_PATH, "w", encoding="utf-8") as f:
            f.write("LAB37 TOOLS: Resolve Whisper - Debug Log\n")
            f.write(f"Python: {sys.executable}\n")
            f.write(f"Version: {sys.version}\n")
            f.write(f"Platform: {sys.platform}\n")
            f.write("---\n")
    except Exception:
        pass

_log_clear()
_log("Script starting...")

try:
    import subprocess
    import tempfile
    import time
    _log("Imports OK")
except Exception as e:
    _log(f"Import failed: {e}")

# --- Locate our install directory ---
try:
    _SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    _log(f"Script dir: {_SCRIPT_DIR}")
except Exception:
    _SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv else os.getcwd()
    _log(f"Script dir (fallback): {_SCRIPT_DIR}")

_POINTER = os.path.join(_SCRIPT_DIR, "resolve_whisper_path.txt")
_log(f"Pointer file: {_POINTER} (exists: {os.path.exists(_POINTER)})")

if os.path.exists(_POINTER):
    with open(_POINTER, "r") as f:
        APP_DIR = f.read().strip()
else:
    APP_DIR = _SCRIPT_DIR

_log(f"App dir: {APP_DIR} (exists: {os.path.isdir(APP_DIR)})")

IS_WIN = sys.platform == "win32"

if IS_WIN:
    VENV_PYTHON = os.path.join(APP_DIR, ".venv", "Scripts", "python.exe")
else:
    VENV_PYTHON = os.path.join(APP_DIR, ".venv", "bin", "python3")

CAPTION_SCRIPT = os.path.join(APP_DIR, "caption.py")

_log(f"Venv Python: {VENV_PYTHON} (exists: {os.path.exists(VENV_PYTHON)})")
_log(f"Caption script: {CAPTION_SCRIPT} (exists: {os.path.exists(CAPTION_SCRIPT)})")

# --- Brand ---
_BRAND = "LAB37 TOOLS"
_TOOL = "Resolve Whisper"
_TITLE = f"{_BRAND}: {_TOOL}"
_GREEN = "#42FC93"
_DARK = "#0D0D0D"


# --- Styled UI helpers ---

def _make_dialog(title, brand_label, brand_color, message, btn_text, btn_fg, btn_bg):
    """Create a branded dialog that auto-sizes to fit content."""
    import tkinter as tk

    root = tk.Tk()
    root.title(title)
    root.configure(bg=_DARK)
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frame = tk.Frame(root, bg=_DARK, padx=28, pady=20)
    frame.pack(fill="both", expand=True)

    tk.Label(
        frame, text=brand_label, font=("Consolas", 9), fg=brand_color, bg=_DARK
    ).pack(anchor="w")

    tk.Label(
        frame, text=message, font=("Segoe UI", 10), fg="#FFFFFF", bg=_DARK,
        wraplength=480, justify="left"
    ).pack(pady=(12, 20), anchor="w")

    tk.Button(
        frame, text=btn_text, font=("Consolas", 10, "bold"),
        fg=btn_fg, bg=btn_bg, activebackground=btn_bg,
        relief="flat", padx=20, pady=4,
        command=root.destroy
    ).pack(anchor="e")

    # Auto-size: let tkinter measure content, then center on screen
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 16, 500)
    h = root.winfo_reqheight() + 8
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    return root


def show_message(title, message):
    _log(f"show_message: {title} - {message}")
    try:
        root = _make_dialog(
            title, _BRAND, _GREEN, message,
            "OK", _DARK, _GREEN,
        )
        root.mainloop()
    except Exception as e:
        _log(f"tkinter failed: {e}")
        print(f"{title}: {message}")


def show_error(title, message):
    _log(f"show_error: {title} - {message}")
    try:
        root = _make_dialog(
            title, f"{_BRAND}  //  ERROR", "#FF3333", message,
            "OK", "#FFFFFF", "#333333",
        )
        root.mainloop()
    except Exception as e:
        _log(f"tkinter failed: {e}")
        print(f"ERROR - {title}: {message}")


def show_progress(title, message):
    _log(f"show_progress: {title} - {message}")
    try:
        import tkinter as tk

        root = tk.Tk()
        root.title(title)
        root.configure(bg=_DARK)
        root.resizable(False, False)
        root.attributes("-topmost", True)

        root.update_idletasks()
        w, h = 500, 170
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")

        frame = tk.Frame(root, bg=_DARK, padx=28, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame, text=_BRAND, font=("Consolas", 9), fg=_GREEN, bg=_DARK
        ).pack(anchor="w")

        label = tk.Label(
            frame, text=message, font=("Segoe UI", 10), fg="#FFFFFF", bg=_DARK,
            wraplength=460, justify="left"
        )
        label.pack(pady=(12, 8), anchor="w")

        # Progress bar (green fill on dark track)
        bar_track = tk.Frame(frame, bg="#1A1C1E", height=8)
        bar_track.pack(fill="x", pady=(0, 8))
        bar_track.pack_propagate(False)

        bar_fill = tk.Frame(bar_track, bg=_GREEN, height=8, width=0)
        bar_fill.place(x=0, y=0, relheight=1.0, width=0)

        pct_label = tk.Label(
            frame, text="", font=("Consolas", 9), fg="#888888", bg=_DARK
        )
        pct_label.pack(anchor="w")

        def update_text(new_message):
            _log(f"progress update: {new_message}")
            try:
                label.config(text=new_message)
                root.update()
            except Exception:
                pass

        def update_bar(percent):
            try:
                track_w = bar_track.winfo_width()
                fill_w = max(1, int(track_w * percent / 100))
                bar_fill.place(x=0, y=0, relheight=1.0, width=fill_w)
                pct_label.config(text=f"{percent}%")
                root.update()
            except Exception:
                pass

        def close():
            try:
                root.destroy()
            except Exception:
                pass

        root.protocol("WM_DELETE_WINDOW", close)
        root.update()
        _log("Progress window created OK")

        return close, update_text, update_bar, root

    except Exception as e:
        _log(f"show_progress failed: {e}")
        print(f"{title}: {message}")
        return (lambda: None), (lambda m: print(m)), (lambda p: None), None


def open_file_selected(file_path):
    """Open Explorer with the file pre-selected (Windows)."""
    try:
        if IS_WIN:
            subprocess.Popen(
                ["explorer", "/select,", file_path],
                creationflags=0x08000000,
            )
        else:
            subprocess.Popen(["open", "-R", file_path])
        _log(f"Opened explorer with file selected: {file_path}")
    except Exception as e:
        _log(f"open_file_selected failed: {e}")


def show_settings():
    """Show settings dialog before transcription. Returns dict or None if cancelled."""
    _log("Showing settings dialog")
    try:
        import tkinter as tk

        result = [None]

        root = tk.Tk()
        root.title(_TITLE)
        root.configure(bg=_DARK)
        root.resizable(False, False)
        root.attributes("-topmost", True)

        frame = tk.Frame(root, bg=_DARK, padx=28, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame, text=_BRAND, font=("Consolas", 9), fg=_GREEN, bg=_DARK
        ).pack(anchor="w")

        tk.Label(
            frame, text="Caption Settings", font=("Segoe UI", 12, "bold"),
            fg="#FFFFFF", bg=_DARK
        ).pack(pady=(8, 16), anchor="w")

        # --- Max words per caption ---
        row1 = tk.Frame(frame, bg=_DARK)
        row1.pack(fill="x", pady=(0, 10))

        tk.Label(
            row1, text="Words per caption:", font=("Segoe UI", 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        words_var = tk.StringVar(value="1")
        words_options = ["1", "2", "3", "4", "6", "Auto"]
        words_menu = tk.OptionMenu(row1, words_var, *words_options)
        words_menu.config(
            font=("Consolas", 10), fg="#FFFFFF", bg="#1A1C1E",
            activebackground="#333333", highlightthickness=0, relief="flat", width=8
        )
        words_menu["menu"].config(
            font=("Consolas", 10), fg="#FFFFFF", bg="#1A1C1E",
            activebackground=_GREEN, activeforeground=_DARK
        )
        words_menu.pack(side="left")

        # --- Max chars per line ---
        row2 = tk.Frame(frame, bg=_DARK)
        row2.pack(fill="x", pady=(0, 20))

        tk.Label(
            row2, text="Characters per line:", font=("Segoe UI", 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        chars_var = tk.StringVar(value="42")
        chars_options = ["32", "37", "42", "47", "52"]
        chars_menu = tk.OptionMenu(row2, chars_var, *chars_options)
        chars_menu.config(
            font=("Consolas", 10), fg="#FFFFFF", bg="#1A1C1E",
            activebackground="#333333", highlightthickness=0, relief="flat", width=8
        )
        chars_menu["menu"].config(
            font=("Consolas", 10), fg="#FFFFFF", bg="#1A1C1E",
            activebackground=_GREEN, activeforeground=_DARK
        )
        chars_menu.pack(side="left")

        # --- Buttons ---
        btn_row = tk.Frame(frame, bg=_DARK)
        btn_row.pack(fill="x")

        def on_cancel():
            result[0] = None
            root.destroy()

        def on_start():
            w = words_var.get()
            result[0] = {
                "max_words": 0 if w == "Auto" else int(w),
                "max_chars": int(chars_var.get()),
            }
            root.destroy()

        tk.Button(
            btn_row, text="Cancel", font=("Consolas", 10),
            fg="#AAAAAA", bg="#1A1C1E", activebackground="#333333",
            relief="flat", padx=16, pady=4,
            command=on_cancel
        ).pack(side="left")

        tk.Button(
            btn_row, text="Generate Captions", font=("Consolas", 10, "bold"),
            fg=_DARK, bg=_GREEN, activebackground="#35D07A",
            relief="flat", padx=16, pady=4,
            command=on_start
        ).pack(side="right")

        # Center on screen
        root.update_idletasks()
        w = max(root.winfo_reqwidth() + 16, 420)
        h = root.winfo_reqheight() + 8
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")

        root.mainloop()

        _log(f"Settings result: {result[0]}")
        return result[0]

    except Exception as e:
        _log(f"Settings dialog failed: {e}")
        # Default settings on failure
        return {"max_words": 0, "max_chars": 42}


def _import_srt(project, timeline, srt_path):
    """Try to import SRT onto subtitle track. Returns True if successful."""

    # Resolve's scripting API (through v20.2) has no method to import
    # SRT files onto the subtitle track. We try ImportIntoTimeline as a
    # forward-compat attempt, then fall back to Explorer file selection.

    _log(f"Attempting SRT import: {srt_path}")

    # Ensure subtitle track exists
    sub_count = timeline.GetTrackCount("subtitle")
    _log(f"Existing subtitle tracks: {sub_count}")
    if sub_count == 0:
        try:
            timeline.AddTrack("subtitle")
            _log("Created subtitle track")
        except Exception as e:
            _log(f"AddTrack('subtitle') failed: {e}")

    # Try ImportIntoTimeline (may work in future Resolve versions)
    try:
        result = timeline.ImportIntoTimeline(srt_path)
        _log(f"ImportIntoTimeline: {result}")
        if result:
            return True
    except Exception as e:
        _log(f"ImportIntoTimeline failed: {e}")

    _log("Auto-import not available in this Resolve version")
    return False


# --- Main ---

def main():
    _log("main() entered")

    # Verify install
    if not os.path.exists(VENV_PYTHON):
        show_error(
            _TITLE,
            f"Python environment not found.\n\n"
            f"Run setup.bat (Windows) or setup.sh (Mac) first.\n\n"
            f"Expected: {VENV_PYTHON}"
        )
        return

    if not os.path.exists(CAPTION_SCRIPT):
        show_error(
            _TITLE,
            f"Caption script not found at:\n{CAPTION_SCRIPT}\n\n"
            f"Re-run setup."
        )
        return

    _log("Install verified OK")

    # Connect to Resolve
    _log("Connecting to Resolve...")
    resolve = None

    # Method 1: Resolve injects 'bmd' global into fuscript environment
    try:
        resolve = bmd.scriptapp("Resolve")  # noqa: F821 - injected by Resolve
        _log(f"Connected via bmd global: {resolve}")
    except Exception as e:
        _log(f"bmd global not available: {e}")

    # Method 2: Try fusion global (available in Fusion page scripts)
    if not resolve:
        try:
            resolve = fusion.GetResolve()  # noqa: F821 - injected by Resolve
            _log(f"Connected via fusion.GetResolve(): {resolve}")
        except Exception as e:
            _log(f"fusion global not available: {e}")

    # Method 3: Import DaVinciResolveScript with path fix
    if not resolve:
        try:
            _modules_paths = [
                os.path.join(os.environ.get("PROGRAMDATA", "C:\\ProgramData"),
                             "Blackmagic Design", "DaVinci Resolve",
                             "Support", "Developer", "Scripting", "Modules"),
            ]
            for p in _modules_paths:
                if os.path.isdir(p) and p not in sys.path:
                    sys.path.insert(0, p)
                    _log(f"Added to sys.path: {p}")

            import DaVinciResolveScript as dvr
            resolve = dvr.scriptapp("Resolve")
            _log(f"Connected via DaVinciResolveScript: {resolve}")
        except Exception as e:
            _log(f"DaVinciResolveScript import failed: {e}")

    if not resolve:
        show_error(
            _TITLE,
            "Could not connect to Resolve.\n\n"
            "Make sure you're running this from inside Resolve\n"
            "(Workspace > Scripts > LAB37 Resolve Whisper)."
        )
        return

    pm = resolve.GetProjectManager()
    project = pm.GetCurrentProject()
    if not project:
        show_error(_TITLE, "No project is open.")
        return

    timeline = project.GetCurrentTimeline()
    if not timeline:
        show_error(_TITLE, "No timeline selected.\n\nSelect a timeline first.")
        return

    timeline_name = timeline.GetName()
    _log(f"Timeline: {timeline_name}")

    fps_str = timeline.GetSetting("timelineFrameRate")
    try:
        fps = float(fps_str)
    except (ValueError, TypeError):
        fps = 24.0

    start_frame = timeline.GetStartFrame()
    end_frame = timeline.GetEndFrame()
    duration_s = (end_frame - start_frame) / fps
    _log(f"Timeline: frame {start_frame}-{end_frame}, {duration_s:.1f}s, FPS: {fps}")

    # Show settings dialog
    settings = show_settings()
    if settings is None:
        _log("User cancelled settings dialog")
        return

    max_words = settings["max_words"]
    max_chars = settings["max_chars"]
    _log(f"Settings: max_words={max_words}, max_chars={max_chars}")

    # Show progress
    close_progress, update_status, update_bar, progress_root = show_progress(
        _TITLE,
        f"Rendering audio from '{timeline_name}'"
    )

    try:
        # Render audio for current in/out range
        # We do NOT set MarkIn/MarkOut - Resolve automatically uses
        # the timeline's in/out points (set with I and O keys).
        # This matches how the built-in transcription tool works.
        tmp_dir = tempfile.mkdtemp(prefix="resolve_whisper_")
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
        wav_name = f"{safe_name}_audio"

        _log(f"Rendering audio to {tmp_dir}")

        # Force WAV output format (Resolve defaults to MOV otherwise)
        if not project.SetCurrentRenderFormatAndCodec("wav", "LinearPCM"):
            _log("SetCurrentRenderFormatAndCodec('wav','LinearPCM') failed, trying 'Wave'")
            project.SetCurrentRenderFormatAndCodec("Wave", "LinearPCM")

        project.SetRenderSettings({
            "ExportAudio": True,
            "ExportVideo": False,
            "TargetDir": tmp_dir,
            "CustomName": wav_name,
            "AudioBitDepth": 16,
            "AudioSampleRate": 16000,
        })
        _log("Render settings applied (using timeline in/out points)")

        job_id = project.AddRenderJob()
        _log(f"Render job ID: {job_id}")

        if not job_id:
            close_progress()
            show_error(_TITLE, "Failed to create render job.\n\nCheck your render settings.")
            return

        # Render ONLY our job (not the entire queue)
        project.StartRendering([job_id])
        _log("Rendering started (our job only)")

        # Wait for render
        while project.IsRenderingInProgress():
            if progress_root:
                try:
                    progress_root.update()
                except Exception:
                    pass
            time.sleep(0.3)

        _log("Rendering finished")

        # Remove our job from the render queue so we don't clutter it
        try:
            project.DeleteRenderJob(job_id)
            _log("Render job cleaned up from queue")
        except Exception:
            pass

        # Switch back to Edit page (rendering puts us on Deliver)
        try:
            resolve.OpenPage("edit")
            _log("Switched back to Edit page")
        except Exception:
            pass

        # Find rendered audio file (WAV preferred, but accept any format)
        audio_path = None
        try:
            files = os.listdir(tmp_dir)
            _log(f"Files in tmp_dir: {files}")
            # Prefer .wav, fall back to any file (mov/mp4/flac all work with whisper)
            for ext in [".wav", ".mov", ".mp4", ".flac", ".mp3", ".aac", ".m4a"]:
                for f in files:
                    if f.lower().endswith(ext):
                        audio_path = os.path.join(tmp_dir, f)
                        break
                if audio_path:
                    break
            # Last resort: take any file
            if not audio_path and files:
                audio_path = os.path.join(tmp_dir, files[0])
        except Exception as e:
            _log(f"Error listing tmp_dir: {e}")

        if not audio_path:
            close_progress()
            show_error(_TITLE, "Audio render finished but no file was created.")
            return

        _log(f"Audio file: {audio_path} ({os.path.getsize(audio_path)} bytes)")

        # Transcribe
        update_status("Transcribing with AI (this may take a moment)")

        output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Captions")
        os.makedirs(output_dir, exist_ok=True)
        srt_path = os.path.join(output_dir, f"{safe_name}.srt")

        cmd = [
            VENV_PYTHON,
            CAPTION_SCRIPT,
            "--file", audio_path,
            "--output", srt_path,
            "--fps", str(fps),
            "--max-words", str(max_words),
            "--max-chars", str(max_chars),
        ]

        _log(f"Running: {' '.join(cmd)}")

        # Use Popen to read progress lines in real-time
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=APP_DIR,
        )

        stderr_lines = []
        # Read stdout for PROGRESS:XX lines, keep UI responsive
        while True:
            line = proc.stdout.readline()
            if not line and proc.poll() is not None:
                break
            if line:
                line = line.strip()
                if line.startswith("PROGRESS:"):
                    try:
                        pct = int(line.split(":")[1])
                        update_bar(pct)
                        update_status(f"Transcribing... {pct}%")
                    except (ValueError, IndexError):
                        pass
                else:
                    _log(f"stdout: {line}")

            # Keep tkinter responsive
            if progress_root:
                try:
                    progress_root.update()
                except Exception:
                    pass

        # Collect remaining stderr
        remaining_stderr = proc.stderr.read()
        if remaining_stderr:
            stderr_lines.append(remaining_stderr)

        returncode = proc.wait()
        stderr_text = "".join(stderr_lines)

        _log(f"Return code: {returncode}")
        if stderr_text:
            _log(f"stderr: {stderr_text[-500:]}")

        # Clean up temp audio
        try:
            os.unlink(audio_path)
            os.rmdir(tmp_dir)
        except Exception:
            pass

        close_progress()

        if returncode != 0:
            error_msg = stderr_text.strip().split("\n")[-1] if stderr_text else "Unknown error"
            show_error(
                _TITLE,
                f"{error_msg}\n\nCheck that setup completed successfully."
            )
            return

        if not os.path.exists(srt_path):
            show_error(_TITLE, "Transcription finished but SRT file was not created.")
            return

        # Count captions
        with open(srt_path, "r", encoding="utf-8") as f:
            content = f.read()
        caption_count = content.count("\n\n")

        _log(f"Done! {caption_count} captions written to {srt_path}")

        # Try auto-import SRT into timeline
        update_status("Importing captions into timeline")
        imported = _import_srt(project, timeline, srt_path)

        close_progress()

        if imported:
            show_message(
                _TITLE,
                f"Done! {caption_count} captions added to '{timeline_name}'.\n\n"
                f"SRT also saved to:\n{srt_path}"
            )
        else:
            # Resolve's scripting API doesn't support SRT import.
            # Open Explorer with the SRT file selected for easy drag or navigation.
            open_file_selected(srt_path)
            show_message(
                _TITLE,
                f"Done! {caption_count} captions for '{timeline_name}'.\n\n"
                f"To import into your timeline:\n"
                f"File > Import > Subtitle\n"
                f"(the file is highlighted in the Explorer window)\n\n"
                f"{srt_path}"
            )

    except subprocess.TimeoutExpired:
        close_progress()
        show_error(_TITLE, "Transcription timed out (>10 minutes).")
    except Exception as e:
        _log(f"Exception in main: {e}\n{traceback.format_exc()}")
        close_progress()
        show_error(_TITLE, f"Unexpected error:\n{e}")


try:
    main()
except Exception as e:
    _log(f"FATAL: {e}\n{traceback.format_exc()}")
    try:
        show_error(_TITLE, f"Fatal error:\n{e}")
    except Exception:
        pass
