"""
Resolve Whisper - Transcribe current timeline with AI.
This script runs inside DaVinci Resolve (Workspace > Scripts > Resolve Whisper).
It renders audio from the timeline, then calls the Whisper engine to generate captions.
"""
import os
import sys
import subprocess
import tempfile
import time
import threading

# --- Locate our install directory ---
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_POINTER = os.path.join(_SCRIPT_DIR, "resolve_whisper_path.txt")

if os.path.exists(_POINTER):
    with open(_POINTER, "r") as f:
        APP_DIR = f.read().strip()
else:
    # Fallback: assume script is in the app dir itself
    APP_DIR = _SCRIPT_DIR

IS_WIN = sys.platform == "win32"

if IS_WIN:
    VENV_PYTHON = os.path.join(APP_DIR, ".venv", "Scripts", "python.exe")
else:
    VENV_PYTHON = os.path.join(APP_DIR, ".venv", "bin", "python3")

CAPTION_SCRIPT = os.path.join(APP_DIR, "caption.py")


# --- Simple UI helpers (tkinter is available in Resolve's Python) ---

def show_message(title, message):
    """Show a simple message dialog."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(title, message)
        root.destroy()
    except Exception:
        print(f"{title}: {message}")


def show_error(title, message):
    """Show an error dialog."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:
        print(f"ERROR - {title}: {message}")


def show_progress(title, message):
    """Show a non-blocking progress window. Returns a close function."""
    try:
        import tkinter as tk

        root = tk.Tk()
        root.title(title)
        root.geometry("400x120")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        # Center on screen
        root.update_idletasks()
        x = (root.winfo_screenwidth() - 400) // 2
        y = (root.winfo_screenheight() - 120) // 2
        root.geometry(f"+{x}+{y}")

        frame = tk.Frame(root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        label = tk.Label(frame, text=message, font=("Arial", 11), wraplength=360)
        label.pack(pady=(0, 10))

        dots_var = tk.StringVar(value="")
        dots_label = tk.Label(frame, textvariable=dots_var, font=("Arial", 11))
        dots_label.pack()

        # Animate dots
        _running = [True]
        _count = [0]

        def animate():
            if _running[0]:
                _count[0] = (_count[0] % 3) + 1
                dots_var.set("." * _count[0])
                root.after(500, animate)

        animate()

        def update_text(new_message):
            try:
                label.config(text=new_message)
                root.update()
            except Exception:
                pass

        def close():
            _running[0] = False
            try:
                root.destroy()
            except Exception:
                pass

        root.protocol("WM_DELETE_WINDOW", close)
        root.update()

        return close, update_text, root

    except Exception:
        print(f"{title}: {message}")
        return (lambda: None), (lambda m: print(m)), None


def open_file(path):
    """Open a file or folder in the system default handler."""
    try:
        if IS_WIN:
            os.startfile(path)
        else:
            subprocess.Popen(["open", path])
    except Exception:
        pass


# --- Main ---

def main():
    # Verify install
    if not os.path.exists(VENV_PYTHON):
        show_error(
            "Resolve Whisper - Not Installed",
            f"Python environment not found.\n\n"
            f"Run setup.bat (Windows) or setup.sh (Mac) first.\n\n"
            f"Expected: {VENV_PYTHON}"
        )
        return

    if not os.path.exists(CAPTION_SCRIPT):
        show_error(
            "Resolve Whisper - Not Installed",
            f"Caption script not found at:\n{CAPTION_SCRIPT}\n\n"
            f"Re-run setup."
        )
        return

    # Connect to Resolve
    try:
        import DaVinciResolveScript as bmd
        resolve = bmd.scriptapp("Resolve")
    except Exception:
        show_error("Resolve Whisper", "Could not connect to Resolve.")
        return

    if not resolve:
        show_error("Resolve Whisper", "Could not connect to Resolve.")
        return

    pm = resolve.GetProjectManager()
    project = pm.GetCurrentProject()
    if not project:
        show_error("Resolve Whisper", "No project is open.")
        return

    timeline = project.GetCurrentTimeline()
    if not timeline:
        show_error("Resolve Whisper", "No timeline selected.\n\nSelect a timeline first.")
        return

    timeline_name = timeline.GetName()
    fps_str = timeline.GetSetting("timelineFrameRate")
    try:
        fps = float(fps_str)
    except (ValueError, TypeError):
        fps = 24.0

    start_frame = timeline.GetStartFrame()
    end_frame = timeline.GetEndFrame()
    duration_s = (end_frame - start_frame) / fps

    # Show progress
    close_progress, update_status, progress_root = show_progress(
        "Resolve Whisper",
        f"Rendering audio from '{timeline_name}' ({duration_s:.0f}s)"
    )

    try:
        # Render audio
        tmp_dir = tempfile.mkdtemp(prefix="resolve_whisper_")
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
        wav_name = f"{safe_name}_audio"

        project.SetRenderSettings({
            "ExportAudio": True,
            "ExportVideo": False,
            "TargetDir": tmp_dir,
            "CustomName": wav_name,
            "AudioCodec": "LinearPCM",
            "AudioBitDepth": 16,
            "AudioSampleRate": 16000,
            "FormatWidth": 0,
            "FormatHeight": 0,
        })

        job_id = project.AddRenderJob()
        if not job_id:
            close_progress()
            show_error("Resolve Whisper", "Failed to create render job.\n\nCheck your render settings.")
            return

        project.StartRendering()

        # Wait for render
        while project.IsRenderingInProgress():
            if progress_root:
                try:
                    progress_root.update()
                except Exception:
                    pass
            time.sleep(0.3)

        # Find rendered WAV
        wav_path = None
        for f in os.listdir(tmp_dir):
            if f.endswith(".wav"):
                wav_path = os.path.join(tmp_dir, f)
                break

        if not wav_path:
            close_progress()
            show_error("Resolve Whisper", "Audio render finished but no WAV file found.")
            return

        # Transcribe using our venv Python
        update_status("Transcribing with AI (this may take a moment)")

        output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Captions")
        os.makedirs(output_dir, exist_ok=True)
        srt_path = os.path.join(output_dir, f"{safe_name}.srt")

        cmd = [
            VENV_PYTHON,
            CAPTION_SCRIPT,
            "--file", wav_path,
            "--output", srt_path,
            "--fps", str(fps),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=APP_DIR,
            timeout=600,  # 10 min max
        )

        # Clean up temp audio
        try:
            os.unlink(wav_path)
            os.rmdir(tmp_dir)
        except Exception:
            pass

        close_progress()

        if result.returncode != 0:
            error_msg = result.stderr.strip().split("\n")[-1] if result.stderr else "Unknown error"
            show_error(
                "Resolve Whisper - Transcription Failed",
                f"{error_msg}\n\nCheck that setup completed successfully."
            )
            return

        if not os.path.exists(srt_path):
            show_error("Resolve Whisper", "Transcription finished but SRT file was not created.")
            return

        # Count captions
        with open(srt_path, "r", encoding="utf-8") as f:
            content = f.read()
        caption_count = content.count("\n\n")

        show_message(
            "Resolve Whisper - Done",
            f"Generated {caption_count} captions for '{timeline_name}'.\n\n"
            f"Saved to:\n{srt_path}\n\n"
            f"To import: File > Import > Subtitle"
        )

        # Open the captions folder
        open_file(output_dir)

    except subprocess.TimeoutExpired:
        close_progress()
        show_error("Resolve Whisper", "Transcription timed out (>10 minutes).")
    except Exception as e:
        close_progress()
        show_error("Resolve Whisper", f"Unexpected error:\n{e}")


main()
