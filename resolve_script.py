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
    import shutil
    import subprocess
    import tempfile
    import threading
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
    with open(_POINTER, "r", encoding="utf-8") as f:
        APP_DIR = f.read().strip()
else:
    APP_DIR = _SCRIPT_DIR

_log(f"App dir: {APP_DIR} (exists: {os.path.isdir(APP_DIR)})")

IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")

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

# Platform-appropriate fonts (Consolas/Segoe UI are Windows-only)
if IS_WIN:
    _FONT_MONO = "Consolas"
    _FONT_UI = "Segoe UI"
elif IS_MAC:
    _FONT_MONO = "Menlo"
    _FONT_UI = "Helvetica Neue"
else:
    _FONT_MONO = "Monospace"
    _FONT_UI = "Sans"


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
        frame, text=brand_label, font=(_FONT_MONO, 9), fg=brand_color, bg=_DARK
    ).pack(anchor="w")

    tk.Label(
        frame, text=message, font=(_FONT_UI, 10), fg="#FFFFFF", bg=_DARK,
        wraplength=480, justify="left"
    ).pack(pady=(12, 20), anchor="w")

    def _dismiss():
        root.quit()
        root.destroy()

    tk.Button(
        frame, text=btn_text, font=(_FONT_MONO, 10, "bold"),
        fg=btn_fg, bg=btn_bg, activebackground=btn_bg,
        relief="flat", padx=20, pady=4,
        command=_dismiss
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
        w, h = 500, 190
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")

        frame = tk.Frame(root, bg=_DARK, padx=28, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame, text=_BRAND, font=(_FONT_MONO, 9), fg=_GREEN, bg=_DARK
        ).pack(anchor="w")

        # Phase indicator (e.g. "Step 1 of 3")
        phase_label = tk.Label(
            frame, text="", font=(_FONT_MONO, 8), fg="#555555", bg=_DARK
        )
        phase_label.pack(anchor="w", pady=(8, 0))

        label = tk.Label(
            frame, text=message, font=(_FONT_UI, 10), fg="#FFFFFF", bg=_DARK,
            wraplength=460, justify="left"
        )
        label.pack(pady=(2, 8), anchor="w")

        # Progress bar
        bar_track = tk.Frame(frame, bg="#1A1C1E", height=8)
        bar_track.pack(fill="x", pady=(0, 8))
        bar_track.pack_propagate(False)

        bar_fill = tk.Frame(bar_track, bg=_GREEN, height=8, width=0)
        bar_fill.place(x=0, y=0, relheight=1.0, width=0)

        # Bottom row: percentage left, elapsed right
        bottom_row = tk.Frame(frame, bg=_DARK)
        bottom_row.pack(fill="x")

        pct_label = tk.Label(
            bottom_row, text="", font=(_FONT_MONO, 9), fg="#888888", bg=_DARK
        )
        pct_label.pack(side="left")

        elapsed_label = tk.Label(
            bottom_row, text="", font=(_FONT_MONO, 9), fg="#555555", bg=_DARK
        )
        elapsed_label.pack(side="right")

        # Elapsed time ticker
        _elapsed = [0]
        _timer_id = [None]

        def _tick():
            _elapsed[0] += 1
            mins = _elapsed[0] // 60
            secs = _elapsed[0] % 60
            try:
                elapsed_label.config(text=f"{mins}:{secs:02d}")
            except Exception:
                return
            _timer_id[0] = root.after(1000, _tick)

        _timer_id[0] = root.after(1000, _tick)

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

        def update_phase(step, total, name):
            try:
                phase_label.config(text=f"Step {step} of {total}")
                label.config(text=name)
                bar_fill.place(x=0, y=0, relheight=1.0, width=0)
                pct_label.config(text="")
                root.update()
            except Exception:
                pass

        def close():
            try:
                if _timer_id[0]:
                    root.after_cancel(_timer_id[0])
                root.quit()
                root.destroy()
            except Exception:
                pass

        root.protocol("WM_DELETE_WINDOW", close)
        root.update()
        _log("Progress window created OK")

        return close, update_text, update_bar, update_phase, root

    except Exception as e:
        _log(f"show_progress failed: {e}")
        print(f"{title}: {message}")
        return (lambda: None), (lambda m: print(m)), (lambda p: None), (lambda s, t, n: None), None


def open_file_selected(file_path):
    """Open file manager with the file pre-selected."""
    try:
        if IS_WIN:
            subprocess.Popen(
                ["explorer", "/select,", file_path],
                creationflags=0x08000000,
            )
        elif IS_MAC:
            subprocess.Popen(["open", "-R", file_path])
        else:
            # Linux: open the containing directory
            subprocess.Popen(["xdg-open", os.path.dirname(file_path)])
        _log(f"Opened explorer with file selected: {file_path}")
    except Exception as e:
        _log(f"open_file_selected failed: {e}")


_PRESETS = {
    "Podcast 16:9": {"max_words": 0, "max_chars": 42, "max_lines": 1},
    # Kept aligned with the Mac preset (presets/LAB37 Reels.lua)
    "Reels":        {"max_words": 3, "max_chars": 22, "max_lines": 1},
    "Custom":       None,
}
_PRESET_NAMES = list(_PRESETS.keys())


def show_settings(timeline_name="", duration_s=0.0, fps=24.0):
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
            frame, text=_BRAND, font=(_FONT_MONO, 9), fg=_GREEN, bg=_DARK
        ).pack(anchor="w")

        tk.Label(
            frame, text="Caption Settings", font=(_FONT_UI, 12, "bold"),
            fg="#FFFFFF", bg=_DARK
        ).pack(pady=(8, 4), anchor="w")

        # --- Timeline context ---
        if timeline_name:
            mins = int(duration_s) // 60
            secs = int(duration_s) % 60
            ctx = f"{timeline_name}  //  {mins}m {secs}s  //  {fps:.0f} fps"
            tk.Label(
                frame, text=ctx, font=(_FONT_MONO, 8),
                fg="#666666", bg=_DARK
            ).pack(anchor="w", pady=(0, 12))

        _menu_cfg = dict(
            font=(_FONT_MONO, 10), fg="#FFFFFF", bg="#1A1C1E",
            activebackground="#333333", highlightthickness=0, relief="flat",
        )
        _dropdown_cfg = dict(
            font=(_FONT_MONO, 10), fg="#FFFFFF", bg="#1A1C1E",
            activebackground=_GREEN, activeforeground=_DARK,
        )

        # --- Preset selector ---
        row_preset = tk.Frame(frame, bg=_DARK)
        row_preset.pack(fill="x", pady=(0, 10))

        tk.Label(
            row_preset, text="Preset:", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        preset_var = tk.StringVar(value=_PRESET_NAMES[0])
        preset_menu = tk.OptionMenu(row_preset, preset_var, *_PRESET_NAMES)
        preset_menu.config(width=14, **_menu_cfg)
        preset_menu["menu"].config(**_dropdown_cfg)
        preset_menu.pack(side="left")

        # --- Output mode ---
        row_output = tk.Frame(frame, bg=_DARK)
        row_output.pack(fill="x", pady=(0, 10))

        tk.Label(
            row_output, text="Output:", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        _OUTPUT_MODES = ["Text+ (native)", "SRT file"]
        output_var = tk.StringVar(value=_OUTPUT_MODES[0])
        output_menu = tk.OptionMenu(row_output, output_var, *_OUTPUT_MODES)
        output_menu.config(width=14, **_menu_cfg)
        output_menu["menu"].config(**_dropdown_cfg)
        output_menu.pack(side="left")

        # --- Language selector ---
        row_lang = tk.Frame(frame, bg=_DARK)
        row_lang.pack(fill="x", pady=(0, 14))

        tk.Label(
            row_lang, text="Language:", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        _LANGUAGES = {"Svenska": "sv", "English": "en", "Auto-detect": None}
        lang_var = tk.StringVar(value="Svenska")
        lang_menu = tk.OptionMenu(row_lang, lang_var, *_LANGUAGES.keys())
        lang_menu.config(width=14, **_menu_cfg)
        lang_menu["menu"].config(**_dropdown_cfg)
        lang_menu.pack(side="left")

        # --- Separator ---
        tk.Frame(frame, bg="#333333", height=1).pack(fill="x", pady=(0, 14))

        # --- Max words per caption ---
        row1 = tk.Frame(frame, bg=_DARK)
        row1.pack(fill="x", pady=(0, 10))

        tk.Label(
            row1, text="Words per caption:", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        words_var = tk.StringVar(value="Auto")
        words_options = ["1", "2", "3", "4", "6", "Auto"]
        words_menu = tk.OptionMenu(row1, words_var, *words_options)
        words_menu.config(width=8, **_menu_cfg)
        words_menu["menu"].config(**_dropdown_cfg)
        words_menu.pack(side="left")

        # --- Max chars per line ---
        row2 = tk.Frame(frame, bg=_DARK)
        row2.pack(fill="x", pady=(0, 10))

        tk.Label(
            row2, text="Characters per line:", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        chars_var = tk.StringVar(value="42")
        chars_options = ["32", "37", "42", "47", "52"]
        chars_menu = tk.OptionMenu(row2, chars_var, *chars_options)
        chars_menu.config(width=8, **_menu_cfg)
        chars_menu["menu"].config(**_dropdown_cfg)
        chars_menu.pack(side="left")

        # --- Max lines per caption ---
        row3 = tk.Frame(frame, bg=_DARK)
        row3.pack(fill="x", pady=(0, 10))

        tk.Label(
            row3, text="Lines per caption:", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, width=20, anchor="w"
        ).pack(side="left")

        lines_var = tk.StringVar(value="1")
        lines_options = ["1", "2"]
        lines_menu = tk.OptionMenu(row3, lines_var, *lines_options)
        lines_menu.config(width=8, **_menu_cfg)
        lines_menu["menu"].config(**_dropdown_cfg)
        lines_menu.pack(side="left")

        # --- Remove punctuation ---
        row_punct = tk.Frame(frame, bg=_DARK)
        row_punct.pack(fill="x", pady=(0, 20))

        strip_punct_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            row_punct, text="Remove punctuation", font=(_FONT_UI, 10),
            fg="#CCCCCC", bg=_DARK, selectcolor="#1A1C1E",
            activebackground=_DARK, activeforeground="#CCCCCC",
            highlightthickness=0, variable=strip_punct_var,
        ).pack(anchor="w")

        # --- Preset auto-fill logic ---
        _detail_menus = [words_menu, chars_menu, lines_menu]

        def _apply_preset(*_args):
            name = preset_var.get()
            vals = _PRESETS.get(name)
            if vals:
                words_var.set("Auto" if vals["max_words"] == 0 else str(vals["max_words"]))
                chars_var.set(str(vals["max_chars"]))
                lines_var.set(str(vals["max_lines"]))
                for m in _detail_menus:
                    m.config(state="disabled")
            else:
                for m in _detail_menus:
                    m.config(state="normal")

        preset_var.trace_add("write", _apply_preset)
        _apply_preset()

        # --- Dynamic button label ---
        start_btn = None

        def _update_btn_label(*_args):
            if start_btn:
                lbl = "Generate Captions" if output_var.get().startswith("Text+") else "Export SRT"
                start_btn.config(text=lbl)

        # --- Buttons ---
        btn_row = tk.Frame(frame, bg=_DARK)
        btn_row.pack(fill="x")

        def on_cancel():
            result[0] = None
            root.quit()
            root.destroy()

        def on_start():
            w = words_var.get()
            result[0] = {
                "max_words": 0 if w == "Auto" else int(w),
                "max_chars": int(chars_var.get()),
                "max_lines": int(lines_var.get()),
                "language": _LANGUAGES[lang_var.get()],
                "strip_punctuation": strip_punct_var.get(),
                "output_mode": "textplus" if output_var.get().startswith("Text+") else "srt",
            }
            root.quit()
            root.destroy()

        tk.Button(
            btn_row, text="Cancel", font=(_FONT_MONO, 10),
            fg="#AAAAAA", bg="#1A1C1E", activebackground="#333333",
            relief="flat", padx=16, pady=4,
            command=on_cancel
        ).pack(side="left")

        start_btn = tk.Button(
            btn_row, text="Generate Captions", font=(_FONT_MONO, 10, "bold"),
            fg=_DARK, bg=_GREEN, activebackground="#35D07A",
            relief="flat", padx=16, pady=4,
            command=on_start
        )
        start_btn.pack(side="right")
        output_var.trace_add("write", _update_btn_label)

        # Center on screen
        root.update_idletasks()
        w = max(root.winfo_reqwidth() + 16, 460)
        h = root.winfo_reqheight() + 8
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")

        root.mainloop()

        _log(f"Settings result: {result[0]}")
        return result[0]

    except Exception as e:
        _log(f"Settings dialog failed: {e}")
        try:
            root.destroy()
        except Exception:
            pass
        return {"max_words": 0, "max_chars": 42, "max_lines": 1, "language": "sv", "strip_punctuation": False, "output_mode": "textplus"}


def _import_srt(_project, timeline, srt_path):
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


# --- Text+ native caption insertion ---

_TEMPLATE_NAME = "_WhisperTemplate"


def _create_transparent_png(path, width=1920, height=1080):
    """Create a transparent PNG at the given resolution (no PIL needed)."""
    import struct
    import zlib

    header = b'\x89PNG\r\n\x1a\n'

    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data) & 0xFFFFFFFF
    ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)

    # All-transparent RGBA: filter byte 0 + (width * 4 zero bytes) per row
    row = b'\x00' + b'\x00' * (width * 4)
    raw = row * height
    compressed = zlib.compress(raw)
    idat_crc = zlib.crc32(b'IDAT' + compressed) & 0xFFFFFFFF
    idat = struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', idat_crc)

    iend_crc = zlib.crc32(b'IEND') & 0xFFFFFFFF
    iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)

    with open(path, 'wb') as f:
        f.write(header + ihdr + idat + iend)


def _find_clip_in_pool(folder, name):
    """Recursively search media pool folders for a clip by name."""
    try:
        for clip in (folder.GetClipList() or []):
            try:
                clip_name = clip.GetClipProperty("Clip Name")
                if clip_name == name:
                    return clip
            except Exception:
                pass
        for sub in (folder.GetSubFolderList() or []):
            result = _find_clip_in_pool(sub, name)
            if result:
                return result
    except Exception as e:
        _log(f"_find_clip_in_pool error: {e}")
    return None


def _get_or_create_template(media_pool, width=1920, height=1080):
    """Find or create the caption template in the media pool."""
    _log("Looking for caption template in media pool...")

    root_folder = media_pool.GetRootFolder()
    if not root_folder:
        _log("Could not access media pool root folder")
        return None

    png_path = os.path.join(APP_DIR, f"{_TEMPLATE_NAME}.png")

    # Always (re)create PNG at current timeline resolution
    _create_transparent_png(png_path, width, height)
    _log(f"Template PNG ready: {png_path} ({width}x{height})")

    # Delete any existing template so we reimport at the correct resolution
    # (Resolve caches media properties -- overwriting the file isn't enough)
    template = _find_clip_in_pool(root_folder, _TEMPLATE_NAME)
    if template:
        try:
            media_pool.DeleteClips([template])
            _log("Deleted old template from media pool")
        except Exception as e:
            _log(f"Could not delete old template: {e}")

    # Import the fresh PNG as our template
    _log("Importing template...")
    try:

        prev_folder = media_pool.GetCurrentFolder()
        media_pool.SetCurrentFolder(root_folder)

        imported = media_pool.ImportMedia([png_path])
        _log(f"ImportMedia result: {imported}")

        if imported and len(imported) > 0:
            try:
                imported[0].SetClipProperty("Clip Name", _TEMPLATE_NAME)
            except Exception:
                pass
            _log("Template created successfully")
            template = imported[0]
        else:
            _log("ImportMedia returned empty")

        if prev_folder:
            media_pool.SetCurrentFolder(prev_folder)

        return template

    except Exception as e:
        _log(f"Template creation failed: {e}")
        return None


def _insert_textplus_captions(media_pool, timeline, captions, fps, template, on_progress=None):
    """Insert clips and add Text+ nodes via Fusion scripting."""
    _log(f"Inserting {len(captions)} Text+ captions...")

    # Add a new video track -- validate it actually worked
    prev_track_count = timeline.GetTrackCount("video")
    try:
        timeline.AddTrack("video")
    except Exception as e:
        _log(f"AddTrack failed: {e}")

    track_index = timeline.GetTrackCount("video")
    if track_index <= prev_track_count:
        _log(f"ERROR: Failed to create caption track (count stayed at {track_index})")
        return False
    _log(f"Caption track index: {track_index} (was {prev_track_count})")

    tl_start = timeline.GetStartFrame() or 0

    # Build clipInfo list for batch insertion
    clip_infos = []
    for cap in captions:
        start_frame = round(cap["start"] * fps) + tl_start
        duration_frames = max(1, round((cap["end"] - cap["start"]) * fps))

        clip_infos.append({
            "mediaPoolItem": template,
            "startFrame": 0,
            "endFrame": duration_frames,
            "trackIndex": track_index,
            "recordFrame": start_frame,
            "mediaType": 1,
        })

    # Batch insert
    items = None
    try:
        items = media_pool.AppendToTimeline(clip_infos)
        _log(f"Batch AppendToTimeline: {len(items) if items else 0} items")
    except Exception as e:
        _log(f"Batch AppendToTimeline failed: {e}")

    # Fallback: one by one
    if not items:
        _log("Falling back to one-by-one insertion...")
        items = []
        for ci in clip_infos:
            try:
                result = media_pool.AppendToTimeline([ci])
                if result:
                    items.extend(result)
            except Exception as e:
                _log(f"Single AppendToTimeline failed: {e}")

    if not items:
        _log("No items were inserted")
        return False

    _log(f"Inserted {len(items)} clips, adding Text+ via Fusion...")

    # Let Resolve finish processing the batch insert
    time.sleep(0.5)

    # Get fresh item references from the timeline track
    # (AppendToTimeline references can be stale for Fusion operations)
    track_items = timeline.GetItemListInTrack("video", track_index)
    if track_items and len(track_items) > 0:
        _log(f"Refreshed: {len(track_items)} items on track {track_index}")
        items = track_items

    # Add TextPlus node to each clip's Fusion comp and set caption text
    success_count = 0
    total = len(items)
    for i, item in enumerate(items):
        if i >= len(captions):
            break
        cap_text = captions[i]["text"]

        if on_progress:
            try:
                on_progress(int((i + 1) / total * 100))
            except Exception:
                pass

        try:
            comp = item.GetFusionCompByIndex(1)

            # PNG clips don't have Fusion comps by default -- create one
            if not comp:
                try:
                    comp = item.AddFusionComp()
                    if comp:
                        _log(f"  Caption {i+1}: created Fusion comp")
                    else:
                        _log(f"  Caption {i+1}: AddFusionComp returned None")
                except Exception as e:
                    _log(f"  Caption {i+1}: AddFusionComp error: {e}")

            if not comp:
                _log(f"  Caption {i+1}: no Fusion comp")
                # Log diagnostic on first failure to help debug
                if success_count == 0 and i < 3:
                    _log("  (If all clips fail here, AddFusionComp may not be "
                         "supported in this Resolve version)")
                continue

            # Diagnostic logging for first clip
            if i == 0:
                try:
                    tools = comp.GetToolList()
                    tool_names = [str(t.Name) for t in tools.values()] if tools else []
                    _log(f"  First clip comp tools: {tool_names}")
                except Exception:
                    pass

            # Check if TextPlus already exists (user-provided template)
            existing = comp.GetToolList(False, "TextPlus")
            if existing:
                for tool in existing.values():
                    tool.SetInput("StyledText", cap_text)
                    success_count += 1
                    break
                continue

            # Add TextPlus node and wire it to MediaOut
            comp.Lock()
            try:
                tp = comp.AddTool("TextPlus", -32768, -32768)
                if tp:
                    tp.SetInput("StyledText", cap_text)
                    tp.SetInput("Font", "Open Sans")
                    tp.SetInput("Style", "Bold")
                    tp.SetInput("Size", 0.042)
                    tp.SetInput("Center", {1: 0.5, 2: 0.08})
                    tp.SetInput("Red1", 1.0)
                    tp.SetInput("Green1", 1.0)
                    tp.SetInput("Blue1", 1.0)
                    tp.SetInput("Enabled2", 1)
                    tp.SetInput("Red2", 0.0)
                    tp.SetInput("Green2", 0.0)
                    tp.SetInput("Blue2", 0.0)
                    tp.SetInput("Thickness2", 0.15)

                    # Connect TextPlus output to MediaOut
                    # (this disconnects MediaIn -- TextPlus is now the sole output)
                    media_out = comp.FindTool("MediaOut1")
                    if media_out:
                        media_out.FindMainInput(1).ConnectTo(tp.FindMainOutput(1))
                        success_count += 1
                    else:
                        _log(f"  Caption {i+1}: no MediaOut1 in comp (TextPlus won't render)")
                else:
                    _log(f"  Caption {i+1}: AddTool returned None")
            finally:
                comp.Unlock()

        except Exception as e:
            _log(f"  Caption {i+1} failed: {e}")

    _log(f"Text+ set on {success_count}/{len(items)} clips")

    # Rename the track
    try:
        timeline.SetTrackName("video", track_index, "Captions")
    except Exception:
        pass

    return success_count > 0


def _run_caption_subprocess(cmd, update_bar, progress_root):
    """Run caption.py subprocess. Returns (returncode, stderr_text)."""
    _log(f"Running: {' '.join(cmd)}")

    popen_kwargs = dict(
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        cwd=APP_DIR,
    )
    if IS_WIN:
        popen_kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW

    proc = subprocess.Popen(cmd, **popen_kwargs)
    try:
        # Drain stderr in a background thread to prevent pipe deadlock
        stderr_lines = []
        def _drain_stderr():
            try:
                for line in proc.stderr:
                    stderr_lines.append(line)
            except Exception:
                pass
        stderr_thread = threading.Thread(target=_drain_stderr, daemon=True)
        stderr_thread.start()

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

        stderr_thread.join(timeout=5.0)

        try:
            returncode = proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            _log("Caption subprocess timed out, killing it")
            proc.kill()
            proc.wait(timeout=5)
            returncode = -1

        stderr_text = "".join(stderr_lines)
        return returncode, stderr_text

    except Exception:
        # Clean up process on any unexpected error
        try:
            proc.kill()
            proc.wait(timeout=5)
        except Exception:
            pass
        raise


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
            _modules_paths = []
            if IS_WIN:
                _modules_paths.append(
                    os.path.join(os.environ.get("PROGRAMDATA", "C:\\ProgramData"),
                                 "Blackmagic Design", "DaVinci Resolve",
                                 "Support", "Developer", "Scripting", "Modules")
                )
            elif IS_MAC:
                _modules_paths.extend([
                    "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules",
                    os.path.expanduser("~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules"),
                ])
            else:
                # Linux
                _modules_paths.extend([
                    "/opt/resolve/Developer/Scripting/Modules",
                    "/opt/resolve/libs/Fusion/Modules",
                ])
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
    if not pm:
        show_error(_TITLE, "Could not access Resolve's project manager.")
        return
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

    try:
        start_frame = timeline.GetStartFrame() or 0
        end_frame = timeline.GetEndFrame() or 0
        duration_s = (end_frame - start_frame) / fps if fps > 0 else 0
        _log(f"Timeline: frame {start_frame}-{end_frame}, {duration_s:.1f}s, FPS: {fps}")
    except Exception as e:
        _log(f"Could not read timeline frame range: {e}")
        duration_s = 0

    # Show settings dialog
    settings = show_settings(timeline_name, duration_s, fps)
    if settings is None:
        _log("User cancelled settings dialog")
        return

    max_words = settings["max_words"]
    max_chars = settings["max_chars"]
    max_lines = settings["max_lines"]
    language = settings.get("language")
    strip_punctuation = settings.get("strip_punctuation", False)
    output_mode = settings.get("output_mode", "srt")
    _log(f"Settings: max_words={max_words}, max_chars={max_chars}, max_lines={max_lines}, language={language}, strip_punct={strip_punctuation}, output={output_mode}")

    # Initialize cleanup vars before try block so they're always bound
    tmp_dir = None
    job_id = None
    close_progress = lambda: None
    update_bar = lambda p: None
    update_phase = lambda s, t, n: None
    progress_root = None

    # Show progress
    close_progress, _update_text, update_bar, update_phase, progress_root = show_progress(
        _TITLE,
        f"Rendering audio from '{timeline_name}'"
    )
    update_phase(1, 3, "Rendering audio")

    try:
        # Render audio for current in/out range
        # We do NOT set MarkIn/MarkOut - Resolve automatically uses
        # the timeline's in/out points (set with I and O keys).
        # This matches how the built-in transcription tool works.
        tmp_dir = tempfile.mkdtemp(prefix="resolve_whisper_")
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
        wav_name = f"{safe_name}_audio"

        _log(f"Rendering audio to {tmp_dir}")

        # Set render format: try MP4/H.264 first (reliable), Wave as fallback.
        # Wave codec list is empty on many Resolve versions, so MP4 audio-only
        # is the safest option. faster-whisper handles MP4 audio fine.
        _format_set = False
        _render_combos = [
            ("MP4", "H264_NVIDIA"),   # NVIDIA GPU encode (fastest)
            ("MP4", "H264"),          # CPU H.264
            ("MP4", "H265_NVIDIA"),
            ("MP4", "H265"),
            ("Wave", "LinearPCM"),
        ]
        for fmt, codec in _render_combos:
            try:
                if project.SetCurrentRenderFormatAndCodec(fmt, codec):
                    _log(f"Render format set: {fmt}/{codec}")
                    _format_set = True
                    break
            except Exception:
                pass

        if not _format_set:
            _log("WARNING: Could not set any render format")
            try:
                fmts = project.GetRenderFormats()
                _log(f"Available formats: {fmts}")
            except Exception:
                pass

        project.SetRenderSettings({
            "ExportAudio": True,
            "ExportVideo": True,
            "TargetDir": tmp_dir,
            "CustomName": wav_name,
        })
        _log("Render settings applied (using timeline in/out points)")

        job_id = project.AddRenderJob()
        _log(f"Render job ID: {job_id}")

        if not job_id:
            close_progress()
            show_error(_TITLE, "Failed to create render job.\n\nCheck your render settings.")
            return

        project.StartRendering([job_id])
        _log("Rendering started")

        # Give Resolve a moment to transition from Ready to Rendering
        time.sleep(2)

        # Wait for render with timeout - check actual job status, not just
        # IsRenderingInProgress() which can return False if render fails instantly
        _render_timeout = 600  # 10 minutes max
        _render_start = time.time()
        _ready_count = 0  # how many consecutive polls show "Ready"
        status = {}
        while True:
            try:
                status = project.GetRenderJobStatus(job_id)
                job_status = status.get("JobStatus", "")
                pct = status.get("CompletionPercentage", 0)
                if pct and progress_root:
                    update_bar(int(pct))
            except Exception:
                job_status = ""

            # Done states
            if job_status.lower() in ("complete", "completed"):
                break
            if job_status.lower() in ("failed", "cancelled", "canceled"):
                _log(f"Render failed with status: {job_status}")
                close_progress()
                show_error(
                    _TITLE,
                    f"Audio render failed ({job_status}).\n\nMake sure your timeline has audio."
                )
                return

            # If status stays "Ready" for 10+ consecutive checks (5+ seconds),
            # the render truly never started
            if not project.IsRenderingInProgress() and job_status.lower() in ("ready", ""):
                _ready_count += 1
                if _ready_count >= 10:
                    _log(f"Render never started (status: {job_status})")
                    close_progress()
                    show_error(
                        _TITLE,
                        "Render never started.\n\nTry a manual render on the Deliver page first."
                    )
                    return
            else:
                _ready_count = 0

            # Timeout guard
            if time.time() - _render_start > _render_timeout:
                _log("Render timed out")
                close_progress()
                show_error(_TITLE, "Audio render timed out after 10 minutes.")
                return

            if progress_root:
                try:
                    progress_root.update()
                except Exception:
                    pass
            time.sleep(0.5)

        _log(f"Render complete: {status}")

        _log("Rendering finished")

        # Remove our job from the render queue so we don't clutter it
        try:
            project.DeleteRenderJob(job_id)
            job_id = None
            _log("Render job cleaned up from queue")
        except Exception:
            pass

        # Switch back to Edit page (rendering puts us on Deliver)
        try:
            resolve.OpenPage("edit")
            _log("Switched back to Edit page")
        except Exception:
            pass

        # Find rendered audio file (accept any format - faster-whisper handles all)
        audio_path = None
        try:
            files = os.listdir(tmp_dir)
            _log(f"Files in tmp_dir: {files}")
            for ext in [".wav", ".mov", ".mp4", ".flac", ".mp3", ".aac", ".m4a", ".mxf"]:
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
            show_error(
                _TITLE,
                "Audio render finished but no file was created."
            )
            return

        try:
            _log(f"Audio file: {audio_path} ({os.path.getsize(audio_path)} bytes)")
        except Exception:
            _log(f"Audio file: {audio_path}")

        # Transcribe
        update_phase(2, 3, "Transcribing audio")

        output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Captions")
        os.makedirs(output_dir, exist_ok=True)
        # Per-run timestamp so Resolve treats each import as a fresh media
        # item (same convention as the Mac path in caption.py).
        run_stamp = time.strftime("%Y%m%d-%H%M%S")
        srt_path = os.path.join(output_dir, f"{safe_name} {run_stamp}.srt")

        cmd = [
            VENV_PYTHON,
            CAPTION_SCRIPT,
            "--file", audio_path,
            "--output", srt_path,
            "--fps", str(fps),
            "--max-words", str(max_words),
            "--max-chars", str(max_chars),
            "--max-lines", str(max_lines),
            # We already drive our own progress window; without this the
            # caption.py subprocess spawns a second one.
            "--no-ui",
        ]
        if language:
            cmd.extend(["--language", language])
        if strip_punctuation:
            cmd.append("--strip-punctuation")

        returncode, stderr_text = _run_caption_subprocess(
            cmd, update_bar, progress_root
        )

        _log(f"Return code: {returncode}")
        if stderr_text:
            _log(f"stderr: ...\n{stderr_text[-2000:]}")

        # Auto-retry on CUDA crash: if the process died with no Python-level
        # error output (only [INFO] lines), it's likely a GPU crash.
        # Retry once on CPU which is slower but reliable.
        if returncode != 0:
            _has_python_error = any(
                not ("[INFO]" in l or l.strip().startswith("PROGRESS:") or not l.strip())
                for l in stderr_text.split("\n")
            )
            if not _has_python_error:
                _log("No error trace in stderr -- possible GPU crash. Retrying on CPU...")
                update_phase(2, 3, "Retrying transcription (CPU)")
                update_bar(0)
                retry_cmd = cmd + ["--device", "cpu"]
                returncode, stderr_text = _run_caption_subprocess(
                    retry_cmd, update_bar, progress_root
                )
                _log(f"Retry return code: {returncode}")
                if stderr_text:
                    _log(f"Retry stderr: ...\n{stderr_text[-2000:]}")

        if returncode != 0:
            close_progress()
            # Find the actual error line (skip INFO/PROGRESS lines)
            error_msg = None
            if stderr_text:
                for line in reversed(stderr_text.strip().split("\n")):
                    line = line.strip()
                    if not line:
                        continue
                    if "[INFO]" in line or line.startswith("PROGRESS:"):
                        continue
                    error_msg = line
                    # Strip timestamp prefix like "11:44:07 [ERROR] ..."
                    if "] " in error_msg:
                        error_msg = error_msg.split("] ", 1)[-1]
                    break
            if not error_msg:
                error_msg = "Transcription failed.\nCheck the debug log on your Desktop."
            show_error(_TITLE, error_msg)
            return

        if not os.path.exists(srt_path):
            close_progress()
            show_error(_TITLE, "Transcription finished but SRT file was not created.")
            return

        # Count captions
        with open(srt_path, "r", encoding="utf-8-sig") as f:
            content = f.read()
        caption_count = sum(1 for line in content.split("\n") if line.strip().isdigit())

        _log(f"Done! {caption_count} captions written to {srt_path}")

        # Insert captions based on output mode
        if output_mode == "textplus":
            update_phase(3, 3, "Inserting captions into timeline")

            # Read JSON sidecar with structured caption data
            import json as _json
            json_path = os.path.splitext(srt_path)[0] + ".json"
            captions_data = None

            if os.path.exists(json_path):
                try:
                    with open(json_path, "r", encoding="utf-8") as jf:
                        captions_data = _json.load(jf)
                    _log(f"Loaded {len(captions_data)} captions from JSON")
                except Exception as e:
                    _log(f"JSON load failed: {e}")

            textplus_ok = False
            if captions_data:
                media_pool = project.GetMediaPool()
                if media_pool:
                    tl_width = int(timeline.GetSetting("timelineResolutionWidth") or 1920)
                    tl_height = int(timeline.GetSetting("timelineResolutionHeight") or 1080)
                    template = _get_or_create_template(media_pool, tl_width, tl_height)
                    if template:
                        textplus_ok = _insert_textplus_captions(
                            media_pool, timeline, captions_data, fps, template,
                            on_progress=update_bar,
                        )
                    else:
                        _log("Could not get/create template")
                else:
                    _log("Could not access media pool")

            close_progress()

            if textplus_ok:
                show_message(
                    _TITLE,
                    f"{caption_count} captions added to '{timeline_name}'\n\n"
                    f"SRT backup: {srt_path}"
                )
            else:
                _log("Text+ insertion failed, falling back to SRT mode")
                open_file_selected(srt_path)
                show_message(
                    _TITLE,
                    f"{caption_count} captions exported\n\n"
                    f"Text+ insertion failed. Import manually:\n"
                    f"File > Import > Subtitle\n\n"
                    f"{srt_path}"
                )
        else:
            # SRT file mode
            update_phase(3, 3, "Importing captions")
            imported = _import_srt(project, timeline, srt_path)

            close_progress()

            if imported:
                show_message(
                    _TITLE,
                    f"{caption_count} captions added to '{timeline_name}'\n\n"
                    f"SRT: {srt_path}"
                )
            else:
                open_file_selected(srt_path)
                show_message(
                    _TITLE,
                    f"{caption_count} captions exported\n\n"
                    f"Import: File > Import > Subtitle\n\n"
                    f"{srt_path}"
                )

    except Exception as e:
        _log(f"Exception in main: {e}\n{traceback.format_exc()}")
        close_progress()
        show_error(_TITLE, f"Unexpected error:\n{e}")

    finally:
        # Subprocess cleanup is handled inside _run_caption_subprocess()

        # Clean up orphaned render job from Resolve's queue
        if job_id is not None:
            try:
                project.DeleteRenderJob(job_id)
                _log("Cleaned up orphaned render job")
            except Exception:
                pass

        # Clean up temp directory (shutil handles non-empty dirs)
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
                _log(f"Cleaned up temp dir: {tmp_dir}")
            except Exception:
                pass

        # Purge Tcl/Tk state from Resolve's Python interpreter.
        # Tkinter's C extension keeps internal state that can destabilize
        # Resolve's process if left behind after dialogs close.
        try:
            import tkinter as _tk_cleanup
            if _tk_cleanup._default_root:
                try:
                    _tk_cleanup._default_root.quit()
                    _tk_cleanup._default_root.destroy()
                except Exception:
                    pass
                _tk_cleanup._default_root = None
        except Exception:
            pass

        import gc
        gc.collect()
        _log("Cleanup complete")


try:
    main()
except Exception as e:
    _log(f"FATAL: {e}\n{traceback.format_exc()}")
    try:
        show_error(_TITLE, f"Fatal error:\n{e}")
    except Exception:
        pass
finally:
    # Final Tcl/Tk purge (covers both normal exit and fatal error dialogs)
    try:
        import tkinter as _tk_final
        if _tk_final._default_root:
            try:
                _tk_final._default_root.quit()
                _tk_final._default_root.destroy()
            except Exception:
                pass
            _tk_final._default_root = None
    except Exception:
        pass
