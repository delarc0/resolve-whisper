"""Settings dialog for the LAB37 Custom preset.

Standalone Tk window (spawned as a subprocess by caption.py --dialog, so it
never runs inside Resolve's process). The user picks language and caption
shape; choices persist to custom_settings.json so the dialog opens prefilled
next time - after one setup it's Enter-and-go.

Protocol: prints the chosen settings as one JSON line on stdout and exits 0.
Exit 1 = cancelled, exit 2 = tkinter unavailable.
"""
import json
import os
import sys

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH = os.path.join(_APP_DIR, "custom_settings.json")

DEFAULTS = {
    "language": "sv",
    "max_words": 0,
    "max_chars": 42,
    "uppercase": True,
    "strip_punctuation": False,
}

LANGUAGES = [
    ("Auto-detect", "auto"),
    ("Svenska", "sv"),
    ("English", "en"),
    ("Suomi", "fi"),
    ("Norsk", "no"),
    ("Dansk", "da"),
    ("Deutsch", "de"),
    ("Español", "es"),
    ("Français", "fr"),
]

STYLES = [
    ("Full sentences", 0),
    ("1 word at a time", 1),
    ("2 words at a time", 2),
    ("3 words (reels)", 3),
    ("4 words at a time", 4),
    ("5 words at a time", 5),
]

LINE_LENGTHS = [
    ("22 characters (9:16 reels)", 22),
    ("32 characters (narrow)", 32),
    ("37 characters (medium)", 37),
    ("42 characters (16:9 standard)", 42),
    ("50 characters (wide)", 50),
]


def load_settings() -> dict:
    """Last-used settings, falling back to defaults per key."""
    merged = dict(DEFAULTS)
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            for key in DEFAULTS:
                if key in saved and type(saved[key]) is type(DEFAULTS[key]):
                    merged[key] = saved[key]
    except (OSError, json.JSONDecodeError, ValueError):
        pass
    return merged


def save_settings(settings: dict):
    """Persist for next run. Non-fatal on failure."""
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
            f.write("\n")
    except OSError:
        pass


def _with_stored(options, value, label_fmt):
    """Ensure a stored off-list value still appears as a choice."""
    if any(v == value for _, v in options):
        return options
    return options + [(label_fmt.format(value), value)]


def run_dialog() -> "dict | None":
    import tkinter as tk
    import ui_theme as th

    initial = load_settings()
    result = [None]

    root = tk.Tk()
    root.title("LAB37 TOOLS: Whisper")
    root.configure(bg=th.BG)
    root.resizable(False, False)
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    th.edge_light(root)

    frame = tk.Frame(root, bg=th.BG, padx=28, pady=20)
    frame.pack(fill="both", expand=True)

    th.brand_header(frame, "Custom Captions",
                    "Match the captions to your content. "
                    "Choices are remembered.")

    th.section_label(frame, "Captions", pady=(16, 6))
    get_style = th.make_dropdown(
        frame, "Caption style",
        _with_stored(STYLES, initial["max_words"], "{} words at a time"),
        initial["max_words"], width=24)
    get_chars = th.make_dropdown(
        frame, "Line length",
        _with_stored(LINE_LENGTHS, initial["max_chars"], "{} characters"),
        initial["max_chars"], width=24)

    th.section_label(frame, "Language")
    get_lang = th.make_dropdown(frame, "Spoken language", LANGUAGES,
                                initial["language"], width=24)

    th.section_label(frame, "Text")
    get_caps = th.make_toggle(frame, "ALL CAPS", initial["uppercase"])
    get_strip = th.make_toggle(frame, "Remove punctuation",
                               initial["strip_punctuation"])

    tk.Frame(frame, bg=th.BORDER, height=1).pack(fill="x", pady=(14, 12))

    btn_row = tk.Frame(frame, bg=th.BG)
    btn_row.pack(fill="x")

    def _cancel(_e=None):
        result[0] = None
        root.destroy()

    def _start(_e=None):
        result[0] = {
            "language": get_lang(),
            "max_words": get_style(),
            "max_chars": get_chars(),
            "uppercase": get_caps(),
            "strip_punctuation": get_strip(),
        }
        root.destroy()

    th.make_button(btn_row, "Cancel", "secondary", _cancel).pack(side="left")
    start_btn = th.make_button(btn_row, "Generate Captions", "primary", _start)
    start_btn.pack(side="right")

    root.bind("<Return>", _start)
    root.bind("<Escape>", _cancel)
    root.protocol("WM_DELETE_WINDOW", _cancel)
    th.place_window(root)
    start_btn.focus_set()
    root.mainloop()
    return result[0]


def main() -> int:
    try:
        import tkinter  # noqa: F401
    except ImportError:
        sys.stderr.write("tkinter not available; settings dialog disabled.\n")
        return 2

    settings = run_dialog()
    if settings is None:
        return 1
    save_settings(settings)
    print("SETTINGS:" + json.dumps(settings))
    return 0


if __name__ == "__main__":
    sys.exit(main())
