"""
LAB37 TOOLS: Resolve Whisper (launcher)
Loads resolve_script.py from the project directory so code changes
are always live without re-running setup.
"""
import os
import sys

_STUB_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in dir() else os.path.dirname(os.path.abspath(sys.argv[0]))
_POINTER = os.path.join(_STUB_DIR, "resolve_whisper_path.txt")

if not os.path.exists(_POINTER):
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        messagebox.showerror(
            "LAB37 TOOLS: Resolve Whisper",
            f"Pointer file not found:\n{_POINTER}\n\nRe-run setup.bat (Windows) or setup.sh (Mac).",
        )
        root.destroy()
    except Exception:
        print(f"ERROR: Pointer file not found: {_POINTER}")
    raise SystemExit(1)

with open(_POINTER, "r", encoding="utf-8") as _f:
    _PROJECT_DIR = _f.read().strip()

_SCRIPT_PATH = os.path.join(_PROJECT_DIR, "resolve_script.py")

if not os.path.exists(_SCRIPT_PATH):
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        messagebox.showerror(
            "LAB37 TOOLS: Resolve Whisper",
            f"Script not found:\n{_SCRIPT_PATH}\n\nRe-run setup.bat (Windows) or setup.sh (Mac).",
        )
        root.destroy()
    except Exception:
        print(f"ERROR: Script not found: {_SCRIPT_PATH}")
    raise SystemExit(1)

# Load and exec the real script.
# - dict(globals()) preserves Resolve's injected 'bmd' and 'fusion' globals
# - __file__ override so resolve_script.py resolves paths to the project dir
# - compile() with the real path gives correct tracebacks
with open(_SCRIPT_PATH, "r", encoding="utf-8") as _f:
    _code = _f.read()

_g = dict(globals())
_g["__file__"] = _SCRIPT_PATH
exec(compile(_code, _SCRIPT_PATH, "exec"), _g)
