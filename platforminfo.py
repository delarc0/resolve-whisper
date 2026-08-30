"""Single source of truth for platform differences.

Every OS branch in the pipeline routes through here so the platform logic is
auditable in one place instead of scattered `if sys.platform` checks. The
Mac and Windows paths never share fragile code: this module only resolves
paths and tools, it does not change pipeline behavior.

Named platforminfo (not `platform`) to avoid shadowing the stdlib module,
which the app dir being on sys.path would otherwise do.
"""
import os
import shutil
import sys

IS_MAC = sys.platform == "darwin"
IS_WIN = sys.platform == "win32"
IS_LINUX = sys.platform.startswith("linux")


def resolve_scripts_dir() -> str:
    """The DaVinci Resolve 'Scripts/Edit' folder where preset .lua files live.

    These are the documented per-user locations; Resolve reads .lua menu
    entries from here on each platform.
    """
    if IS_MAC:
        return os.path.expanduser(
            "~/Library/Application Support/Blackmagic Design/"
            "DaVinci Resolve/Fusion/Scripts/Edit")
    if IS_WIN:
        appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
        return os.path.join(
            appdata, "Blackmagic Design", "DaVinci Resolve", "Support",
            "Fusion", "Scripts", "Edit")
    # Linux
    return os.path.expanduser(
        "~/.local/share/DaVinciResolve/Fusion/Scripts/Edit")


def resolve_script_api() -> str:
    """RESOLVE_SCRIPT_API dir (holds Modules/DaVinciResolveScript)."""
    if IS_MAC:
        return ("/Library/Application Support/Blackmagic Design/"
                "DaVinci Resolve/Developer/Scripting")
    if IS_WIN:
        programdata = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
        return os.path.join(
            programdata, "Blackmagic Design", "DaVinci Resolve", "Support",
            "Developer", "Scripting")
    return "/opt/resolve/Developer/Scripting"


def resolve_script_lib() -> str:
    """Absolute path to the Resolve scripting library (fusionscript)."""
    if IS_MAC:
        return ("/Applications/DaVinci Resolve/DaVinci Resolve.app/"
                "Contents/Libraries/Fusion/fusionscript.so")
    if IS_WIN:
        programfiles = os.environ.get("ProgramFiles", r"C:\Program Files")
        return os.path.join(
            programfiles, "Blackmagic Design", "DaVinci Resolve",
            "fusionscript.dll")
    return "/opt/resolve/libs/Fusion/fusionscript.so"


def log_dir() -> str:
    """Directory for the tool's log files.

    Deliberately NOT tempfile.gettempdir() on Mac: a GUI-launched Resolve
    hands its children a private TMPDIR (/var/folders/.../T), so a log
    written there is invisible to anyone following the documented
    `cat /tmp/resolve_whisper*.log`. The Lua launcher already redirects
    stdout to /tmp, so both halves agree here.
    """
    if IS_WIN:
        import tempfile
        return os.environ.get("TEMP") or tempfile.gettempdir()
    return "/tmp"


def cache_dir() -> str:
    """Writable scratch that survives reboots, for things like the update check.

    Deliberately NOT inside the app dir: that is a git checkout, and a state
    file written there would show up in `git status --porcelain`, which is
    how version.py decides whether the build is modified. The tool would
    report every install as carrying local changes.
    """
    if IS_WIN:
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        path = os.path.join(base, "lab37-whisper")
    else:
        path = os.path.expanduser("~/.cache/lab37-whisper")
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        return ""
    return path


def add_cuda_dll_dir():
    """Make CTranslate2 (the faster-whisper GPU backend) find cuBLAS/cuDNN.

    CTranslate2 needs those CUDA DLLs but does not look inside torch's bundle
    on its own. The CUDA torch wheel ships them in `torch/lib`, so on Windows
    we add that dir to the DLL search path before the model loads. No-op off
    Windows or when torch isn't installed. Safe to call more than once.
    """
    if not IS_WIN:
        return
    try:
        import torch
        lib = os.path.join(os.path.dirname(torch.__file__), "lib")
        if os.path.isdir(lib) and hasattr(os, "add_dll_directory"):
            os.add_dll_directory(lib)
    except Exception:
        pass  # falls back to CPU at model load if CUDA libs can't be found


def bootstrap_resolve_env():
    """Make `import DaVinciResolveScript` work without the launcher exporting
    env vars.

    The Lua presets used to export RESOLVE_SCRIPT_API/LIB + PYTHONPATH before
    running us, which meant fragile per-OS shell quoting. Instead we set them
    here from the known install locations if they aren't already set, and add
    the Modules dir to sys.path directly. Honors pre-set env vars so a
    nonstandard install can still override via the environment.
    """
    api = os.environ.get("RESOLVE_SCRIPT_API") or resolve_script_api()
    lib = os.environ.get("RESOLVE_SCRIPT_LIB") or resolve_script_lib()
    os.environ.setdefault("RESOLVE_SCRIPT_API", api)
    os.environ.setdefault("RESOLVE_SCRIPT_LIB", lib)
    modules = os.path.join(api, "Modules")
    if os.path.isdir(modules) and modules not in sys.path:
        sys.path.insert(0, modules)


def _tool_candidates(name: str):
    """Well-known install dirs to check when a tool isn't on PATH."""
    if IS_WIN:
        exe = name if name.lower().endswith(".exe") else name + ".exe"
        programfiles = os.environ.get("ProgramFiles", r"C:\Program Files")
        return [
            os.path.join(programfiles, "ffmpeg", "bin", exe),
            os.path.join(r"C:\ffmpeg\bin", exe),
        ]
    return [f"/opt/homebrew/bin/{name}", f"/usr/local/bin/{name}"]


def find_tool(name: str):
    """Locate ffmpeg/ffprobe even when the host launches us with a bare PATH.

    Resolve spawns our subprocess without the user's shell PATH, so a plain
    `which` misses Homebrew (Mac) or a non-PATH ffmpeg (Windows). Check PATH
    first, then well-known locations.
    """
    found = shutil.which(name)
    if found:
        return found
    for candidate in _tool_candidates(name):
        if os.path.exists(candidate):
            return candidate
    return None
