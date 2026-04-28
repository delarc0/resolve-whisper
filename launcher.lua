-- LAB37 Resolve Whisper -- shared launcher
-- Each preset .lua file sets PRESET_ARGS then dofile()s this script.
--
-- Mac Resolve 20 only lists .lua files in the Scripts menu (not .py), so
-- this Lua wrapper is the menu entry. It shells out to caption.py via the
-- project's venv. PATH is augmented to include /opt/homebrew/bin since
-- Resolve's GUI subprocess inherits a minimal env without it.

local args = PRESET_ARGS
if not args or args == "" then
    print("[LAB37] PRESET_ARGS not set -- this script is loaded by preset wrappers")
    return
end
PRESET_ARGS = nil  -- reset for the next invocation

local script_dir = debug.getinfo(2, "S").source:sub(2):match("(.*/)") or "./"
local pointer = script_dir .. "resolve_whisper_path.txt"
local f = io.open(pointer, "r")
if not f then
    print(string.format("[LAB37] Pointer file not found: %s", pointer))
    return
end
local app_dir = f:read("*all"):gsub("%s+$", "")
f:close()

local python = app_dir .. "/.venv/bin/python3"
local caption = app_dir .. "/caption.py"
local log = "/tmp/resolve_whisper.log"

local cmd = string.format(
    [[(export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"; ]] ..
    [[export RESOLVE_SCRIPT_API=%q; ]] ..
    [[export RESOLVE_SCRIPT_LIB=%q; ]] ..
    [[export PYTHONPATH="$RESOLVE_SCRIPT_API/Modules"; ]] ..
    [[cd %q && %q %q %s) > %q 2>&1 &]],
    "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting",
    "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so",
    app_dir, python, caption, args, log
)

print(string.format("[LAB37] Launched: caption.py %s", args))
print(string.format("[LAB37] Output: %s", log))
os.execute(cmd)
