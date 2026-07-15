-- LAB37 Custom -- opens a settings dialog before running
--
-- The user picks language, caption style, and line length; choices are
-- remembered in custom_settings.json for next time.
local PRESET_ARGS = "--dialog"

print("[LAB37 Custom] Starting...")

local RESOLVE_API = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting"
local RESOLVE_LIB = "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so"

-- Fail loudly if Resolve moved (nonstandard install location); the
-- subprocess would otherwise die silently into the log.
local lf = io.open(RESOLVE_LIB, "r")
if lf then lf:close() else
    print("[LAB37 Custom] ERROR: fusionscript.so not found at: " .. RESOLVE_LIB)
    print("[LAB37 Custom] Is Resolve installed in a nonstandard location?")
    return
end

local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*/)") or "./"
local f = io.open(script_dir .. "resolve_whisper_path.txt", "r")
if not f then
    print("[LAB37 Custom] ERROR: pointer file not found")
    return
end
local app_dir = f:read("*all"):gsub("%s+$", "")
f:close()

local cmd = string.format(
    [[(export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"; ]] ..
    [[export RESOLVE_SCRIPT_API=%q; ]] ..
    [[export RESOLVE_SCRIPT_LIB=%q; ]] ..
    [[export PYTHONPATH="$RESOLVE_SCRIPT_API/Modules"; ]] ..
    [[cd %q && %q %q %s) > %q 2>&1 &]],
    RESOLVE_API, RESOLVE_LIB,
    app_dir, app_dir .. "/.venv/bin/python3", app_dir .. "/caption.py",
    PRESET_ARGS, "/tmp/resolve_whisper.log"
)

print("[LAB37 Custom] Launching subprocess.")
os.execute(cmd)
