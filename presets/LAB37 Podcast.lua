-- LAB37 Podcast -- standard SRT subtitles, full sentences, Swedish

local PRESET_ARGS = "--language sv --max-words 0 --max-chars 42 --max-lines 1"

print("[LAB37 Podcast] Starting...")

local RESOLVE_API = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting"
local RESOLVE_LIB = "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so"

local lf = io.open(RESOLVE_LIB, "r")
if lf then lf:close() else
    print("[LAB37 Podcast] ERROR: fusionscript.so not found at: " .. RESOLVE_LIB)
    print("[LAB37 Podcast] Is Resolve installed in a nonstandard location?")
    return
end

local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*/)") or "./"
local f = io.open(script_dir .. "resolve_whisper_path.txt", "r")
if not f then
    print("[LAB37 Podcast] ERROR: pointer file not found")
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

print("[LAB37 Podcast] Launching subprocess.")
os.execute(cmd)
