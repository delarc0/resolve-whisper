-- LAB37 Check -- runs the pre-flight check (env + Resolve + Audio Only preset)
-- Run this after install or after a Resolve update to confirm the tool is healthy.
local PRESET_ARGS = "--check"

print("[LAB37 Check] Starting...")

local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*/)") or "./"
local pointer = script_dir .. "resolve_whisper_path.txt"
print("[LAB37 Check] Looking for pointer at: " .. pointer)

local f = io.open(pointer, "r")
if not f then
    print("[LAB37 Check] ERROR: pointer file not found")
    return
end
local app_dir = f:read("*all"):gsub("%s+$", "")
f:close()
print("[LAB37 Check] App dir: " .. app_dir)

local python = app_dir .. "/.venv/bin/python3"
local caption = app_dir .. "/caption.py"
local log_path = "/tmp/resolve_whisper_check.log"

local cmd = string.format(
    [[(export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"; ]] ..
    [[export RESOLVE_SCRIPT_API=%q; ]] ..
    [[export RESOLVE_SCRIPT_LIB=%q; ]] ..
    [[export PYTHONPATH="$RESOLVE_SCRIPT_API/Modules"; ]] ..
    [[cd %q && %q %q %s) > %q 2>&1 &]],
    "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting",
    "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so",
    app_dir, python, caption, PRESET_ARGS, log_path
)

print("[LAB37 Check] Launching subprocess (output: " .. log_path .. ")")
os.execute(cmd)
print("[LAB37 Check] Subprocess launched in background. Check the log file when done.")
