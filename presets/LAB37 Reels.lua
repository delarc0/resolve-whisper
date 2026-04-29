-- LAB37 Reels -- single-line SRT, 1-3 words, no punctuation
-- Self-contained: no dofile/require -- Resolve's Lua sandbox doesn't always
-- handle them the way standard Lua does.
--
-- max-words=3 (not 4): 4 reads as dense on a 9x16 reel even when it fits
-- max-chars. Combined with the connector-suppression in srt.py, output
-- variation lands naturally at 1-3 words depending on speech rhythm.
local PRESET_ARGS = "--language sv --max-words 3 --max-chars 22 --max-lines 1 --strip-punctuation"

print("[LAB37 Reels] Starting...")

local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*/)") or "./"
local pointer = script_dir .. "resolve_whisper_path.txt"
print("[LAB37 Reels] Looking for pointer at: " .. pointer)

local f = io.open(pointer, "r")
if not f then
    print("[LAB37 Reels] ERROR: pointer file not found")
    return
end
local app_dir = f:read("*all"):gsub("%s+$", "")
f:close()
print("[LAB37 Reels] App dir: " .. app_dir)

local python = app_dir .. "/.venv/bin/python3"
local caption = app_dir .. "/caption.py"
local log_path = "/tmp/resolve_whisper.log"

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

print("[LAB37 Reels] Launching subprocess (output: " .. log_path .. ")")
os.execute(cmd)
print("[LAB37 Reels] Subprocess launched in background.")
