-- LAB37 Reels -- single-line SRT, 1-3 words, no punctuation, Swedish
LAB37_TOOL = "Reels"
LAB37_ARGS = "--language sv --max-words 3 --max-chars 22 --max-lines 1 --strip-punctuation"

-- Resolve the install dir from the pointer file next to this preset, then
-- hand off to the shared cross-platform launcher (see preset_launch.lua).
local sep = package.config:sub(1, 1)
local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*[/\\])") or ("." .. sep)
local f = io.open(script_dir .. "resolve_whisper_path.txt", "r")
if not f then
    print("[LAB37 " .. LAB37_TOOL .. "] ERROR: pointer file not found (re-run setup)")
    return
end
LAB37_APP_DIR = f:read("*all"):gsub("%s+$", "")
f:close()
local launcher = LAB37_APP_DIR .. sep .. "preset_launch.lua"
local lf = io.open(launcher, "r")
if not lf then
    print("[LAB37 " .. LAB37_TOOL .. "] ERROR: launcher missing at " .. launcher .. " (re-run setup)")
    return
end
lf:close()
dofile(launcher)
