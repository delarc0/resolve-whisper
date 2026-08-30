-- LAB37 Custom -- opens a settings dialog before running
--
-- Resolve's Lua sandbox does NOT pass globals into a dofile()'d chunk, so the
-- tool name/args/install dir are handed to the launcher as ARGUMENTS. Setting
-- them as LAB37_* globals looked right and silently launched nothing.
local TOOL = "Custom"
local ARGS = "--dialog"

-- Resolve the install dir from the pointer file next to this preset, then
-- hand off to the shared cross-platform launcher (see preset_launch.lua).
local sep = package.config:sub(1, 1)
local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*[/\\])") or ("." .. sep)
local f = io.open(script_dir .. "resolve_whisper_path.txt", "r")
if not f then
    print("[LAB37 " .. TOOL .. "] ERROR: pointer file not found (re-run setup)")
    return
end
local app_dir = f:read("*all"):gsub("%s+$", "")
f:close()

local launcher = app_dir .. sep .. "preset_launch.lua"
local lf = io.open(launcher, "r")
if not lf then
    print("[LAB37 " .. TOOL .. "] ERROR: launcher missing at " .. launcher .. " (re-run setup)")
    return
end
lf:close()

local launch = dofile(launcher)
if type(launch) ~= "function" then
    -- Launcher is older than this preset (pulled without re-running setup).
    print("[LAB37 " .. TOOL .. "] ERROR: launcher is out of date; run ./update.sh (Mac) or update.ps1 (Windows)")
    return
end
launch(TOOL, ARGS, app_dir)
