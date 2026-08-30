-- LAB37 Update -- pulls the latest version and re-runs setup
--
-- Deliberately NOT routed through preset_launch.lua: that runs caption.py
-- detached and silent, which is right for a caption job and wrong here. An
-- update rewrites the menu entries and reinstalls dependencies, so the user
-- needs to SEE it happen and see any error. This opens a real terminal
-- window running the update script.
local TOOL = "Update"

local sep = package.config:sub(1, 1)
local is_win = sep == "\\"
local script_dir = debug.getinfo(1, "S").source:sub(2):match("(.*[/\\])") or ("." .. sep)

local function say(msg) print("[LAB37 " .. TOOL .. "] " .. msg) end

local f = io.open(script_dir .. "resolve_whisper_path.txt", "r")
if not f then
    say("ERROR: pointer file not found (re-run setup)")
    return
end
local app_dir = f:read("*all"):gsub("%s+$", "")
f:close()

local script = app_dir .. sep .. (is_win and "update.ps1" or "update.sh")
local sf = io.open(script, "r")
if sf then sf:close() else
    say("ERROR: update script not found at " .. script)
    say("This install predates it; ask Erik for the one-time update steps.")
    return
end

say("Opening a terminal to run the update...")
say("Quit and reopen Resolve when it finishes.")

local cmd
if is_win then
    -- -NoExit leaves the window open so the user can read the result.
    cmd = string.format(
        [[start "" powershell -NoExit -ExecutionPolicy Bypass -File "%s"]],
        script)
else
    -- Single-quote for sh: %q escapes for Lua and would leave $ and ` live.
    local function shq(s)
        return "'" .. tostring(s):gsub("'", "'\\''") .. "'"
    end
    -- `open -a Terminal <script>` runs it in a visible window, detached from
    -- Resolve. Running it inside Resolve's process would freeze the UI for
    -- the length of a dependency install.
    cmd = string.format([[open -a Terminal %s]], shq(script))
end

os.execute(cmd)
