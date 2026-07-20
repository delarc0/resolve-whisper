-- Shared cross-platform launcher for the LAB37 Whisper presets.
--
-- Lives in the app dir (NOT Resolve's Scripts/Edit folder) so it never shows
-- up as its own menu entry. Each preset sets LAB37_TOOL + LAB37_ARGS +
-- LAB37_APP_DIR and then dofile()s this. Keeping the launch logic in one
-- place means a fix lands once, not five times.
--
-- caption.py finds Resolve's scripting module itself (platforminfo
-- .bootstrap_resolve_env), so this launcher only has to run Python detached
-- with the preset's args. No env exports = no fragile per-OS shell quoting.
--
-- Detached on purpose: the pipeline must not run inside Resolve's process
-- (Tk-in-process froze Resolve on Mac, and a long transcription would block
-- the UI regardless).

local TOOL = LAB37_TOOL or "Whisper"
local ARGS = LAB37_ARGS or ""
local app_dir = LAB37_APP_DIR

local function say(msg) print("[LAB37 " .. TOOL .. "] " .. msg) end

if not app_dir or app_dir == "" then
    say("ERROR: app dir not resolved (re-run setup)")
    return
end

local is_win = package.config:sub(1, 1) == "\\"

local PY, CAPTION, LOGFILE
if is_win then
    PY = app_dir .. "\\.venv\\Scripts\\python.exe"
    CAPTION = app_dir .. "\\caption.py"
    -- `start /b` redirection to a real file is unreliable for a detached
    -- child, so discard shell output to NUL; caption.py writes the run log
    -- itself (a Python FileHandler) on Windows, which is dependable.
    LOGFILE = "NUL"
else
    PY = app_dir .. "/.venv/bin/python3"
    CAPTION = app_dir .. "/caption.py"
    LOGFILE = "/tmp/resolve_whisper.log"
end

say("Starting...")

-- Fail loudly if the venv python is missing (setup not run / folder moved);
-- the subprocess would otherwise die silently into the log.
local pf = io.open(PY, "r")
if pf then pf:close() else
    say("ERROR: Python venv not found at: " .. PY)
    say("Run setup again (the install folder may have moved).")
    return
end

local cmd
if is_win then
    -- os.execute wraps this in `cmd /c "..."`, so keep it a plain command;
    -- start /b runs it detached. Inner quotes guard spaces in the paths.
    cmd = string.format(
        [[start /b "" "%s" "%s" %s > "%s" 2>&1]],
        PY, CAPTION, ARGS, LOGFILE
    )
else
    -- Homebrew ffmpeg lives outside Resolve's PATH; export it so mlx_whisper
    -- and the ffprobe fallback can find it. Trailing & detaches.
    cmd = string.format(
        [[(export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"; ]] ..
        [[cd %q && %q %q %s) > %q 2>&1 &]],
        app_dir, PY, CAPTION, ARGS, LOGFILE
    )
end

say("Launching subprocess.")
os.execute(cmd)
