# Assisted install runbook (for Claude / AI agents)

You are helping a **non-technical user** install LAB37 Resolve Whisper
(AI-powered captions for DaVinci Resolve Studio) on their **Mac or Windows
PC**. Your job is to do everything you can yourself (terminal commands, file
checks, reading logs) and only involve the user for GUI clicks or their
password.

## How to behave

- Talk to the user in **plain, friendly language**. No jargon, no raw error
  dumps. If the user writes in Swedish, answer in Swedish.
- Run every command **yourself** (Terminal on Mac, PowerShell on Windows).
  Never ask the user to open a terminal or type commands.
- Verify each phase before moving to the next. Don't assume success.
- The user only ever needs to do three kinds of things: click things in
  DaVinci Resolve, click things in OS dialogs, and type their password when
  the system asks.
- If a step fails twice and you can't resolve it, stop and tell the user to
  message Erik at **hello@lab37.io** with a screenshot. Don't leave the
  machine in a half-configured state without explaining what happened.

## Pick the platform first

Detect the OS and follow the matching install track, then both platforms
share the Resolve-side steps:

- **macOS** → do **Phase 0-3 (Mac)** below.
- **Windows** → do **Phase 0-3 (Windows)** in the "Windows track" section near
  the end, then come back.
- **Both then do** Phase 4 (Resolve settings + health check) and Phase 5
  (first captions), which are identical on both platforms.

---

# Mac track

## Phase 0 - Preflight (all yours, no user action)

Check and report each of these before touching anything:

```
uname -m                      # must be arm64 (Apple Silicon). Intel Macs are NOT supported.
sw_vers -productVersion       # macOS 13+ recommended
df -h ~                       # need ~6 GB free (3 GB model + deps)
ls "/Applications/DaVinci Resolve/DaVinci Resolve.app"   # Resolve must be installed
xcode-select -p               # command line tools (needed for git/brew)
```

- **Intel Mac** → stop. The tool needs Apple Silicon (mlx). Tell the user to
  contact Erik.
- **Resolve missing** → the user must install DaVinci Resolve **Studio**
  first (they should have a LAB37 license key; if unsure, ask Erik). The free
  version will not work - the scripting API is Studio-only. You can't easily
  tell free from Studio on disk; the health check in Phase 4 will expose it.
- **xcode-select fails** → run `xcode-select --install`, tell the user a
  macOS dialog will appear and they should click "Install" (takes a few
  minutes). Wait for it to finish before continuing.

## Phase 1 - Homebrew and Python

```
command -v brew || echo MISSING
```

If missing, install it:

```
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Warn the user first: *"macOS will ask for your computer password - that's
normal, it's how the Mac approves new software. The password won't be shown
as you type."* After install, make brew available in the current shell:

```
eval "$(/opt/homebrew/bin/brew shellenv)"
```

Then Python:

```
python3 --version    # need 3.10+
```

If too old or missing: `brew install python@3.12` and use
`PATH="/opt/homebrew/opt/python@3.12/libexec/bin:$PATH"` for the rest.

## Phase 2 - Get the code

Clone to a **permanent location** - NOT Downloads or Desktop. The install
path gets written into Resolve's scripts folder, so moving the folder later
breaks the tool.

```
mkdir -p ~/Tools
git clone https://github.com/delarc0/resolve-whisper.git ~/Tools/resolve-whisper
```

## Phase 3 - Run setup

```
cd ~/Tools/resolve-whisper && ./setup.sh
```

This creates a Python environment, installs dependencies (mlx-whisper,
silero-vad, ffmpeg via brew if missing), copies the menu entries into
Resolve, and downloads the ~3 GB AI model. Takes 5-15 minutes depending on
wifi. Narrate progress to the user so they know it's working, and read the
output yourself - every section should end in OK.

# Shared steps (Mac and Windows)

## Phase 4 - Resolve settings + health check (user does GUI, you verify)

1. Ask the user to open **DaVinci Resolve**.
2. Scripting permission: guide them to
   **DaVinci Resolve menu > Preferences > System > General >
   "External scripting using"** and set it to **Local**. Then Save and
   **restart Resolve**.
3. Ask them to open any project and any timeline (a scratch/test project is
   fine - ask them to pick one that isn't important).
4. Guide them to **Workspace > Scripts > Edit > LAB37 Check** in the menu bar.
5. Wait ~15 seconds, then read the result yourself:

```
# Mac
cat /tmp/resolve_whisper_check.log
# Windows (PowerShell)
Get-Content "$env:TEMP\resolve_whisper_check.log"
```

Every row must say **PASS**. Common failures:
- `Connect to Resolve - FAIL` → Resolve not running, external scripting not
  set to Local, or the free (non-Studio) version → check Preferences again;
  if scripting is Local and it still fails, it's probably the free version:
  the user needs a Studio license from Erik.
- `mlx_whisper importable - FAIL` (Mac) / `faster_whisper importable - FAIL`
  (Windows) → re-run setup (`./setup.sh` on Mac, `.\setup.ps1` on Windows)
  and read its output.
- No log file at all → the menu entry didn't run; re-run setup so the Lua
  launchers get copied, restart Resolve, try again.

## Phase 5 - First captions (walk the user through it)

1. In a timeline with speech, set **in and out points** (press `I` at the
   start, `O` at the end of the section).
2. **Workspace > Scripts > Edit >** pick one:
   - **LAB37 Reels** - Swedish, 1-3 word captions for 9:16 reels
   - **LAB37 Podcast** - Swedish, full sentences
   - **LAB37 Auto** - auto-detect language, full sentences
   - **LAB37 Custom** - a settings window opens first (language, caption
     style, line length, ALL CAPS). Best choice when the user's content
     doesn't match the presets; it remembers their choices.
3. A small progress window appears. First run takes a couple of minutes
   (model loads into memory); after that ~5 seconds per minute of audio.
4. When it finishes, if the timeline had **no subtitle track**, the captions
   are **placed on a new subtitle track automatically** - the user is done.
5. If the timeline **already had subtitle tracks** (or placement failed),
   the SRT is in the **Media Pool** instead: have the user right-click it
   there and choose **Insert Selected Subtitles to Timeline**.
6. **IMPORTANT: never drag SRT files from Finder/Explorer onto the timeline.**
   Resolve 21.0.2 has a crash bug in that path. Always use the Media Pool
   right-click route. Warn the user about this explicitly.
7. Captions come out in ALL CAPS by design. Font, size, and outline are set
   **once per timeline** by the user: select the subtitle track, open the
   **Inspector > Track Style** tab, and style it there. Every caption on the
   track follows, including re-generations.

Done - tell the user that steps 1-4 are the whole everyday workflow and that
SRT files also save to their Desktop in a `Captions` folder.

# Windows track

Do these instead of the Mac Phases 0-3, then go back to the shared Phase 4.
Run everything in **PowerShell** yourself.

## Phase 0 (Windows) - Preflight

```powershell
[Environment]::Is64BitOperatingSystem            # must be True
(Get-CimInstance Win32_OperatingSystem).Caption  # Windows 10/11
Get-PSDrive C | Select-Object Free               # need ~6 GB free
Test-Path "C:\Program Files\Blackmagic Design\DaVinci Resolve\Resolve.exe"  # Resolve installed?
python --version                                 # or: py --version
```

- **Resolve missing** → the user installs DaVinci Resolve **Studio** first
  (free version has no scripting API). The Phase 4 health check will expose
  free-vs-Studio.
- **Python missing / older than 3.10** → install Python 3.10+ from python.org
  and tell the user to tick **"Add python.exe to PATH"** in the installer.
  You can't do the GUI install for them; guide them, then continue.
- Windows runs on **CPU** for now (GPU is a planned follow-up), so it's slower
  than Mac on long timelines. That's expected, not a fault.

## Phase 1 (Windows) - Get the code

Clone to a **permanent location** (not Downloads/Desktop; the path is written
into Resolve's scripts folder, so moving it later breaks the tool):

```powershell
git clone https://github.com/delarc0/resolve-whisper "$env:USERPROFILE\Tools\resolve-whisper"
```

If `git` is missing, install it (`winget install --id Git.Git -e`) or download
the ZIP from GitHub and expand it to that folder.

## Phase 2 (Windows) - Run setup

```powershell
cd "$env:USERPROFILE\Tools\resolve-whisper"
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

This creates the Python environment, installs faster-whisper (CPU mode),
copies the menu entries into Resolve, and downloads the ~3 GB model. Narrate
progress; every section should end in OK. If ffmpeg was just installed via
winget, open a fresh PowerShell so it's on PATH before continuing.

## Phase 3 (Windows) - none

There's no separate step 3; once setup.ps1 finishes cleanly, go to **Phase 4**
(shared). The scripting-permission setting and health check are identical.

---

## Troubleshooting quick reference

| Symptom | Fix |
|---|---|
| `permission denied: ./setup.sh` (Mac) | `chmod +x setup.sh` and retry |
| `setup.ps1 cannot be loaded` (Windows) | run it with `powershell -ExecutionPolicy Bypass -File .\setup.ps1` |
| brew not found after install (Mac) | `eval "$(/opt/homebrew/bin/brew shellenv)"` |
| `python` not recognized (Windows) | Python not on PATH; reinstall from python.org with "Add to PATH" ticked, or use `py` |
| ffmpeg not found (Windows) | `winget install --id Gyan.FFmpeg -e`, then open a new PowerShell |
| tkinter warning during setup (Mac) | fine - setup installs python-tk automatically; only cosmetic if it persists |
| Model download fails | check wifi, re-run setup (it resumes) |
| "Another caption run is in progress" | previous run still working; wait or cancel from its progress window |
| Resolve crashes on SRT drag | known Resolve 21.0.2 bug - use Media Pool right-click route, never drag |
| "'Audio Only' preset not found" | run LAB37 Check (creates it); on Resolve 21 it's built in |
| Wrong language in captions | use the LAB37 Auto preset, or tell the user to always pick the preset matching the content language |

## Notes

- Logs (read them yourself before asking the user anything):
  - Mac: `/tmp/resolve_whisper.log` (runs), `/tmp/resolve_whisper_check.log`
    (health check)
  - Windows: `%TEMP%\resolve_whisper.log` and `%TEMP%\resolve_whisper_check.log`
- Config knobs: `caption_config.json` in the install folder (created on
  first run). Defaults are correct for LAB37 work - don't change them unless
  asked.
- After any DaVinci Resolve update: re-run setup (`./setup.sh` on Mac,
  `.\setup.ps1` on Windows), then LAB37 Check.
