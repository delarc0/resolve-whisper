# resolve-whisper

Local AI captions for DaVinci Resolve Studio (a LAB37 tool). Renders timeline
audio via the render queue, transcribes locally with Whisper (mlx-whisper on
Apple Silicon, faster-whisper on Windows), writes an SRT, and imports it into
Resolve's Media Pool.

Cross-platform (Mac + Windows), one codebase. All OS differences route
through `platforminfo.py`; the Mac and Windows transcription backends split
at a single branch in `transcribe.py` and share no fragile code, so a change
to one platform can't destabilize the other. Mac is the primary/most-tested
platform.

## Installing this for a user?

Follow **AGENT_INSTALL.md** - a step-by-step runbook written for AI agents
assisting non-technical users. Human-facing instructions live in README.md.

## Architecture (Mac is the primary platform)

- `presets/*.lua` - Resolve menu entries (Workspace > Scripts > Edit); thin:
  each sets tool/args, reads the pointer file for the install dir, then
  dofile()s the shared launcher
- `preset_launch.lua` - the one cross-platform launcher (lives in the app
  dir, NOT Scripts/Edit, so it isn't its own menu entry). Runs caption.py
  detached with the preset's args; no env exports (caption.py bootstraps the
  Resolve module path itself), so no fragile per-OS shell quoting
- `platforminfo.py` - single OS boundary: flags, Resolve scripts/API/lib
  paths, ffmpeg/ffprobe finder, and `bootstrap_resolve_env()` which sets
  RESOLVE_SCRIPT_* and adds the Modules dir to sys.path
- `caption.py` - the pipeline: render timeline audio ("Audio Only" preset) →
  transcribe → SRT → auto-import to Media Pool. Also `--check` (pre-flight
  health check), `--file` (standalone, no Resolve), and `--dialog` (Custom
  preset settings window) modes
- `transcribe.py` - Whisper wrapper. Backend splits at ONE branch: mlx +
  silero-vad pre-pass (Mac) vs faster-whisper with its built-in VAD
  (Windows/other). The silero pre-pass and its ffmpeg/AAC decode fallback
  are Mac-path-only
- `srt.py` - word→caption chunking and SRT generation. Two chunkers: greedy
  (Reels, max_words > 0) and balanced (Auto/Podcast, max_words == 0) which
  splits at sentence/silence boundaries then picks break points by cost
  minimization (even fill, connector penalty, pause bonus) to avoid orphan
  fragments
- `progress_ui.py` - Tk progress window; talks to caption.py via an atomic
  JSON status file; Cancel delivers SIGTERM
- `settings_ui.py` - settings dialog for the Custom preset (caption.py
  --dialog spawns it); persists last-used choices to custom_settings.json,
  returns them via a SETTINGS: JSON line on stdout
- `ui_theme.py` - shared Tk theme: apps/DESIGN.md charcoal tokens plus
  Aqua-safe widgets (macOS ignores colors on native tk.Button/OptionMenu,
  so buttons/selects/toggles are Label/Frame-based). Verify UI changes with
  screenshots; Aqua renders wrong silently
- `config.py` - defaults, `caption_config.json` generation, per-key type
  validation (a bad value used to crash AFTER the render), and
  `model_for_language()`: Swedish routes to KB-Whisper (KBLab, ~47% lower
  WER on Swedish), everything else to large-v3. Mac has no MLX build of
  KB-Whisper yet and falls back automatically

Cross-platform via one codebase (see the isolation note up top). Windows
install is `setup.ps1`; Mac is `setup.sh`. The OLD Windows path
(resolve_script.py + launcher_stub.py + setup.bat) was removed 2026-07-15
because it ran Tk dialogs INSIDE Resolve's process (froze Resolve) and had
drifted from the hardened pipeline. Don't resurrect that in-process design;
the current Windows support reuses the same detached-subprocess pipeline as
Mac, which is why it can't drift. Windows-specific gotcha: Cancel sends
SIGTERM, which is a hard kill on Windows (no cleanup finally blocks), so a
cancelled Windows run may leave a render job behind - acceptable for now.

Caption placement: SRT is imported to the Media Pool and placed on a fresh
subtitle track via `AppendToTimeline` - which returns None even on success
(verify by counting track items) and fails when another collaborator holds
the timeline lock. Subtitle track STYLE (font/size/stroke) has no scripting
API at all; users style the track once per timeline in the Inspector. ALL
CAPS therefore lives in the text (`uppercase` config).

## Invariants worth not breaking

These were each a real bug found in the 2026-07 audit; the comments at each
site explain why. Summary so they aren't "simplified" away:

- **The run lock is a kernel lock** (flock/msvcrt) on a held-open fd, not a
  file's existence. That is what survives SIGKILL. Do not go back to
  PID-liveness probing: `os.kill(pid, 0)` TERMINATES on Windows.
- **The status file is per-run** (`...status.<pid>.json`). A shared path let
  a lingering window adopt the next run's pid and cancel it.
- **Caption placement offset**: `SelectAllFrames` wins over MarkIn, and a
  duration cross-check guards the rest. Shifting captions wrongly is worse
  than not shifting them, so anything ambiguous returns 0.
- **Never fuse or drop caption text**: punctuation removal replaces with a
  space (not nothing) and protects decimals - "3,5 miljoner" must not become
  "35 miljoner". Line overflow is never truncated; losing a word is worse
  than an extra line.
- **Timing is normalised at the WORD level** before chunking, and the
  zero-duration guard compares MILLISECONDS (the SRT's resolution).
- **SRT cleanup only sweeps the tool's own default folder**, and logs every
  filename it deletes.

## Development

- Tests: `./.venv/bin/python -m unittest discover -s tests` - pure logic,
  no Resolve required. Run after any change.
- Live verification needs Resolve Studio running: `caption.py --check`
  (or Workspace > Scripts > Edit > LAB37 Check).
- Resolve API version gotchas (21.x job-key renames, the factory Audio Only
  preset, the SRT drag-crash bug) are handled by compat shims and documented
  inline where they live. Don't remove them; Resolve 18-21 must all work.
- Never make captions import via Finder drag - Resolve 21.0.2 segfaults in
  its drag handler. The Media Pool route is deliberate.
