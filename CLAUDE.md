# resolve-whisper

Local AI captions for DaVinci Resolve Studio (a LAB37 tool). Renders timeline
audio via the render queue, transcribes locally with Whisper (mlx-whisper on
Apple Silicon, faster-whisper on Windows), writes an SRT, and imports it into
Resolve's Media Pool.

## Installing this for a user?

Follow **AGENT_INSTALL.md** - a step-by-step runbook written for AI agents
assisting non-technical users. Human-facing instructions live in README.md.

## Architecture (Mac is the primary platform)

- `presets/*.lua` - Resolve menu entries (Workspace > Scripts > Edit);
  self-contained launchers that shell out to caption.py with preset flags
- `caption.py` - the pipeline: render timeline audio ("Audio Only" preset) →
  transcribe → SRT → auto-import to Media Pool. Also `--check` (pre-flight
  health check) and `--file` (standalone, no Resolve) modes
- `transcribe.py` - Whisper wrapper + silero-vad pre-pass (prevents
  word-timestamp drift after pauses)
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
- `config.py` - defaults + `caption_config.json` generation

Mac-only. The old Windows path (resolve_script.py + launcher_stub.py +
setup.bat) was removed 2026-07-15: it ran Tk dialogs inside Resolve's process
which freezes Resolve on macOS, and it had drifted from the hardened
pipeline. Don't resurrect it.

Caption placement: SRT is imported to the Media Pool and placed on a fresh
subtitle track via `AppendToTimeline` - which returns None even on success
(verify by counting track items) and fails when another collaborator holds
the timeline lock. Subtitle track STYLE (font/size/stroke) has no scripting
API at all; users style the track once per timeline in the Inspector. ALL
CAPS therefore lives in the text (`uppercase` config).

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
