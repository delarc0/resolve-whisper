# resolve-whisper

Local AI captions for DaVinci Resolve Studio (a LAB37 tool). Renders timeline
audio via the render queue, transcribes locally with Whisper (mlx-whisper on
Apple Silicon, faster-whisper on Windows), writes an SRT, and imports it into
Resolve's Media Pool.

## Installing this for a user?

Follow **AGENT_INSTALL.md** — a step-by-step runbook written for AI agents
assisting non-technical users. Human-facing instructions live in README.md.

## Architecture (Mac is the primary platform)

- `presets/*.lua` — Resolve menu entries (Workspace > Scripts > Edit);
  self-contained launchers that shell out to caption.py with preset flags
- `caption.py` — the pipeline: render timeline audio ("Audio Only" preset) →
  transcribe → SRT → auto-import to Media Pool. Also `--check` (pre-flight
  health check) and `--file` (standalone, no Resolve) modes
- `transcribe.py` — Whisper wrapper + silero-vad pre-pass (prevents
  word-timestamp drift after pauses)
- `srt.py` — word→caption chunking (pause/punctuation logic, Reels mode) and
  SRT generation
- `progress_ui.py` — Tk progress window; talks to caption.py via an atomic
  JSON status file; Cancel delivers SIGTERM
- `config.py` — defaults + `caption_config.json` generation
- `resolve_script.py`, `launcher_stub.py`, `setup.bat` — Windows path (secondary)

## Development

- Tests: `./.venv/bin/python -m unittest discover -s tests` — pure logic,
  no Resolve required. Run after any change.
- Live verification needs Resolve Studio running: `caption.py --check`
  (or Workspace > Scripts > Edit > LAB37 Check).
- Resolve API version gotchas (21.x job-key renames, the factory Audio Only
  preset, the SRT drag-crash bug) are handled by compat shims and documented
  inline where they live. Don't remove them; Resolve 18-21 must all work.
- Never make captions import via Finder drag — Resolve 21.0.2 segfaults in
  its drag handler. The Media Pool route is deliberate.
