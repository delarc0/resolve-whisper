# LAB37 TOOLS: Resolve Whisper

AI-powered captions for DaVinci Resolve Studio. Runs locally on your GPU. No cloud, no subscription, no waiting.

Built by LAB37 // lab37.se

---

## Install (one time)

### With an AI assistant (recommended for non-technical users)

If you have Claude (or another AI assistant that can run commands on your
Mac), paste this and let it do the work:

> Please install the LAB37 Resolve Whisper caption tool on this Mac.
> Clone https://github.com/delarc0/resolve-whisper and follow the runbook in
> AGENT_INSTALL.md exactly. I'm not technical, so do everything you can
> yourself and guide me in simple steps when you need me.

### Mac (Apple Silicon)
1. Make sure Python 3 is installed (`brew install python` if not)
2. Open Terminal in this folder and run: `./setup.sh`
3. Wait for it to finish (downloads a ~3 GB AI model on first run)

Setup installs everything and adds the scripts to Resolve automatically.

---

## Use

1. Open **DaVinci Resolve Studio** and select a timeline
2. Set your **in and out points** (I and O keys) for the section you want captioned
3. Go to **Workspace > Scripts > Edit** and pick a preset:
   - **LAB37 Reels** - single-line Swedish captions, 1-3 words, no punctuation (9:16 reels)
   - **LAB37 Podcast** - plain Swedish SRT, full sentences
   - **LAB37 Auto** - auto-detect language, plain SRT
   - **LAB37 Check** - health check, run after install or a Resolve update
4. Wait for the progress window to finish. If the timeline has no subtitle
   track yet, the captions are **placed on a new subtitle track
   automatically** - you're done.
5. If the timeline already has subtitle tracks (e.g. you're re-generating),
   the SRT lands in the Media Pool instead: right-click it there >
   **Insert Selected Subtitles to Timeline**.

Never drag SRT files from Finder onto the timeline - Resolve 21.0.2 has a
crash bug in its drag handler (worst on collaboration projects with locked
timelines). Auto-placement can also fail on collaboration timelines locked
by another editor; the Media Pool right-click route works once you hold the
lock.

## Caption style (font, size, stroke)

Resolve's subtitle styling has no scripting API, so the tool can't set it
for you. Style each timeline **once** and it sticks:

1. After the first generation, select the subtitle track
2. Inspector > **Track Style**: set font (e.g. Montserrat), size, stroke
3. Every caption on that track updates, and future re-generations onto the
   same track inherit the style automatically

ALL CAPS is handled by the tool itself (`"uppercase": true` in the config,
on by default) since Resolve has no caps transform.

Each run writes a new file named `<timeline> YYYYMMDD-HHMMSS.srt` so Resolve
never re-imports a stale cached version. The tool's own files older than 30
days are cleaned up automatically (set `keep_srt_days` to 0 in the config to
disable); other SRT files in the folder are never touched.

---

## Settings

After running once, a `caption_config.json` file appears in the install folder. Open it in any text editor to change:

| Setting | Default | What it does |
|---|---|---|
| `language` | `"sv"` | Transcription language. `null` = auto-detect. Note: presets pass their own language flag (Reels/Podcast force `sv`, Auto forces auto-detect), so this mainly applies to `--file` mode. |
| `min_word_probability` | `0.0` | Drop words below this Whisper confidence (0 = off). Try `0.3` on noisy footage if junk words appear. |
| `max_chars_per_line` | `42` | Characters per subtitle line |
| `max_lines` | `1` | Lines per subtitle block |
| `min_duration_s` | `1.0` | Shortest a subtitle stays on screen (seconds) |
| `max_duration_s` | `7.0` | Longest a subtitle stays on screen (seconds) |
| `keep_srt_days` | `30` | Auto-delete this tool's old SRTs from the output folder (0 = never) |
| `uppercase` | `true` | ALL CAPS caption text (Resolve has no caps transform) |

Most people won't need to change anything.

---

## Health check

Run this once after install and after any DaVinci Resolve update:

**From Resolve:** Workspace > Scripts > Edit > **LAB37 Check**
**Or terminal:** `./.venv/bin/python caption.py --check`

It validates the Python environment, ffmpeg, the VAD model, the Resolve
connection, and the `Audio Only` render preset (creating it if missing).
Output goes to `/tmp/resolve_whisper_check.log` when launched from Resolve.

## Troubleshooting

Run **LAB37 Check** first - it catches almost everything below automatically.

**Script doesn't appear in Workspace > Scripts**
Re-run `./setup.sh`. It copies the scripts to Resolve's scripts folder.

**"Python environment not found" error**
Re-run setup. The virtual environment may not have been created.

**Transcription takes forever**
First run is slower because it loads the AI model into memory. After that, expect roughly 5 seconds per minute of audio.

**Captions are in the wrong language**
Use the **LAB37 Auto** preset (auto-detects), or set `"language"` in `caption_config.json` for `--file` mode.

**Nothing happens when I click the script**
Make sure you have a timeline selected (not just a project open). The script needs an active timeline to work with.

**Resolve crashes when dragging the SRT into the timeline**
Known Resolve 21.0.2 bug (segfault in the drag handler, especially on cloud
collaboration projects where another machine holds timeline locks). Don't
drag from Finder. The tool imports the SRT into the Media Pool for you:
right-click it there > **Insert Selected Subtitles to Timeline**. Also make
sure no collaborator has the timeline open so you hold the lock.

**"Another caption run is in progress"**
A previous run is still working (or its progress window is open). Wait for it
or cancel it from its progress window, then try again.

**"'Audio Only' render preset not found"**
Run **LAB37 Check** - it'll create the preset for you.

**"'Audio Only' preset is misconfigured"**
The preset was edited (probably exported as video by accident). Recreate it:
`./.venv/bin/python create_audio_only_preset.py --force`

**Render starts but never completes / log shows "Render timed out"**
Check `/tmp/resolve_whisper.log`. If the render genuinely got stuck (not just slow), it usually means another job is queued and the render queue got confused. Open the Deliver page, clear stuck jobs, and run again.

**Resolve scripts menu vanished after a Resolve update**
Re-run setup so the Lua launchers get re-copied to the new Scripts folder, then run **LAB37 Check** to confirm.

---

## Advanced: command line

You can also use it outside of Resolve to caption any audio/video file:

```
.venv/bin/python caption.py --file video.mp4
.venv/bin/python caption.py --file interview.wav --language sv
.venv/bin/python caption.py --file clip.mov --language auto
```

SRT file appears next to the source file.

---

## Requirements

- DaVinci Resolve Studio 18+ (scripting API is Studio only; Resolve 21 supported)
- Python 3.10+
- Apple Silicon Mac (the mlx Whisper backend is Metal-only)
- ~3 GB disk space (AI model + dependencies)
- ~2-3 GB VRAM during transcription

---

LAB37 TOOLS // lab37.se
