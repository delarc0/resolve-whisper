# LAB37 TOOLS: Resolve Whisper

AI-powered captions for DaVinci Resolve Studio. Runs locally on your GPU. No cloud, no subscription, no waiting.

Built by LAB37 // lab37.se

---

## Install (one time)

### Windows
1. Make sure Python 3.10+ is installed ([python.org](https://python.org) - check "Add to PATH")
2. Double-click **`setup.bat`**
3. Wait for it to finish (downloads ~1.5 GB AI model on first run)

### Mac
1. Make sure Python 3 is installed (`brew install python` if not)
2. Open Terminal in this folder and run: `./setup.sh`
3. Wait for it to finish

Setup installs everything and adds the script to Resolve automatically.

---

## Use

1. Open **DaVinci Resolve Studio** and select a timeline
2. Set your **in and out points** (I and O keys) for the section you want captioned
3. Go to **Workspace > Scripts > LAB37 Resolve Whisper**
4. Wait for the progress window to finish
5. A Captions folder opens on your Desktop with the `.srt` file
6. In Resolve: **File > Import > Subtitle** and select the `.srt`

Captions appear on a subtitle track in your timeline.

---

## Settings

After running once, a `caption_config.json` file appears in the install folder. Open it in any text editor to change:

| Setting | Default | What it does |
|---|---|---|
| `language` | `null` (auto) | Force a language: `"sv"`, `"en"`, `"fi"`, etc. |
| `max_chars_per_line` | `42` | Characters per subtitle line |
| `max_lines` | `2` | Lines per subtitle block |
| `min_duration_s` | `1.0` | Shortest a subtitle stays on screen (seconds) |
| `max_duration_s` | `7.0` | Longest a subtitle stays on screen (seconds) |
| `highlight_low_confidence` | `true` | Marks words the AI was unsure about with `[brackets?]` |

Most people won't need to change anything. If transcription quality is bad, try setting `language` to your language instead of auto-detect.

---

## Troubleshooting

**Script doesn't appear in Workspace > Scripts**
Re-run `setup.bat` / `setup.sh`. It copies the script to Resolve's scripts folder.

**"Python environment not found" error**
Re-run setup. The virtual environment may not have been created.

**Transcription takes forever**
First run is slower because it loads the AI model into memory. After that, expect roughly 5 seconds per minute of audio.

**Captions are in the wrong language**
Open `caption_config.json` and set `"language": "sv"` (or your language code) instead of `null`.

**Nothing happens when I click the script**
Make sure you have a timeline selected (not just a project open). The script needs an active timeline to work with.

---

## Advanced: command line

You can also use it outside of Resolve to caption any audio/video file:

```
.venv\Scripts\python caption.py --file video.mp4
.venv\Scripts\python caption.py --file interview.wav --language sv
```

SRT file appears next to the source file.

---

## Requirements

- DaVinci Resolve Studio 18+ (scripting API is Studio only)
- Python 3.10+
- NVIDIA GPU with CUDA (Windows) or Apple Silicon (Mac)
- ~3 GB disk space (AI model + dependencies)
- ~2-3 GB VRAM during transcription

---

LAB37 TOOLS // lab37.se
