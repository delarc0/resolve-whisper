# PRD: Resolve Captions - AI Transcription for DaVinci Resolve

## Problem

DaVinci Resolve's built-in transcription is slow, inaccurate, and painful. It costs LAB37 multiple hours per week fixing bad captions. We already have a proven Whisper pipeline in Bark that runs locally on GPU - we just need to point it at timeline audio instead of a microphone.

## Solution

A DaVinci Resolve script that extracts audio from the current timeline, runs it through faster-whisper (same model Bark uses), and generates an SRT file with word-level timestamps. Optionally imports it back into Resolve's subtitle track.

## What We Reuse from Bark

| Bark module | What we take | Changes needed |
|---|---|---|
| `transcriber.py` | WhisperModel init, faster-whisper/mlx-whisper backends, VAD filtering | Add `word_timestamps=True`, return segment timing data instead of flat string |
| `clean_text()` | Filler word removal, hallucination filter | None - works as-is |
| `config.py` pattern | JSON config with defaults, load/save | New config keys for caption preferences |
| Model choice | `deepdml/faster-whisper-large-v3-turbo-ct2` (Win), `mlx-community/whisper-large-v3-turbo` (Mac) | Same models |

**Not reused:** overlay.py, audio.py, keyboard_hook.py, feedback.py, tray.py - all real-time dictation stuff.

## Architecture

```
resolve-captions/
  caption.py          # Main script - Resolve API integration + orchestration
  transcribe.py       # Whisper engine (extracted/adapted from Bark's transcriber.py)
  srt.py              # SRT file generation from word-level timestamps
  config.py           # Script config (language, max chars per line, etc.)
  caption_config.json # User settings (auto-created)
```

Runs as a Resolve Utility script (available on all pages) or standalone CLI.

## Core Workflow

```
1. User runs script from Resolve's Scripts menu (or CLI)
2. Script gets current timeline from Resolve API
3. Extracts audio:
   Option A: Render timeline audio as WAV via Resolve render API
   Option B: Read source media file paths from timeline items, extract audio via ffprobe/ffmpeg (faster)
4. Whisper transcribes with word_timestamps=True
5. Script generates SRT with configurable line length / timing
6. SRT file saved next to project or in configured output dir
7. User imports SRT via File > Import > Subtitle (or auto-import if API supports it)
```

## Transcription Engine

Adapted from Bark's `transcriber.py`:

```python
segments, info = model.transcribe(
    audio_path,                    # WAV file path (not numpy array)
    beam_size=5,
    language=config["language"],   # None = auto-detect
    word_timestamps=True,          # THE KEY ADDITION - gives us per-word timing
    vad_filter=True,
    vad_parameters=dict(
        min_silence_duration_ms=500,
        speech_pad_ms=200,
    ),
)

# Each segment has:
#   segment.start, segment.end (seconds)
#   segment.text
#   segment.words[] -> each word has .word, .start, .end, .probability
```

Word-level timestamps let us do tight 1-2 word captions (TikTok/Reels style) or traditional 2-line subtitles.

## SRT Generation

Configurable caption styles:

### Style 1: Traditional subtitles (default)
- Max 42 chars per line, max 2 lines
- Group words into natural phrases using punctuation + timing gaps
- Min duration: 1s, max duration: 7s
- 2-frame gap between subtitles (for Resolve rendering)

### Style 2: Short-form / social (TikTok/Reels)
- Max 3-5 words per caption
- Aggressive timing - each caption appears with the spoken word
- Single line, centered
- Good for vertical video / social cuts

### Style 3: Word-by-word (karaoke)
- One word at a time
- Exact word timestamps
- For kinetic typography / motion graphics workflows

Output example (Style 1):
```srt
1
00:00:01,200 --> 00:00:03,800
Vi har jobbat med det har
projektet i tre manader

2
00:00:04,100 --> 00:00:06,500
och resultatet blev helt
fantastiskt
```

## Config

```json
{
    "language": null,
    "caption_style": "traditional",
    "max_chars_per_line": 42,
    "max_lines": 2,
    "min_duration_s": 1.0,
    "max_duration_s": 7.0,
    "gap_frames": 2,
    "output_dir": null,
    "highlight_low_confidence": true,
    "confidence_threshold": 0.7,
    "beam_size": 5
}
```

`highlight_low_confidence`: Marks words with probability < threshold in output so you know what to double-check. Saves review time.

## Resolve API Integration

### Getting timeline info
```python
import DaVinciResolveScript as bmd
resolve = bmd.scriptapp('Resolve')
project = resolve.GetProjectManager().GetCurrentProject()
timeline = project.GetCurrentTimeline()
fps = timeline.GetSetting("timelineFrameRate")
```

### Audio extraction (two approaches)

**Approach A - Render from timeline (reliable, slower):**
```python
project.SetRenderSettings({
    "ExportAudio": True,
    "ExportVideo": False,
    "TargetDir": temp_dir,
    "CustomName": "caption_audio",
    "AudioCodec": "LinearPCM",
    "AudioBitDepth": 16,
    "AudioSampleRate": 16000,
    "Format": "wav"
})
project.AddRenderJob()
project.StartRendering()
# Poll until render complete, then transcribe the WAV
```

**Approach B - Read source media directly (fast, for simple timelines):**
```python
# Get media file paths from timeline items
for i in range(timeline.GetTrackCount("audio")):
    items = timeline.GetItemListInTrack("audio", i + 1)
    for item in items:
        media = item.GetMediaPoolItem()
        filepath = media.GetClipProperty("File Path")
        # Extract audio with ffmpeg: ffmpeg -i filepath -ar 16000 -ac 1 -f wav output.wav
```

**Recommendation:** Start with Approach A (render). It handles all edge cases (multiple clips, effects, speed changes). Approach B is a future optimization.

### Subtitle import

Resolve's scripting API has no `ImportSubtitle()` method. Options:

1. **Manual import** (v1): Script generates SRT, prints path, user does File > Import > Subtitle. Simple and reliable.
2. **Clipboard + hotkey** (v1.5): Copy SRT path to clipboard, user pastes in import dialog.
3. **Programmatic subtitle creation** (v2): Use `timeline.CreateSubtitlesFromSRT()` if available in newer Resolve versions, or create Text+ clips on a track via API.

**v1 ships with manual import.** It's one extra click and eliminates a class of API compatibility bugs.

## UI / UX

### Option A: Resolve Script Menu (recommended for v1)
- User clicks: Workspace > Scripts > resolve-captions
- Console window shows progress: "Extracting audio... Transcribing... Done!"
- Opens file explorer to SRT location when done

### Option B: Standalone CLI
```bash
python caption.py                          # Auto-detect from Resolve
python caption.py --file video.mp4         # Direct file input
python caption.py --language sv            # Force Swedish
python caption.py --style social           # Short-form captions
python caption.py --style karaoke          # Word-by-word
```

CLI mode is useful for batch processing and automation.

### Option C: Simple tkinter GUI (v2)
- Small window with: language dropdown, style selector, "Transcribe" button, progress bar
- Preview of generated captions before export
- "Import to Resolve" button

**v1: Both Script Menu + CLI. v2: Add GUI.**

## Platform Support

| | Windows | Mac |
|---|---|---|
| Whisper backend | faster-whisper + CUDA | mlx-whisper (Metal) |
| Resolve API | DaVinciResolveScript module | Same |
| Audio extraction | Resolve render API | Same |
| Python | 3.10+ (match Resolve's Python) | Same |

Same dual-platform approach as Bark. Model loading code is already cross-platform.

## Performance Estimates

Based on Bark's model performance:

| Content length | Transcription time (RTX GPU) | Transcription time (M-series Mac) |
|---|---|---|
| 1 min | ~3-5s | ~5-8s |
| 5 min | ~15-25s | ~25-40s |
| 30 min | ~90-150s | ~150-240s |
| 1 hour | ~3-5 min | ~5-8 min |

Plus ~10-30s for audio extraction via Resolve render. Still massively faster than Resolve's built-in.

## Confidence Flagging

faster-whisper gives per-word probability scores. We use these to:

1. Flag low-confidence words in SRT output (e.g., wrap in brackets or add comment)
2. Generate a "review list" of timestamps where the model was unsure
3. Let the editor jump straight to problem spots instead of watching the whole thing

Example flagged output:
```srt
3
00:00:07,200 --> 00:00:09,100
vi anvander [Kubernetes?] for
att hantera deployments
```

## Installation

### As Resolve Script (recommended)
```
1. Clone/copy resolve-captions/ to Resolve's Scripts/Utility folder
2. Install Python deps: pip install faster-whisper torch
3. First run downloads the Whisper model (~1.5GB, one time)
4. Access from Workspace > Scripts > caption
```

### Standalone
```
1. Clone repo
2. pip install -r requirements.txt
3. python caption.py --file video.mp4
```

## MVP Scope (v1)

- [x] Extract audio from current Resolve timeline (render approach)
- [x] Transcribe with faster-whisper, word-level timestamps
- [x] Generate SRT file (traditional subtitle style)
- [x] Auto-detect language (or force via config)
- [x] Filler word removal + hallucination filtering (from Bark)
- [x] Confidence flagging for low-probability words
- [x] Cross-platform (Windows CUDA + Mac Metal)
- [x] CLI mode for direct file input
- [ ] Manual SRT import into Resolve (user does File > Import > Subtitle)

## v2 Scope

- [ ] Short-form caption style (social/TikTok)
- [ ] Word-by-word karaoke style
- [ ] Simple tkinter GUI with preview
- [ ] Batch processing (multiple timelines/files)
- [ ] Direct file audio extraction via ffmpeg (skip Resolve render for speed)
- [ ] Speaker diarization (who said what) via pyannote or similar
- [ ] Auto-import subtitles via Resolve API if available
- [ ] Timeline markers at low-confidence spots for quick review

## v3 Scope

- [ ] Translation (transcribe in Swedish, generate English subtitles)
- [ ] Custom vocabulary / proper noun correction
- [ ] Integration with Bark (live dictation + post-production captions from same engine)
- [ ] Resolve Fusion integration for styled subtitles

## Open Questions

1. **Naming**: "Resolve Captions"? "Bark Captions"? Keep it in the Bark repo or separate?
2. **Resolve version**: What version of Resolve are we targeting? (API differences between 18/19/Studio vs Free)
3. **SRT vs other formats**: Start with SRT only, or also support VTT/TTML?
4. **Audio extraction**: Render approach (reliable) vs ffmpeg direct (fast) - start with render?
5. **Resolve Free vs Studio**: The scripting API is Studio-only. Do we need CLI-only mode for Free users?
