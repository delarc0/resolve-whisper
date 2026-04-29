"""
Resolve Whisper - AI-powered caption generation for DaVinci Resolve Studio.

Usage:
    From Resolve Scripts menu:   Workspace > Scripts > caption
    From command line:           python caption.py --file video.mp4
                                 python caption.py --language sv
                                 python caption.py --file audio.wav --output captions.srt
"""
import argparse
import logging
import os
import sys
import tempfile
import time

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def get_resolve():
    """Connect to a running DaVinci Resolve Studio instance."""
    try:
        import DaVinciResolveScript as bmd
        resolve = bmd.scriptapp("Resolve")
        if resolve is None:
            log.error("Could not connect to Resolve. Is it running?")
            return None
        return resolve
    except ImportError:
        log.error(
            "DaVinciResolveScript module not found. "
            "Make sure Resolve's scripting modules are on PYTHONPATH.\n"
            "  Windows: set PYTHONPATH=%PROGRAMDATA%\\Blackmagic Design\\"
            "DaVinci Resolve\\Support\\Developer\\Scripting\\Modules\n"
            "  Mac: export PYTHONPATH=/Library/Application Support/"
            "Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules"
        )
        return None


def get_timeline_info(resolve):
    """Get current project, timeline, and frame rate."""
    pm = resolve.GetProjectManager()
    project = pm.GetCurrentProject()
    if not project:
        log.error("No project open in Resolve.")
        return None, None, None

    timeline = project.GetCurrentTimeline()
    if not timeline:
        log.error("No timeline selected in Resolve.")
        return project, None, None

    fps_str = timeline.GetSetting("timelineFrameRate")
    try:
        fps = float(fps_str)
    except (ValueError, TypeError):
        fps = 24.0
        log.warning(f"Could not read frame rate (got '{fps_str}'), defaulting to {fps}")

    name = timeline.GetName()
    duration = timeline.GetEndFrame() - timeline.GetStartFrame()
    dur_s = duration / fps
    log.info(f"Timeline: '{name}' | {dur_s:.1f}s | {fps}fps")

    return project, timeline, fps


def render_audio(project, timeline, output_dir: str) -> str:
    """Render timeline audio via Resolve's Quick Export API.

    Quick Export does NOT modify the project's persistent render settings
    (TargetDir, CustomName, format, codec). Compared to AddRenderJob, it
    fires a one-shot render with our params and leaves the user's normal
    Deliver-page settings completely untouched.
    """
    timeline_name = timeline.GetName()
    safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
    if not safe_name.strip():
        safe_name = "timeline_audio"
    wav_name = f"{safe_name}_audio"

    log.info("Rendering timeline audio via Quick Export...")

    _AUDIO_ONLY_NAMES = ("Audio Only", "audio only", "Audio-Only", "audio-only")
    try:
        presets = project.GetQuickExportRenderPresets() or []
        log.info(f"Quick Export presets: {presets}")
    except Exception as e:
        log.warning(f"GetQuickExportRenderPresets failed: {e}")
        presets = []

    chosen = next((n for n in _AUDIO_ONLY_NAMES if n in presets), None)
    if not chosen:
        log.error("'Audio Only' preset not in Quick Export list.")
        return None

    try:
        result = project.RenderWithQuickExport(chosen, {
            "TargetDir": output_dir,
            "CustomName": wav_name,
            "EnableUpload": False,
        })
        log.info(f"RenderWithQuickExport returned: {result}")
    except Exception as e:
        log.error(f"RenderWithQuickExport failed: {e}")
        return None

    # Quick Export's blocking semantics aren't documented; poll for the
    # output file with a generous timeout in case it returns before the
    # write is complete.
    _AUDIO_EXTS = (".wav", ".mov", ".mp4", ".flac", ".mp3", ".aac", ".m4a", ".mxf")
    _timeout = 600  # 10 minutes
    _start = time.time()
    audio_path = None
    while time.time() - _start < _timeout:
        try:
            files = os.listdir(output_dir)
        except FileNotFoundError:
            files = []
        for ext in _AUDIO_EXTS:
            audio_path = next(
                (os.path.join(output_dir, f) for f in files if f.lower().endswith(ext)),
                None,
            )
            if audio_path:
                break
        if audio_path:
            break
        time.sleep(0.5)

    if not audio_path:
        log.error("Render completed but audio file not found within timeout.")
        return None

    try:
        size_mb = os.path.getsize(audio_path) / (1024 * 1024)
        log.info(f"Audio rendered: {audio_path} ({size_mb:.1f} MB)")
    except Exception:
        log.info(f"Audio rendered: {audio_path}")
    return audio_path


def run_resolve_mode(args):
    """Full pipeline: Resolve timeline -> audio -> transcribe -> SRT."""
    import shutil
    from config import cfg
    from transcribe import Transcriber
    from srt import write_srt

    # Apply CLI overrides
    if args.language:
        cfg["language"] = args.language
    if args.max_words is not None:
        cfg["max_words_per_caption"] = args.max_words
    if args.max_chars is not None:
        cfg["max_chars_per_line"] = args.max_chars
    if args.max_lines is not None:
        cfg["max_lines"] = args.max_lines

    resolve = get_resolve()
    if not resolve:
        return 1

    project, timeline, fps = get_timeline_info(resolve)
    if not timeline:
        return 1

    # Render audio to temp dir
    tmp_dir = tempfile.mkdtemp(prefix="resolve_whisper_")
    wav_path = render_audio(project, timeline, tmp_dir)
    if not wav_path:
        return 1

    # Transcribe
    log.info("Loading Whisper model...")
    t0 = time.time()
    transcriber = Transcriber()
    load_time = time.time() - t0
    log.info(f"Model loaded in {load_time:.1f}s")

    log.info("Transcribing...")
    t0 = time.time()
    segments = transcriber.transcribe(wav_path)
    tx_time = time.time() - t0

    word_count = sum(len(s.words) for s in segments)
    log.info(f"Transcribed {word_count} words in {tx_time:.1f}s")

    if not segments:
        log.warning("No speech detected in timeline audio.")
        return 1

    strip_punct = getattr(args, "strip_punctuation", False)

    # Generate SRT
    output_dir = args.output_dir or cfg["output_dir"]
    if not output_dir:
        # Default: same directory as Resolve project database? Not accessible.
        # Use a "Captions" folder on Desktop instead.
        output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Captions")

    os.makedirs(output_dir, exist_ok=True)

    timeline_name = timeline.GetName()
    safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
    srt_path = os.path.join(output_dir, f"{safe_name}.srt")

    success = write_srt(segments, srt_path, fps, strip_punctuation=strip_punct)
    if not success:
        return 1

    # Clean up temp render
    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass

    log.info("")
    log.info(f"SRT saved to: {srt_path}")
    log.info("Drag from Finder onto a subtitle track in your timeline.")
    _open_folder(output_dir)
    return 0


def run_file_mode(args):
    """Transcribe a file directly without Resolve."""
    from config import cfg

    # Apply device override BEFORE importing Transcriber
    # (it reads _config.DEVICE when creating the model)
    if getattr(args, "device", None):
        import config as _cfg_mod
        _cfg_mod.DEVICE = args.device
        _cfg_mod.COMPUTE_TYPE = "int8" if args.device == "cpu" else "float16"

    from transcribe import Transcriber
    from srt import write_srt, write_captions_json

    if args.language:
        cfg["language"] = args.language
    if args.max_words is not None:
        cfg["max_words_per_caption"] = args.max_words
    if args.max_chars is not None:
        cfg["max_chars_per_line"] = args.max_chars
    if args.max_lines is not None:
        cfg["max_lines"] = args.max_lines

    input_path = os.path.abspath(args.file)
    if not os.path.exists(input_path):
        log.error(f"File not found: {input_path}")
        return 1

    log.info(f"Input: {input_path}")

    # Determine output path
    if args.output:
        srt_path = os.path.abspath(args.output)
    else:
        base = os.path.splitext(input_path)[0]
        srt_path = f"{base}.srt"

    # Transcribe
    log.info("Loading Whisper model...")
    t0 = time.time()
    transcriber = Transcriber()
    load_time = time.time() - t0
    log.info(f"Model loaded in {load_time:.1f}s")

    def _report_progress(pct):
        # Machine-readable line for resolve_script.py to parse
        print(f"PROGRESS:{pct}", flush=True)

    log.info("Transcribing...")
    t0 = time.time()
    segments = transcriber.transcribe(input_path, on_progress=_report_progress)
    tx_time = time.time() - t0

    word_count = sum(len(s.words) for s in segments)
    log.info(f"Transcribed {word_count} words in {tx_time:.1f}s")

    if not segments:
        log.warning("No speech detected.")
        return 1

    strip_punct = getattr(args, "strip_punctuation", False)

    # Determine FPS (default 24 for standalone files)
    fps = args.fps or 24.0

    success = write_srt(segments, srt_path, fps, strip_punctuation=strip_punct)
    if not success:
        return 1

    # Write JSON sidecar for Text+ insertion mode
    json_path = os.path.splitext(srt_path)[0] + ".json"
    write_captions_json(segments, json_path, fps, strip_punctuation=strip_punct)

    log.info("")
    log.info(f"SRT saved to: {srt_path}")
    return 0


def _open_folder(path: str):
    """Open a folder in the system file manager."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser(
        description="Resolve Whisper - AI caption generation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python caption.py                       Transcribe current Resolve timeline
  python caption.py --file video.mp4      Transcribe a video file directly
  python caption.py --language sv         Force Swedish language
  python caption.py --file a.wav -o a.srt Specify output file
        """,
    )
    parser.add_argument(
        "--file", "-f",
        help="Audio/video file to transcribe (skips Resolve, direct file mode)",
    )
    parser.add_argument(
        "--output", "-o",
        help="Output SRT file path (default: same name as input with .srt)",
    )
    parser.add_argument(
        "--output-dir",
        help="Output directory for SRT (Resolve mode)",
    )
    parser.add_argument(
        "--language", "-l",
        help="Language code (e.g., 'sv', 'en'). Default: auto-detect",
    )
    parser.add_argument(
        "--fps",
        type=float,
        help="Frame rate for subtitle gap calculation (default: 24.0 or from Resolve)",
    )
    parser.add_argument(
        "--max-words",
        type=int,
        help="Max words per caption (0=auto). Overrides config.",
    )
    parser.add_argument(
        "--max-chars",
        type=int,
        help="Max characters per line. Overrides config.",
    )
    parser.add_argument(
        "--max-lines",
        type=int,
        help="Max lines per caption. Overrides config.",
    )
    parser.add_argument(
        "--strip-punctuation",
        action="store_true",
        help="Remove all punctuation from captions.",
    )
    parser.add_argument(
        "--device",
        choices=["cuda", "cpu"],
        help="Force device (cuda or cpu). Default: auto-detect.",
    )

    args = parser.parse_args()

    if args.file:
        return run_file_mode(args)
    else:
        return run_resolve_mode(args)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log.info("Cancelled.")
        sys.exit(1)
    except Exception:
        log.exception("Fatal error")
        sys.exit(1)
