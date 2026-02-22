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
    """Render timeline audio as WAV file via Resolve's render API."""
    timeline_name = timeline.GetName()
    safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
    wav_name = f"{safe_name}_audio"
    wav_path = os.path.join(output_dir, f"{wav_name}.wav")

    log.info("Rendering timeline audio...")

    project.SetRenderSettings({
        "ExportAudio": True,
        "ExportVideo": False,
        "TargetDir": output_dir,
        "CustomName": wav_name,
        "AudioCodec": "LinearPCM",
        "AudioBitDepth": 16,
        "AudioSampleRate": 16000,
        "FormatWidth": 0,
        "FormatHeight": 0,
    })

    job_id = project.AddRenderJob()
    if not job_id:
        log.error("Failed to add render job. Check Resolve render settings.")
        return None

    project.StartRendering()

    # Poll until done
    while project.IsRenderingInProgress():
        time.sleep(0.5)

    status = project.GetRenderJobStatus(job_id)
    if status and status.get("JobStatus") == "Complete":
        # Find the rendered file
        target = status.get("TargetDir", output_dir)
        rendered = os.path.join(target, f"{wav_name}.wav")
        if os.path.exists(rendered):
            size_mb = os.path.getsize(rendered) / (1024 * 1024)
            log.info(f"Audio rendered: {rendered} ({size_mb:.1f} MB)")
            return rendered

    # Fallback: search output dir for the file
    for f in os.listdir(output_dir):
        if f.endswith(".wav") and safe_name in f:
            found = os.path.join(output_dir, f)
            log.info(f"Audio rendered: {found}")
            return found

    log.error("Render completed but WAV file not found.")
    return None


def run_resolve_mode(args):
    """Full pipeline: Resolve timeline -> audio -> transcribe -> SRT."""
    from config import cfg
    from transcribe import Transcriber
    from srt import write_srt

    # Apply CLI overrides
    if args.language:
        cfg["language"] = args.language

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

    success = write_srt(segments, srt_path, fps)
    if not success:
        return 1

    # Clean up temp audio
    try:
        os.unlink(wav_path)
        os.rmdir(tmp_dir)
    except Exception:
        pass

    log.info("")
    log.info(f"SRT saved to: {srt_path}")
    log.info("To import: File > Import > Subtitle in Resolve")

    # Open output folder
    _open_folder(output_dir)

    return 0


def run_file_mode(args):
    """Transcribe a file directly without Resolve."""
    from config import cfg
    from transcribe import Transcriber
    from srt import write_srt

    if args.language:
        cfg["language"] = args.language

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

    log.info("Transcribing...")
    t0 = time.time()
    segments = transcriber.transcribe(input_path)
    tx_time = time.time() - t0

    word_count = sum(len(s.words) for s in segments)
    log.info(f"Transcribed {word_count} words in {tx_time:.1f}s")

    if not segments:
        log.warning("No speech detected.")
        return 1

    # Determine FPS (default 24 for standalone files)
    fps = args.fps or 24.0

    success = write_srt(segments, srt_path, fps)
    if not success:
        return 1

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
