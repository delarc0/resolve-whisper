"""
Resolve Whisper - AI-powered caption generation for DaVinci Resolve Studio.

Usage:
    From Resolve Scripts menu:   Workspace > Scripts > caption
    From command line:           python caption.py --file video.mp4
                                 python caption.py --language sv
                                 python caption.py --file audio.wav --output captions.srt
"""
import argparse
import json
import logging
import os
import subprocess
import sys
import tempfile
import time

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


STATUS_FILE = os.path.join(tempfile.gettempdir(), "resolve_whisper_status.json")
LOCK_FILE = os.path.join(tempfile.gettempdir(), "resolve_whisper.lock")


def _acquire_run_lock() -> bool:
    """Claim the single-run lock or report another run is in progress.

    Atomic: O_CREAT|O_EXCL closes the check-then-write race between two
    near-simultaneous launches. A stale lock (recorded PID dead) is removed
    and acquisition retried once. PermissionError from the liveness probe
    means the PID exists under another user -> treat as alive.
    """
    for _attempt in (1, 2):
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(str(os.getpid()))
            return True
        except FileExistsError:
            if sys.platform == "win32":
                # os.kill(pid, 0) TERMINATES the target on Windows, so probe
                # by lock age instead: runs never legitimately exceed 2h.
                try:
                    age = time.time() - os.path.getmtime(LOCK_FILE)
                except OSError:
                    continue  # lock vanished between open and stat; retry
                if age < 7200:
                    log.error("Another caption run appears to be in progress. Wait for it to finish.")
                    return False
                try:
                    os.unlink(LOCK_FILE)
                except OSError:
                    pass
                continue
            try:
                with open(LOCK_FILE, "r", encoding="utf-8") as f:
                    other_pid = int(f.read().strip())
                os.kill(other_pid, 0)  # signal 0 == probe; raises if not alive
                log.error(f"Another caption run is in progress (pid {other_pid}). Wait for it to finish or cancel it.")
                return False
            except PermissionError:
                # Process exists under another user: definitely alive.
                log.error("Another caption run is in progress (different user). Wait for it to finish.")
                return False
            except (ValueError, ProcessLookupError, OSError):
                # Stale or garbage lock: remove it and retry the atomic open.
                try:
                    os.unlink(LOCK_FILE)
                except OSError:
                    pass
                continue
        except OSError as e:
            log.warning(f"Could not create lock file ({e}); proceeding without lock.")
            return True
    return False


def _release_run_lock():
    """Release the run lock if we own it. Best-effort; never raises."""
    try:
        with open(LOCK_FILE, "r", encoding="utf-8") as f:
            owner = int(f.read().strip())
        if owner == os.getpid():
            os.unlink(LOCK_FILE)
    except (OSError, ValueError):
        pass


# Our SRTs end in " YYYYMMDD-HHMMSS.srt" (see run_resolve_mode). The cleanup
# must only ever touch files matching this stamp -- output_dir is
# user-configurable and can contain hand-edited or delivered SRTs.
_OWN_SRT_RE = None  # compiled lazily; module-level re import kept minimal


def _is_own_srt(name: str) -> bool:
    global _OWN_SRT_RE
    if _OWN_SRT_RE is None:
        import re
        _OWN_SRT_RE = re.compile(r" \d{8}-\d{6}\.srt$", re.IGNORECASE)
    return bool(_OWN_SRT_RE.search(name))


def _cleanup_old_srts(output_dir: str, keep_days: int):
    """Delete OUR old .srt files in output_dir. Non-fatal.

    Only files matching the tool's own " YYYYMMDD-HHMMSS.srt" stamp are
    candidates; foreign SRTs in a shared folder are never touched.
    """
    if keep_days <= 0 or not os.path.isdir(output_dir):
        return
    cutoff = time.time() - keep_days * 86400
    removed = 0
    try:
        for name in os.listdir(output_dir):
            if not _is_own_srt(name):
                continue
            path = os.path.join(output_dir, name)
            try:
                if os.path.getmtime(path) < cutoff:
                    os.unlink(path)
                    removed += 1
            except OSError:
                pass
    except OSError as e:
        log.debug(f"Could not list output dir for cleanup: {e}")
        return
    if removed:
        log.info(f"Cleaned up {removed} SRT file(s) older than {keep_days} days.")


def _start_progress_estimator(audio_duration_s: float):
    """Run a background thread that fakes a progress percentage based on
    elapsed time vs an estimated transcription duration.

    mlx_whisper.transcribe() blocks without progress callbacks on Mac, so
    the UI would otherwise sit on 0% for minutes on long files. The
    estimate is rough -- M5 Pro hits ~5x real-time on large-v3 -- but it's
    enough to show the bar moving. Caps at 95% so the real completion
    callback gets to write 100%.

    Returns (stop_event, thread) so the caller can stop and join it.
    """
    import threading

    stop = threading.Event()
    # Conservative 4x real-time so the bar advances at most as fast as the
    # actual transcription -- better to under-promise than to sit on 95%
    # while transcription drags on.
    expected = max(audio_duration_s / 4.0, 1.0)

    def _run():
        t0 = time.time()
        while not stop.is_set():
            elapsed = time.time() - t0
            pct = int(min(elapsed / expected * 95.0, 95.0))
            _write_status("transcribing", f"~{pct}% complete", progress=pct)
            if stop.wait(0.5):
                return

    th = threading.Thread(target=_run, daemon=True)
    th.start()
    return stop, th


def _start_heartbeat(stage: str, message: str):
    """Re-write the same status every 5s so the UI's stale-file detector
    (90s) doesn't close the window during long blocking calls -- the
    first-run model download inside Transcriber() can take many minutes.

    Returns (stop_event, thread); call stop.set() + thread.join() when done.
    """
    import threading

    stop = threading.Event()

    def _run():
        while not stop.wait(5.0):
            _write_status(stage, message)

    th = threading.Thread(target=_run, daemon=True)
    th.start()
    return stop, th


def _write_status(stage: str, message: str = "", progress: int = -1):
    """Write current pipeline state for the progress UI to read.

    Best-effort: never raises. The UI is a separate process polling this file.
    PID is included so the UI's Cancel button can deliver SIGTERM to us.
    Atomic (tmp + os.replace): the estimator thread and the main thread both
    write here, and the UI's Cancel handler reads it -- a half-written file
    must never be observable, or Cancel silently no-ops.
    """
    import threading
    tmp_path = f"{STATUS_FILE}.{os.getpid()}.{threading.get_ident()}.tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump({
                "stage": stage,
                "progress": progress,
                "message": message,
                "ts": time.time(),
                "pid": os.getpid(),
            }, f)
        os.replace(tmp_path, STATUS_FILE)
    except Exception as e:
        log.debug(f"status write failed: {e}")
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def _install_cancel_signal():
    """Convert SIGTERM into KeyboardInterrupt.

    The progress-UI Cancel button sends SIGTERM to our PID. By raising
    KeyboardInterrupt, we let the existing try/finally cleanup paths run
    (delete render job, restore Deliver state, remove temp dir) instead of
    being killed dead.
    """
    if sys.platform == "win32":
        return  # Windows signal handling is too limited for this trick.
    try:
        import signal

        def _handler(signum, frame):
            raise KeyboardInterrupt(f"signal {signum}")

        signal.signal(signal.SIGTERM, _handler)
    except Exception as e:
        log.debug(f"signal handler install failed: {e}")


def _spawn_progress_ui():
    """Launch progress_ui.py as a subprocess. Returns the Popen or None.

    Failure is silent: if Tk isn't available or the spawn fails, the pipeline
    runs without UI feedback (logs only).
    """
    ui_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "progress_ui.py")
    if not os.path.exists(ui_script):
        return None
    try:
        # Reset status file so the new UI process doesn't pick up a stale "done"
        _write_status("starting", "Connecting to Resolve...")
        return subprocess.Popen(
            [sys.executable, ui_script, STATUS_FILE],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
    except Exception as e:
        log.debug(f"progress UI spawn failed: {e}")
        return None


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


_AUDIO_PRESET_NAME = "Audio Only"
_RENDER_TIMEOUT_S = 1800  # 30 min upper bound; long debates land well under this
_AUDIO_EXTS = (".wav", ".flac", ".mp3", ".aac", ".m4a", ".aif", ".aiff")


def _safe(call, *args, **kwargs):
    """Call a Resolve API method, catching exceptions and returning a default.

    Resolve's bridged objects throw opaque RuntimeErrors on bad calls in some
    versions; we'd rather log + continue than crash the pipeline.
    """
    default = kwargs.pop("_default", None)
    try:
        return call(*args, **kwargs)
    except Exception as e:
        log.debug(f"{getattr(call, '__name__', 'call')} failed: {e}")
        return default


def _start_rendering_compat(project, job_id) -> bool:
    """Call StartRendering across Resolve API drift.

    Doc/example variants seen in the wild:
      project.StartRendering(jobId)
      project.StartRendering(jobId, isInteractiveMode=False)
      project.StartRendering([jobId], False)
    Try the safest first, then fall back. Treat any truthy return as success.
    """
    for attempt in (
        lambda: project.StartRendering(job_id),
        lambda: project.StartRendering(job_id, False),
        lambda: project.StartRendering([job_id], False),
    ):
        try:
            result = attempt()
        except TypeError:
            continue
        except Exception as e:
            log.debug(f"StartRendering attempt failed: {e}")
            continue
        if result:
            return True
    return False


def _validate_audio_only_settings(settings: dict) -> list:
    """Return a list of human-readable problems with an audio-only render config.

    Empty list = preset is correctly configured. Pure helper, unit-testable.
    """
    problems = []
    if not isinstance(settings, dict):
        return ["render settings unreadable"]

    # Resolve sometimes uses string "0"/"1", sometimes bool. Normalise both.
    def _is_falsey(v):
        return v in (False, 0, "0", "false", "False", None, "")

    def _is_truthy(v):
        return v in (True, 1, "1", "true", "True")

    # Resolve 21 renamed the job-dict keys: ExportVideo -> IsExportVideo,
    # ExportAudio -> IsExportAudio. Check whichever spelling is present.
    video_key = next((k for k in ("ExportVideo", "IsExportVideo") if k in settings), None)
    audio_key = next((k for k in ("ExportAudio", "IsExportAudio") if k in settings), None)
    if video_key and not _is_falsey(settings[video_key]):
        problems.append(f"{video_key} is {settings[video_key]!r}, want false")
    if audio_key and not _is_truthy(settings[audio_key]):
        problems.append(f"{audio_key} is {settings[audio_key]!r}, want true")

    # Audio container: Resolve reports audio format under different keys
    # depending on version. We just look for evidence of WAV / PCM.
    audio_codec = settings.get("AudioCodec", "")
    if isinstance(audio_codec, str) and audio_codec and "pcm" not in audio_codec.lower() and "linearpcm" not in audio_codec.lower().replace(" ", ""):
        # Don't fail outright on codec mismatch - Whisper handles many codecs
        # via ffmpeg. Just warn through the problems list at info level.
        log.info(f"Audio codec is {audio_codec!r} (PCM preferred but not required).")

    return problems


def _expected_output_path(job_settings: dict, output_dir: str, fallback_name: str) -> str:
    """Compute the expected rendered file path from a job's settings."""
    target_dir = job_settings.get("TargetDir") or output_dir
    custom_name = job_settings.get("CustomName") or fallback_name
    # Resolve 21 exposes the exact output filename on the job; trust it first.
    out_name = job_settings.get("OutputFilename")
    if out_name:
        candidate = os.path.join(target_dir, out_name)
        if os.path.exists(candidate):
            return candidate
    # Resolve appends the extension based on format. We probe extensions
    # rather than try to predict from the format string.
    for ext in _AUDIO_EXTS:
        candidate = os.path.join(target_dir, custom_name + ext)
        if os.path.exists(candidate):
            return candidate
    # Fallback to scanning the dir for files matching CustomName
    try:
        files = os.listdir(target_dir)
    except FileNotFoundError:
        return ""
    for ext in _AUDIO_EXTS:
        match = next(
            (os.path.join(target_dir, f) for f in files
             if f.lower().endswith(ext) and custom_name.lower() in f.lower()),
            "",
        )
        if match:
            return match
    return ""


def _delete_job_if_ours(project, job_id: str, pre_existing_ids: set):
    """Only delete the job we added; never touch jobs the user had queued."""
    if not job_id:
        return
    if job_id in pre_existing_ids:
        log.warning(f"Refusing to delete job {job_id}: was already in queue.")
        return
    _safe(project.DeleteRenderJob, job_id)


# Format/codec ids for a WAV audio-only setup, in preference order.
# Resolve 21 dropped the 'LinearPCM' codec id (GetRenderCodecs('wav') returns
# {} and job dicts report 'lpcm'); Resolve 18-20 want 'LinearPCM'.
_WAV_FORMAT_CANDIDATES = (
    ("wav", "LinearPCM"),   # Resolve 18-20
    ("wav", "lpcm"),        # Resolve 21 job-dict codec id
    ("wav", ""),            # Resolve 21: audio formats expose no codec list
    ("Wave", "LinearPCM"),  # display-name spelling, defensive
    ("Wave", "lpcm"),
    ("Wave", ""),
)


def _set_wav_format_compat(project):
    """Switch the Deliver page to WAV across Resolve versions.

    Tries each known format/codec id pair until one sticks. Returns the
    (format, codec) pair that worked, or None if every candidate failed.
    """
    for fmt, codec in _WAV_FORMAT_CANDIDATES:
        if _safe(project.SetCurrentRenderFormatAndCodec, fmt, codec):
            return (fmt, codec)
    return None


def _restore_deliver_state(project, saved_fmt: dict, saved_mode):
    """Put the Deliver page back to roughly what the user had before our run.

    We can restore format/codec/mode (the API exposes getters) but not the
    user's preset NAME (no API for that), so the dropdown will show 'Custom'
    rather than the original preset name. The format/codec being right means
    the user's next video export doesn't start on Audio Only. We also clear
    our injected TargetDir/CustomName so the user's path field isn't filled
    with our temp location.
    """
    if isinstance(saved_fmt, dict):
        fmt = saved_fmt.get("format")
        codec = saved_fmt.get("codec")
        if fmt and codec:
            _safe(project.SetCurrentRenderFormatAndCodec, fmt, codec)
    if saved_mode is not None:
        _safe(project.SetCurrentRenderMode, saved_mode)
    # Clear our temp-dir leftovers so the Deliver page doesn't show
    # /var/folders/... in the path field.
    _safe(project.SetRenderSettings, {"TargetDir": "", "CustomName": ""})


def render_audio(project, timeline, output_dir: str) -> str:
    """Render timeline audio via the regular render queue.

    Quick Export's preset universe is built-in only and can't see the user's
    'Audio Only' preset, so we use LoadRenderPreset + AddRenderJob. To avoid
    leaving the Deliver page in a polluted state (temp /var/folders paths
    sticking around), we re-load the preset at the end so the page resets
    to a clean known state with no TargetDir override.

    Hardened:
      - Polls GetRenderJobStatus(jobId) so other queued jobs don't confuse us
      - Tracks pre-existing job ids; we only ever delete our own
      - Validates the loaded preset before adding the job
      - Resolves output path from job-settings, not a directory scan
      - Snapshots format/codec/mode so the Deliver page can be returned to
        a video-export-ready state instead of being stuck on Audio Only
    """
    timeline_name = timeline.GetName()
    safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
    if not safe_name.strip():
        safe_name = "timeline_audio"
    wav_name = f"{safe_name}_audio"

    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError as e:
            log.error(f"Cannot create output dir {output_dir}: {e}")
            return None
    if not os.access(output_dir, os.W_OK):
        log.error(f"Output dir not writable: {output_dir}")
        return None

    presets = _safe(project.GetRenderPresetList, _default=[]) or []
    if _AUDIO_PRESET_NAME not in presets:
        log.error(f"'{_AUDIO_PRESET_NAME}' render preset not found in this project.")
        log.error("Run: python caption.py --check  (will create it)")
        return None

    # Snapshot the user's Deliver page state BEFORE we touch anything, so we
    # can put them back on a video-export-ready state at the end. We can't
    # snapshot the preset NAME (no API for that), but format/codec/mode is
    # enough that the user doesn't end up stuck on the Audio Only preset
    # next time they go to Deliver to export video.
    saved_fmt = _safe(project.GetCurrentRenderFormatAndCodec, _default={}) or {}
    saved_mode = _safe(project.GetCurrentRenderMode)

    # Snapshot pre-existing job ids so we never delete jobs the user queued.
    existing_jobs = _safe(project.GetRenderJobList, _default=[]) or []
    pre_existing_ids = {j.get("JobId") for j in existing_jobs if isinstance(j, dict) and j.get("JobId")}

    log.info(f"Loading '{_AUDIO_PRESET_NAME}' preset...")
    if not _safe(project.LoadRenderPreset, _AUDIO_PRESET_NAME):
        log.error(f"LoadRenderPreset('{_AUDIO_PRESET_NAME}') failed.")
        return None

    if not _safe(project.SetRenderSettings, {
        "TargetDir": output_dir,
        "CustomName": wav_name,
    }):
        log.warning("SetRenderSettings returned falsy; continuing.")

    job_id = _safe(project.AddRenderJob)
    if not job_id:
        log.error("AddRenderJob returned no id.")
        _restore_deliver_state(project, saved_fmt, saved_mode)
        return None

    # Validate the queued job actually exports audio (catches a tampered preset).
    job_settings = {}
    all_jobs = _safe(project.GetRenderJobList, _default=[]) or []
    for j in all_jobs:
        if isinstance(j, dict) and j.get("JobId") == job_id:
            job_settings = j
            break
    problems = _validate_audio_only_settings(job_settings)
    if problems:
        log.error(
            f"'{_AUDIO_PRESET_NAME}' preset is misconfigured: " + "; ".join(problems)
        )
        log.error("Recreate it: python caption.py --check")
        _delete_job_if_ours(project, job_id, pre_existing_ids)
        _restore_deliver_state(project, saved_fmt, saved_mode)
        return None

    log.info("Rendering timeline audio...")
    audio_path = None
    try:
        if not _start_rendering_compat(project, job_id):
            log.error("StartRendering failed across all known signatures.")
            return None

        start = time.time()
        last_pct = -1
        empty_polls = 0
        while time.time() - start < _RENDER_TIMEOUT_S:
            status = _safe(project.GetRenderJobStatus, job_id, _default={}) or {}

            # Resolve quitting/crashing mid-render makes every API call
            # return empty through _safe. Detect it instead of spinning
            # silently for the full 30-minute timeout.
            if not status:
                empty_polls += 1
                if empty_polls >= 20:  # ~10s of nothing back
                    if not _safe(project.GetName):
                        log.error("Lost connection to Resolve (did it quit?). Aborting.")
                        _write_status("error", "Lost connection to Resolve.")
                        return None
                    empty_polls = 0  # project still alive; transient blips
            else:
                empty_polls = 0

            job_status = status.get("JobStatus", "")
            pct = status.get("CompletionPercentage", 0)
            try:
                pct_int = int(pct)
            except (ValueError, TypeError):
                pct_int = 0
            if pct_int != last_pct and pct_int % 10 == 0:
                log.info(f"  render {pct_int}%")
                last_pct = pct_int
            # Heartbeat every poll: long renders otherwise trip the UI's
            # 90s stale-file auto-close, taking the Cancel button with it.
            _write_status("rendering_audio", f"Rendering... {pct_int}%", progress=pct_int)

            if job_status == "Complete":
                break
            if job_status in ("Failed", "Cancelled"):
                log.error(f"Render ended with status: {job_status}")
                return None
            time.sleep(0.5)
        else:
            log.error(f"Render timed out after {_RENDER_TIMEOUT_S}s.")
            _safe(project.StopRendering)
            return None

        # Resolve output path from the job's actual settings.
        job_settings_post = {}
        for j in (_safe(project.GetRenderJobList, _default=[]) or []):
            if isinstance(j, dict) and j.get("JobId") == job_id:
                job_settings_post = j
                break
        audio_path = _expected_output_path(job_settings_post, output_dir, wav_name)
        if not audio_path or not os.path.exists(audio_path):
            log.error("Render reported complete but audio file not found.")
            return None

        try:
            size_mb = os.path.getsize(audio_path) / (1024 * 1024)
            log.info(f"Audio rendered: {audio_path} ({size_mb:.1f} MB)")
        except Exception:
            log.info(f"Audio rendered: {audio_path}")
        return audio_path

    except KeyboardInterrupt:
        log.warning("Interrupted; stopping render and cleaning up...")
        _safe(project.StopRendering)
        raise
    finally:
        _delete_job_if_ours(project, job_id, pre_existing_ids)
        _restore_deliver_state(project, saved_fmt, saved_mode)


def run_resolve_mode(args):
    """Full pipeline: Resolve timeline -> audio -> transcribe -> SRT."""
    import shutil
    from config import cfg
    from transcribe import Transcriber
    from srt import write_srt

    # Apply CLI overrides. "--language auto" means explicit auto-detect
    # (cfg language defaults to "sv", so the Auto preset must be able to
    # override it back to None).
    if args.language:
        cfg["language"] = None if args.language.lower() in ("auto", "none") else args.language
    if args.max_words is not None:
        cfg["max_words_per_caption"] = args.max_words
    if args.max_chars is not None:
        cfg["max_chars_per_line"] = args.max_chars
    if args.max_lines is not None:
        cfg["max_lines"] = args.max_lines

    ui_proc = None if getattr(args, "no_ui", False) else _spawn_progress_ui()
    _install_cancel_signal()
    tmp_dir = None

    try:
        _write_status("starting", "Connecting to Resolve...")
        resolve = get_resolve()
        if not resolve:
            _write_status("error", "Resolve isn't running. Open Resolve Studio and try again.")
            return 1

        project, timeline, fps = get_timeline_info(resolve)
        if not project:
            _write_status("error", "No project open in Resolve. Open a project first.")
            return 1
        if not timeline:
            _write_status("error", "No timeline selected. Open a timeline and set in/out points.")
            return 1

        _write_status("rendering_audio", f"Timeline: {timeline.GetName()}")
        tmp_dir = tempfile.mkdtemp(prefix="resolve_whisper_")
        wav_path = render_audio(project, timeline, tmp_dir)
        if not wav_path:
            _write_status(
                "error",
                "Audio render failed. Check in/out points and the 'Audio Only' preset (run LAB37 Check).",
            )
            return 1

        _write_status("loading_model", "Warming up Whisper...")
        log.info("Loading Whisper model...")
        t0 = time.time()
        # Heartbeat: the first-ever run downloads ~3GB inside Transcriber(),
        # far beyond the UI's 90s stale-file auto-close.
        hb_stop, hb_th = _start_heartbeat("loading_model", "Loading AI model (first run downloads ~3 GB)...")
        try:
            transcriber = Transcriber()
        except Exception as e:
            msg = str(e).lower()
            if "huggingface" in msg or "connection" in msg or "no such file" in msg:
                _write_status("error", "Model download failed. Check internet connection and retry.")
            else:
                _write_status("error", f"Whisper model failed to load: {e}")
            return 1
        finally:
            hb_stop.set()
            hb_th.join(timeout=1.0)
        load_time = time.time() - t0
        log.info(f"Model loaded in {load_time:.1f}s")

        _write_status("transcribing", "Listening to audio...", progress=0)
        log.info("Transcribing...")

        # Background progress estimator: mlx_whisper.transcribe blocks without
        # callbacks on Mac, so we fake-advance the bar based on elapsed time.
        # On Windows faster_whisper still calls _on_progress for real updates.
        audio_dur = transcriber.get_audio_duration(wav_path)
        prog_stop, prog_th = _start_progress_estimator(audio_dur)

        def _on_progress(pct):
            # Real progress from faster_whisper -- stop the estimator and use this.
            if not prog_stop.is_set():
                prog_stop.set()
            try:
                pct_int = max(0, min(int(pct), 100))
            except (ValueError, TypeError):
                return
            _write_status("transcribing", f"{pct_int}% complete", progress=pct_int)

        t0 = time.time()
        try:
            segments = transcriber.transcribe(wav_path, on_progress=_on_progress)
        finally:
            prog_stop.set()
            prog_th.join(timeout=1.0)
        tx_time = time.time() - t0

        word_count = sum(len(s.words) for s in segments)
        log.info(f"Transcribed {word_count} words in {tx_time:.1f}s")

        if not segments:
            log.warning("No speech detected in timeline audio.")
            _write_status(
                "error",
                "No speech detected. Check in/out points cover the spoken section, and audio is unmuted.",
            )
            return 1

        strip_punct = getattr(args, "strip_punctuation", False)

        output_dir = args.output_dir or cfg["output_dir"]
        if not output_dir:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Captions")
        os.makedirs(output_dir, exist_ok=True)
        _cleanup_old_srts(output_dir, int(cfg.get("keep_srt_days", 30) or 0))

        timeline_name = timeline.GetName()
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
        # Per-run timestamp suffix so Resolve treats each import as a fresh
        # media item -- importing an SRT with a name Resolve has seen before
        # will reuse the cached pool item and ignore the file changes.
        run_stamp = time.strftime("%Y%m%d-%H%M%S")
        srt_path = os.path.join(output_dir, f"{safe_name} {run_stamp}.srt")

        _write_status("writing_srt", "Writing captions to disk...")
        success = write_srt(segments, srt_path, fps, strip_punctuation=strip_punct)
        if not success:
            _write_status("error", "SRT write failed.")
            return 1

        # Switch back to Edit page so the user lands where they expect
        # (Resolve briefly switches to Deliver during AddRenderJob in some versions).
        try:
            resolve.OpenPage("edit")
        except Exception as e:
            log.debug(f"OpenPage('edit') failed: {e}")

        # Import the SRT into the Media Pool, then place it on a subtitle
        # track via AppendToTimeline. Never via Finder drag -- Resolve
        # 21.0.2 can segfault in its drag handler (dragEnterEvent).
        #
        # AppendToTimeline quirks (learned the hard way):
        # - it returns None for Subtitle items EVEN ON SUCCESS, so the only
        #   reliable signal is counting items on the subtitle track after
        # - it fails (places nothing) when another collaborator holds the
        #   timeline lock -- that's the fallback path below
        # - it targets subtitle track 1, so we only auto-place onto a
        #   timeline with NO existing subtitle tracks; if tracks exist we
        #   must not risk stacking onto someone's hand-edited captions
        imported_to_pool = False
        auto_placed = False
        media_pool = _safe(project.GetMediaPool)
        pool_items = None
        if media_pool:
            pool_items = _safe(media_pool.ImportMedia, [srt_path])
            imported_to_pool = bool(pool_items)

        if imported_to_pool:
            existing_sub_tracks = _safe(timeline.GetTrackCount, "subtitle", _default=0) or 0
            if existing_sub_tracks == 0:
                _safe(timeline.AddTrack, "subtitle")
                _safe(media_pool.AppendToTimeline, list(pool_items))
                placed = _safe(timeline.GetItemListInTrack, "subtitle", 1, _default=[]) or []
                auto_placed = len(placed) > 0
                if auto_placed:
                    log.info(f"Captions placed on subtitle track ({len(placed)} items).")
            else:
                log.info("Timeline already has subtitle track(s); not auto-placing to avoid mixing with existing captions.")

        log.info("")
        log.info(f"SRT saved to: {srt_path}")
        if auto_placed:
            _write_status("done", "Captions are on your timeline.", progress=100)
        elif imported_to_pool:
            log.info("SRT imported into the Media Pool.")
            log.info("In Resolve: right-click it > Insert Selected Subtitles to Timeline.")
            log.info("(Avoid dragging SRTs from Finder -- Resolve 21.0.2 can crash on that.)")
            _write_status("done", "In Media Pool: right-click > Insert Selected Subtitles", progress=100)
        else:
            log.info("Drag import is unreliable on Resolve 21; use File > Import > Subtitle.")
            _open_folder(output_dir)
            _write_status("done", os.path.basename(srt_path), progress=100)
        return 0
    except KeyboardInterrupt:
        log.warning("Cancelled by user.")
        _write_status("error", "Cancelled.")
        return 1
    except Exception as e:
        _write_status("error", str(e))
        raise
    finally:
        # Always clean up the temp render dir, success or fail.
        if tmp_dir:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass
        # Don't kill the UI process; it auto-closes after seeing 'done' / 'error'.
        # If we exited via 'return 1' without writing a terminal status,
        # mark errored so the UI doesn't hang.
        if ui_proc is not None and ui_proc.poll() is None:
            try:
                with open(STATUS_FILE, "r", encoding="utf-8") as f:
                    cur = json.load(f)
                if cur.get("stage") not in ("done", "error"):
                    _write_status("error", "Pipeline ended unexpectedly.")
            except Exception:
                pass


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
        cfg["language"] = None if args.language.lower() in ("auto", "none") else args.language
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

    # Spawn UI by default; disable with --no-ui.
    ui_proc = None if getattr(args, "no_ui", False) else _spawn_progress_ui()
    _install_cancel_signal()

    # Same try/except/finally discipline as run_resolve_mode: a crash or a
    # Cancel-induced KeyboardInterrupt must write a terminal status, or the
    # progress window hangs until its 90s stale timeout.
    try:
        _write_status("loading_model", "Warming up Whisper...")
        log.info("Loading Whisper model...")
        t0 = time.time()
        hb_stop, hb_th = _start_heartbeat("loading_model", "Loading AI model (first run downloads ~3 GB)...")
        try:
            transcriber = Transcriber()
        except Exception as e:
            _write_status("error", f"Whisper model failed to load: {e}")
            return 1
        finally:
            hb_stop.set()
            hb_th.join(timeout=1.0)
        load_time = time.time() - t0
        log.info(f"Model loaded in {load_time:.1f}s")

        audio_dur = transcriber.get_audio_duration(input_path)
        prog_stop, prog_th = _start_progress_estimator(audio_dur)

        def _on_progress(pct):
            if not prog_stop.is_set():
                prog_stop.set()
            try:
                pct_int = max(0, min(int(pct), 100))
            except (ValueError, TypeError):
                return
            _write_status("transcribing", f"{pct_int}% complete", progress=pct_int)

        _write_status("transcribing", "Listening to audio...", progress=0)
        log.info("Transcribing...")
        t0 = time.time()
        try:
            segments = transcriber.transcribe(input_path, on_progress=_on_progress)
        finally:
            prog_stop.set()
            prog_th.join(timeout=1.0)
        tx_time = time.time() - t0

        word_count = sum(len(s.words) for s in segments)
        log.info(f"Transcribed {word_count} words in {tx_time:.1f}s")

        if not segments:
            log.warning("No speech detected.")
            _write_status("error", "No speech detected.")
            return 1

        strip_punct = getattr(args, "strip_punctuation", False)

        # Determine FPS (default 24 for standalone files)
        fps = args.fps or 24.0

        _write_status("writing_srt", "Writing captions to disk...")
        success = write_srt(segments, srt_path, fps, strip_punctuation=strip_punct)
        if not success:
            _write_status("error", "SRT write failed.")
            return 1

        # Write JSON sidecar for Text+ insertion mode
        json_path = os.path.splitext(srt_path)[0] + ".json"
        write_captions_json(segments, json_path, fps, strip_punctuation=strip_punct)

        log.info("")
        log.info(f"SRT saved to: {srt_path}")
        _write_status("done", os.path.basename(srt_path), progress=100)
        return 0
    except KeyboardInterrupt:
        log.warning("Cancelled by user.")
        _write_status("error", "Cancelled.")
        return 1
    except Exception as e:
        _write_status("error", str(e))
        raise
    finally:
        if ui_proc is not None and ui_proc.poll() is None:
            try:
                with open(STATUS_FILE, "r", encoding="utf-8") as f:
                    cur = json.load(f)
                if cur.get("stage") not in ("done", "error"):
                    _write_status("error", "Pipeline ended unexpectedly.")
            except Exception:
                pass


def run_check_mode(args):
    """Pre-flight check: validate environment + Resolve + preset.

    Run this once after install or after a Resolve update. Creates the
    'Audio Only' preset if missing. Reports each step PASS/FAIL with a
    clear remedy.
    """
    failures = []
    warnings = []

    def _row(name, ok, detail=""):
        prefix = "PASS" if ok else "FAIL"
        log.info(f"  [{prefix}] {name}{(' - ' + detail) if detail else ''}")
        if not ok:
            failures.append(name)

    log.info("=== Resolve Whisper pre-flight check ===")

    # 1. Python version
    pyver = sys.version_info
    _row(
        "Python 3.10+",
        pyver >= (3, 10),
        f"running {pyver.major}.{pyver.minor}.{pyver.micro}",
    )

    # 2. mlx_whisper importable (Mac) / faster_whisper (other)
    if sys.platform == "darwin":
        try:
            import mlx_whisper  # noqa: F401
            _row("mlx_whisper importable", True)
        except Exception as e:
            _row("mlx_whisper importable", False, str(e))
    else:
        try:
            from faster_whisper import WhisperModel  # noqa: F401
            _row("faster_whisper importable", True)
        except Exception as e:
            _row("faster_whisper importable", False, str(e))

    # 3. ffmpeg on PATH
    import shutil as _sh
    ffmpeg_path = _sh.which("ffmpeg")
    _row("ffmpeg on PATH", bool(ffmpeg_path), ffmpeg_path or "install: brew install ffmpeg")

    # 3a. silero-vad importable + ONNX model loads (Mac path needs this for
    # accurate word timing; Windows uses faster-whisper's built-in VAD).
    if sys.platform == "darwin":
        try:
            from silero_vad import load_silero_vad
            m = load_silero_vad(onnx=True)
            _row("silero-vad (ONNX) loadable", bool(m))
        except Exception as e:
            _row(
                "silero-vad (ONNX) loadable", False,
                f"{e} -- pip install silero-vad onnxruntime",
            )

    # 4. Resolve scripting module
    try:
        import DaVinciResolveScript as _bmd  # noqa: F401
        _row("DaVinciResolveScript importable", True)
    except Exception as e:
        _row("DaVinciResolveScript importable", False, str(e))

    # 5. Connect to Resolve (only if module imported)
    resolve = None
    if "DaVinciResolveScript importable" not in failures:
        resolve = get_resolve()
        _row("Connect to Resolve", bool(resolve), "is Resolve running?")

    project = None
    timeline = None
    if resolve:
        pm = resolve.GetProjectManager()
        project = pm.GetCurrentProject() if pm else None
        _row("Project open", bool(project))
        if project:
            timeline = project.GetCurrentTimeline()
            _row("Timeline open", bool(timeline))

    # 6. Output dir writable
    output_dir = (
        args.output_dir
        or os.path.join(os.path.expanduser("~"), "Desktop", "Captions")
    )
    try:
        os.makedirs(output_dir, exist_ok=True)
        ok = os.access(output_dir, os.W_OK)
        _row("Output dir writable", ok, output_dir)
    except OSError as e:
        _row("Output dir writable", False, f"{output_dir}: {e}")

    # 7. Audio Only preset present + valid (only meaningful with project).
    # Snapshot the Deliver page first: probing/creating loads the Audio Only
    # preset, and the user shouldn't find their export page stuck on it.
    if project:
        saved_fmt = _safe(project.GetCurrentRenderFormatAndCodec, _default={}) or {}
        saved_mode = _safe(project.GetCurrentRenderMode)
        try:
            presets = project.GetRenderPresetList() or []
            if _AUDIO_PRESET_NAME in presets:
                _row("'Audio Only' preset exists", True)
                # Validate its settings by adding a test job, inspecting, deleting
                existing_jobs = project.GetRenderJobList() or []
                pre_existing_ids = {j.get("JobId") for j in existing_jobs if isinstance(j, dict) and j.get("JobId")}
                project.LoadRenderPreset(_AUDIO_PRESET_NAME)
                # Resolve 21: AddRenderJob returns no id unless a target dir
                # is set. Use the temp dir; we delete the probe job anyway.
                _safe(project.SetRenderSettings, {"TargetDir": tempfile.gettempdir()})
                test_job_id = None
                try:
                    test_job_id = project.AddRenderJob()
                    if test_job_id:
                        job_settings = {}
                        for j in (project.GetRenderJobList() or []):
                            if isinstance(j, dict) and j.get("JobId") == test_job_id:
                                job_settings = j
                                break
                        problems = _validate_audio_only_settings(job_settings)
                        if problems:
                            _row("'Audio Only' preset settings", False, "; ".join(problems))
                            warnings.append(
                                "Recreate the preset: ./.venv/bin/python create_audio_only_preset.py --force"
                            )
                        else:
                            _row("'Audio Only' preset settings", True)
                    else:
                        _row("'Audio Only' preset settings", False, "AddRenderJob returned no id")
                finally:
                    # Delete the probe even if validation raises, or it
                    # leaks into the user's render queue.
                    _delete_job_if_ours(project, test_job_id, pre_existing_ids)
            else:
                log.info(f"  [....] 'Audio Only' preset not found, attempting to create...")
                try:
                    pair = _set_wav_format_compat(project)
                    if not pair:
                        # Resolve 21 removed audio formats from
                        # SetCurrentRenderFormatAndCodec entirely, but it also
                        # ships 'Audio Only' as a factory preset, so this
                        # branch (preset missing AND can't create) should be
                        # unreachable there. Reaching it means something odd.
                        _row(
                            "Create 'Audio Only' preset", False,
                            "could not switch Deliver page to WAV. On Resolve 21 "
                            "'Audio Only' is a factory preset and should already exist; "
                            "check the Deliver page's preset list.",
                        )
                    else:
                        settings = {
                            "ExportVideo": False,
                            "ExportAudio": True,
                            "AudioBitDepth": 24,
                            "AudioSampleRate": 48000,
                        }
                        if pair[1]:
                            settings["AudioCodec"] = pair[1]
                        project.SetRenderSettings(settings)
                        if project.SaveAsNewRenderPreset(_AUDIO_PRESET_NAME):
                            _row("Create 'Audio Only' preset", True, f"format/codec: {pair[0]}/{pair[1] or '(default)'}")
                        else:
                            _row("Create 'Audio Only' preset", False, "SaveAsNewRenderPreset failed")
                except Exception as e:
                    _row("Create 'Audio Only' preset", False, str(e))
        finally:
            _restore_deliver_state(project, saved_fmt, saved_mode)

    log.info("")
    if failures:
        log.error(f"Pre-flight FAILED on: {', '.join(failures)}")
        for w in warnings:
            log.warning(w)
        return 1
    log.info("Pre-flight OK. Tool is ready to use.")
    for w in warnings:
        log.warning(w)
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
    parser.add_argument(
        "--check",
        action="store_true",
        help="Run pre-flight environment + Resolve checks, create 'Audio Only' "
             "preset if missing, then exit. Run after install/Resolve update.",
    )
    parser.add_argument(
        "--no-ui",
        action="store_true",
        help="Don't spawn the floating progress window.",
    )

    args = parser.parse_args()

    # Everything that touches Resolve's Deliver page or the status file is
    # exclusive -- including --check, which loads presets and adds probe
    # jobs and would race a live caption run.
    if not _acquire_run_lock():
        # Do NOT write the shared status file here: the only reader is the
        # WINNING run's progress window, and an "error" status would make it
        # show a false failure and auto-close, killing its Cancel button.
        log.error("Another caption run is in progress. Wait for it or cancel it from its progress window.")
        return 1
    try:
        if args.check:
            return run_check_mode(args)
        if args.file:
            return run_file_mode(args)
        else:
            return run_resolve_mode(args)
    finally:
        _release_run_lock()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log.info("Cancelled.")
        sys.exit(1)
    except Exception:
        log.exception("Fatal error")
        sys.exit(1)
