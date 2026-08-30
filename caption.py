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


# Per-run status file: the progress UI is told its exact path via argv. A
# single shared path let a lingering window from a previous run adopt the
# NEXT run's pid and cancel it (cancel -> relaunch is a common sequence).
STATUS_FILE = os.path.join(tempfile.gettempdir(),
                           f"resolve_whisper_status.{os.getpid()}.json")
LOCK_FILE = os.path.join(tempfile.gettempdir(), "resolve_whisper.lock")

# Held open for the process lifetime; the kernel drops the lock when this
# process dies for ANY reason, which is what makes the scheme crash-proof.
_LOCK_FD = None


def _acquire_run_lock() -> bool:
    """Claim the single-run lock, or report that another run holds it.

    Uses a KERNEL-held advisory lock (fcntl.flock / msvcrt.locking) on a
    long-lived fd rather than the existence of a file. The kernel releases it
    when the process exits however it exits -- normal return, exception,
    SIGKILL, TerminateProcess, power loss -- so a crashed or force-quit run
    can never wedge the tool. That also removes the PID-liveness probe
    entirely, which had three separate failure modes:
      - Windows had no safe probe (os.kill(pid, 0) TERMINATES the target), so
        it fell back to a 2h age heuristic that locked the tool out for two
        hours after any hard-killed run;
      - a recycled PID could make a dead lock look alive forever;
      - the create-then-write gap and the unlink-by-path retry let two runs
        acquire the "exclusive" lock simultaneously.
    The pid is still written into the file, but only as a human-readable
    breadcrumb for logs -- never as the liveness signal.
    """
    global _LOCK_FD
    if _LOCK_FD is not None:
        return True  # already held by this process; never orphan the old fd
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_RDWR, 0o644)
    except OSError as e:
        log.warning(f"Could not open lock file ({e}); proceeding without lock.")
        return True

    try:
        if sys.platform == "win32":
            import msvcrt
            # Lock 1 byte, non-blocking. Raises OSError if already held.
            msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except (OSError, ImportError) as e:
        # ImportError: no locking primitive available -> don't block the user.
        if isinstance(e, ImportError):
            log.warning("No file-locking primitive available; proceeding without lock.")
            _LOCK_FD = fd
            return True
        holder = ""
        try:
            os.lseek(fd, 0, os.SEEK_SET)
            holder = os.read(fd, 32).decode("utf-8", "replace").strip()
        except OSError:
            pass
        os.close(fd)
        if holder:
            log.error(f"Another caption run is in progress (pid {holder}). "
                      "Wait for it to finish or cancel it from its progress window.")
        else:
            log.error("Another caption run is in progress. "
                      "Wait for it to finish or cancel it from its progress window.")
        return False

    # We own it: record our pid for the log breadcrumb.
    try:
        os.ftruncate(fd, 0)
        os.lseek(fd, 0, os.SEEK_SET)
        os.write(fd, str(os.getpid()).encode("utf-8"))
    except OSError:
        pass
    _LOCK_FD = fd
    return True


def _release_run_lock():
    """Release the run lock. Best-effort; never raises.

    Closing the fd drops the kernel lock. The file itself is left behind on
    purpose: unlinking it races with another process that already has it
    open, and an empty leftover file is harmless (the lock is the fd, not
    the file's existence).
    """
    global _LOCK_FD
    if _LOCK_FD is None:
        return
    try:
        if sys.platform == "win32":
            import msvcrt
            try:
                os.lseek(_LOCK_FD, 0, os.SEEK_SET)
                msvcrt.locking(_LOCK_FD, msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
        os.close(_LOCK_FD)
    except OSError:
        pass
    finally:
        _LOCK_FD = None


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


def _sweep_stale_status_files(max_age_s: int = 86400):
    """Delete leftover per-run status files from previous runs.

    The status file carries the pid so a lingering progress window can't
    adopt the next run's pid, but that means nothing overwrites the old
    ones. Windows %TEMP% is never swept by the OS, so they would pile up
    forever. Only our own prefix, only files older than a day, never the
    one this run is using.
    """
    directory = os.path.dirname(STATUS_FILE)
    keep = os.path.basename(STATUS_FILE)
    cutoff = time.time() - max_age_s
    try:
        for name in os.listdir(directory):
            if name == keep:
                continue
            if not (name.startswith("resolve_whisper_status.")
                    and name.endswith(".json")):
                continue
            path = os.path.join(directory, name)
            try:
                if os.path.getmtime(path) < cutoff:
                    os.unlink(path)
            except OSError:
                pass
    except OSError:
        pass


def _default_output_dir() -> str:
    return os.path.join(os.path.expanduser("~"), "Desktop", "Captions")


def _cleanup_old_srts(output_dir: str, keep_days: int):
    """Delete OUR old .srt files in output_dir. Non-fatal.

    Two guards, because this deletes files:
      - only names carrying the tool's own " YYYYMMDD-HHMMSS.srt" stamp, and
      - only inside the tool's OWN output folder. The stamp alone is not
        proof of authorship ("Kundleverans FINAL 20250612-093000.srt" matches
        it), and output_dir is user-configurable -- pointing it at a delivery
        folder would have quietly destroyed month-old client subtitles.
    Every deleted filename is logged so the action is auditable.
    """
    if keep_days <= 0 or not os.path.isdir(output_dir):
        return
    if os.path.normpath(output_dir) != os.path.normpath(_default_output_dir()):
        log.debug("Custom output dir; skipping automatic SRT cleanup.")
        return
    cutoff = time.time() - keep_days * 86400
    removed = []
    try:
        for name in os.listdir(output_dir):
            if not _is_own_srt(name):
                continue
            path = os.path.join(output_dir, name)
            try:
                if os.path.getmtime(path) < cutoff:
                    os.unlink(path)
                    removed.append(name)
            except OSError:
                pass
    except OSError as e:
        log.debug(f"Could not list output dir for cleanup: {e}")
        return
    if removed:
        log.info(f"Cleaned up {len(removed)} SRT file(s) older than "
                 f"{keep_days} days: {', '.join(sorted(removed))}")


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
    # while transcription drags on. Unknown duration (0.0) means the
    # estimate would be pure fiction (instant 95% that sits for the whole
    # run); show an honest indeterminate bar instead (progress=-1).
    expected = max(audio_duration_s / 4.0, 1.0) if audio_duration_s > 0 else None

    def _run():
        t0 = time.time()
        while not stop.is_set():
            if expected is None:
                _write_status("transcribing", "Transcribing audio...",
                              progress=-1)
            else:
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


# Set by a background thread if this install is behind origin. Surfaced in
# the progress window, which is the one screen every user actually looks at:
# the same warning in the check log is invisible to anyone who runs the tool
# from Resolve's menu and never opens a log.
_UPDATE_NOTICE = ""


def _start_update_check():
    """Ask (at most once a day, off the hot path) whether an update exists.

    A daemon thread so it cannot delay or outlive the run, and every failure
    inside is swallowed: learning about an update is never worth risking a
    caption job.
    """
    def _probe():
        global _UPDATE_NOTICE
        try:
            import version
            behind = version.update_available()
            if behind > 0:
                plural = "s" if behind > 1 else ""
                _UPDATE_NOTICE = (
                    f"Update available ({behind} version{plural} behind) - "
                    "Workspace > Scripts > LAB37 Update")
        except Exception:
            pass

    try:
        import threading
        threading.Thread(target=_probe, daemon=True).start()
    except Exception as e:
        log.debug(f"update check not started: {e}")


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
                "notice": _UPDATE_NOTICE,
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
        _sweep_stale_status_files()
        # Reset status file so the new UI process doesn't pick up a stale "done"
        _write_status("starting", "Connecting to Resolve...")
        # On Windows use pythonw.exe for the Tk window so it doesn't flash a
        # console. Falls back to the normal interpreter if pythonw is absent.
        ui_python = sys.executable
        if sys.platform == "win32":
            candidate = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            if os.path.exists(candidate):
                ui_python = candidate
        return subprocess.Popen(
            [ui_python, ui_script, STATUS_FILE],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
    except Exception as e:
        log.debug(f"progress UI spawn failed: {e}")
        return None


def _run_settings_dialog():
    """Show the Custom settings dialog (settings_ui.py subprocess).

    Returns the chosen settings dict, or None if the user cancelled or the
    dialog could not run. Blocks until the user decides; runs BEFORE the run
    lock is taken so an open dialog never blocks other caption runs.
    """
    dialog_script = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "settings_ui.py")
    if not os.path.exists(dialog_script):
        log.error("settings_ui.py missing - reinstall (run setup.sh).")
        return None
    try:
        proc = subprocess.run(
            [sys.executable, dialog_script],
            capture_output=True, text=True,
        )
    except Exception as e:
        log.error(f"Could not open the settings dialog: {e}")
        return None
    if proc.returncode == 1:
        log.info("Cancelled from the settings dialog.")
        return None
    for line in proc.stdout.splitlines():
        if line.startswith("SETTINGS:"):
            try:
                return json.loads(line[len("SETTINGS:"):])
            except json.JSONDecodeError:
                break
    log.error("Settings dialog failed (no settings returned). "
              f"stderr: {proc.stderr.strip()[:200]}")
    return None


def get_resolve():
    """Connect to a running DaVinci Resolve Studio instance."""
    try:
        import platforminfo
        platforminfo.bootstrap_resolve_env()
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
# .mp4/.mov are here because Resolve 21's factory "Audio Only" preset writes
# AAC into an .mp4 container -- without them the file-resolution fallback
# misses the render entirely and the run dies after doing all the work.
_AUDIO_EXTS = (".wav", ".flac", ".mp3", ".aac", ".m4a", ".aif", ".aiff",
               ".mp4", ".mov")


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

    return problems


def _audio_codec_warning(settings: dict) -> str:
    """Return a warning string when the render audio codec isn't PCM/WAV, else "".

    Non-PCM (e.g. Resolve 21's factory AAC) is not fatal - duration, VAD, and
    transcription all decode via ffmpeg/ffprobe - but WAV skips those fallback
    decodes, so it's worth surfacing instead of burying at info level.
    """
    if not isinstance(settings, dict):
        return ""
    audio_codec = settings.get("AudioCodec", "")
    if (
        isinstance(audio_codec, str)
        and audio_codec
        and "pcm" not in audio_codec.lower()
        and "linearpcm" not in audio_codec.lower().replace(" ", "")
    ):
        return (
            f"audio codec is {audio_codec!r} - works (decoded via ffmpeg), "
            "but a WAV/PCM preset transcribes faster"
        )
    return ""


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


def _restore_page(resolve, saved_page):
    """Put the user back on the page they started on.

    Resolve switches itself to Deliver during AddRenderJob/StartRendering,
    so ANY exit before the end of the happy path (render failed, no speech,
    cancelled, Resolve lost) used to strand the user on the Deliver page
    looking at a queue. Restoring the page they were actually on beats
    hardcoding "edit": a colourist who started on Color should land on Color.
    """
    if not resolve or not saved_page:
        return
    try:
        if resolve.GetCurrentPage() != saved_page:
            resolve.OpenPage(saved_page)
    except Exception as e:
        log.debug(f"page restore failed: {e}")


def _restore_deliver_state(project, saved_fmt: dict, saved_mode,
                           saved_target_dir=None, saved_custom_name=None):
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
    # Put the user's own output location back. Blanking these unconditionally
    # (the old behaviour) destroyed the user's Deliver page path and filename
    # on every run AND every health check -- our temp dir was not the only
    # thing being cleared. If we never captured a snapshot, leave the fields
    # alone rather than guessing.
    if saved_target_dir is not None or saved_custom_name is not None:
        _safe(project.SetRenderSettings, {
            "TargetDir": saved_target_dir or "",
            "CustomName": saved_custom_name or "",
        })


def _audio_duration_s(path: str) -> float:
    """Duration of a rendered audio file in seconds, or 0.0 if unknown.

    Used to cross-check what the render ACTUALLY covered. ffprobe reads any
    container Resolve emits (including the AAC/.mp4 the 21 factory preset
    produces, which libsndfile cannot open).
    """
    try:
        import platforminfo
        ffprobe = platforminfo.find_tool("ffprobe")
        if not ffprobe:
            return 0.0
        proc = subprocess.run(
            [ffprobe, "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", path],
            capture_output=True, text=True, timeout=60,
        )
        return max(float(proc.stdout.strip()), 0.0)
    except Exception:
        return 0.0


def _render_offset_s(job_settings: dict, timeline, fps: float,
                     audio_path: str = None) -> float:
    """Seconds between the timeline start and the start of the rendered audio.

    When the render covers an in/out RANGE, the audio (and every Whisper
    timestamp derived from it) is zero-based at the IN point, while Resolve
    maps an imported SRT's 00:00:00 to the TIMELINE start. Without this
    offset every caption lands at the top of the timeline instead of over
    the speech.

    Deciding the offset is subtle, because a job dict can carry MarkIn/
    MarkOut that the render IGNORED:
      - "SelectAllFrames": True means exactly that (Resolve's own docs:
        "the settings MarkIn and MarkOut are ignored"). Shifting then would
        corrupt captions that were already correct, so it wins outright.
      - Otherwise MarkIn is used, but only after a duration cross-check:
        if the rendered audio is as long as the whole timeline, the range
        was not applied whatever the marks say.
    Anything ambiguous returns 0.0 -- placing captions from the timeline
    start is the recoverable failure; shifting them wrongly is not.
    """
    if not fps or fps <= 0:
        return 0.0

    if isinstance(job_settings, dict):
        for key in ("SelectAllFrames", "selectAllFrames"):
            if job_settings.get(key):
                return 0.0

    start_frame = _safe(timeline.GetStartFrame, _default=None)
    if start_frame is None:
        return 0.0

    mark_in = None
    if isinstance(job_settings, dict):
        for key in ("MarkIn", "markIn"):
            if key in job_settings:
                mark_in = job_settings[key]
                break

    # No MarkIn on the job means we cannot know what range was rendered.
    # timeline.GetMarkInOut() is NOT a safe fallback: it reports the LIVE
    # marks (which the user may have changed during a long render) and its
    # frame base is ambiguous versus GetStartFrame.
    try:
        offset_frames = int(mark_in) - int(start_frame)
    except (TypeError, ValueError):
        return 0.0

    # Negative means no in-point, or marks in a different frame base.
    if offset_frames <= 0:
        return 0.0
    offset = offset_frames / float(fps)

    # Cross-check against the audio we actually got: if it spans the whole
    # timeline, the marks were not applied and shifting would be wrong.
    if audio_path:
        audio_dur = _audio_duration_s(audio_path)
        end_frame = _safe(timeline.GetEndFrame, _default=None)
        if audio_dur > 0 and end_frame is not None:
            try:
                timeline_dur = (int(end_frame) - int(start_frame)) / float(fps)
            except (TypeError, ValueError):
                timeline_dur = 0.0
            # Tolerance covers container/encoder padding only (a few
            # frames), not a real in-point. A wider window suppressed
            # genuine sub-second in-points.
            tolerance = max(0.25, 3.0 / float(fps))
            if timeline_dur > 0 and abs(audio_dur - timeline_dur) <= tolerance:
                log.info("Rendered audio spans the whole timeline; "
                         "ignoring in/out marks for caption placement.")
                return 0.0
    return offset


def render_audio(project, timeline, output_dir: str, fps: float = 0.0):
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
            return None, 0.0
    if not os.access(output_dir, os.W_OK):
        log.error(f"Output dir not writable: {output_dir}")
        return None, 0.0

    presets = _safe(project.GetRenderPresetList, _default=[]) or []
    if _AUDIO_PRESET_NAME not in presets:
        log.error(f"'{_AUDIO_PRESET_NAME}' render preset not found in this project.")
        log.error("Run: python caption.py --check  (will create it)")
        return None, 0.0

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

    # There is no getter for the Deliver page's output path/filename, but a
    # throwaway probe job reports them -- so snapshot via one, then delete
    # it. Without this the restore blanked the user's own output location on
    # every run and every health check.
    saved_target_dir = None
    saved_custom_name = None
    probe_id = _safe(project.AddRenderJob)
    if probe_id:
        try:
            for j in (_safe(project.GetRenderJobList, _default=[]) or []):
                if isinstance(j, dict) and j.get("JobId") == probe_id:
                    saved_target_dir = j.get("TargetDir")
                    saved_custom_name = j.get("CustomName")
                    break
        finally:
            # Delete the probe even if the readback raises, or it leaks
            # into the user's render queue.
            _delete_job_if_ours(project, probe_id, pre_existing_ids)

    log.info(f"Loading '{_AUDIO_PRESET_NAME}' preset...")
    if not _safe(project.LoadRenderPreset, _AUDIO_PRESET_NAME):
        log.error(f"LoadRenderPreset('{_AUDIO_PRESET_NAME}') failed.")
        # A failed load can still have applied part of the preset, so put
        # the Deliver page back rather than assuming nothing changed.
        _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name)
        return None, 0.0

    if not _safe(project.SetRenderSettings, {
        "TargetDir": output_dir,
        "CustomName": wav_name,
    }):
        log.warning("SetRenderSettings returned falsy; continuing.")

    job_id = _safe(project.AddRenderJob)
    if not job_id:
        log.error("AddRenderJob returned no id.")
        _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name)
        return None, 0.0

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
        _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name)
        return None, 0.0
    codec_note = _audio_codec_warning(job_settings)
    if codec_note:
        log.warning(f"'{_AUDIO_PRESET_NAME}' preset: {codec_note}")

    log.info("Rendering timeline audio...")
    audio_path = None
    try:
        if not _start_rendering_compat(project, job_id):
            log.error("StartRendering failed across all known signatures.")
            return None, 0.0

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
                        return None, 0.0
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
                return None, 0.0
            time.sleep(0.5)
        else:
            log.error(f"Render timed out after {_RENDER_TIMEOUT_S}s.")
            _safe(project.StopRendering)
            return None, 0.0

        # Resolve output path from the job's actual settings.
        job_settings_post = {}
        for j in (_safe(project.GetRenderJobList, _default=[]) or []):
            if isinstance(j, dict) and j.get("JobId") == job_id:
                job_settings_post = j
                break
        audio_path = _expected_output_path(job_settings_post, output_dir, wav_name)
        if not audio_path or not os.path.exists(audio_path):
            log.error("Render reported complete but audio file not found.")
            return None, 0.0

        # How far into the timeline this audio starts (in/out range renders
        # are zero-based at the IN point). Merge both job snapshots: the
        # post-render copy can drop MarkIn, the pre-render one can predate a
        # setting. SelectAllFrames is a safety guard, so honour it if EITHER
        # snapshot reports it -- a wrong shift is worse than no shift.
        merged_job = dict(job_settings or {})
        merged_job.update(job_settings_post or {})
        for key in ("SelectAllFrames", "selectAllFrames"):
            if (job_settings or {}).get(key) or (job_settings_post or {}).get(key):
                merged_job[key] = True
        offset_s = _render_offset_s(merged_job, timeline, fps, audio_path)
        if offset_s > 0:
            log.info(f"In-point offset: captions shift +{offset_s:.2f}s to match the timeline.")

        try:
            size_mb = os.path.getsize(audio_path) / (1024 * 1024)
            log.info(f"Audio rendered: {audio_path} ({size_mb:.1f} MB)")
        except Exception:
            log.info(f"Audio rendered: {audio_path}")
        return audio_path, offset_s

    except KeyboardInterrupt:
        log.warning("Interrupted; stopping render and cleaning up...")
        _safe(project.StopRendering)
        raise
    finally:
        _delete_job_if_ours(project, job_id, pre_existing_ids)
        _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name)


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
    _start_update_check()
    tmp_dir = None
    resolve = None
    saved_page = None

    try:
        _write_status("starting", "Connecting to Resolve...")
        resolve = get_resolve()
        if not resolve:
            _write_status("error", "Resolve isn't running. Open Resolve Studio and try again.")
            return 1
        # Where the user was before we started. The finally puts them back;
        # rendering drags them to Deliver otherwise.
        saved_page = _safe(resolve.GetCurrentPage)

        project, timeline, fps = get_timeline_info(resolve)
        if not project:
            _write_status("error", "No project open in Resolve. Open a project first.")
            return 1
        if not timeline:
            _write_status("error", "No timeline selected. Open a timeline and set in/out points.")
            return 1

        _write_status("rendering_audio", f"Timeline: {timeline.GetName()}")
        tmp_dir = tempfile.mkdtemp(prefix="resolve_whisper_")
        wav_path, timeline_offset_s = render_audio(project, timeline, tmp_dir, fps)
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
            output_dir = _default_output_dir()
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
        # offset_s puts the captions where the speech actually is when the
        # user rendered an in/out range instead of the whole timeline.
        success = write_srt(segments, srt_path, fps, strip_punctuation=strip_punct,
                            offset_s=timeline_offset_s)
        if not success:
            _write_status("error", "SRT write failed.")
            return 1

        # Switch back to Edit page so the user lands where they expect
        # (Resolve briefly switches to Deliver during AddRenderJob in some
        # versions). Also make Edit the restore target, so the finally that
        # handles the failure paths doesn't undo this on the way out.
        saved_page = "edit"
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
        # - it targets subtitle track 1, so we only auto-place when no
        #   subtitle track holds any items. The guard is about not stacking
        #   onto someone's hand-edited captions, so it counts ITEMS, not
        #   tracks: Resolve leaves an empty 'Subtitle 1' behind as soon as
        #   anyone has touched subtitles once, and keying off track count
        #   sent those users to the Media Pool fallback for no reason.
        imported_to_pool = False
        auto_placed = False
        media_pool = _safe(project.GetMediaPool)
        pool_items = None
        if media_pool:
            pool_items = _safe(media_pool.ImportMedia, [srt_path])
            imported_to_pool = bool(pool_items)

        # AppendToTimeline targets the project's CURRENT timeline, but our
        # guard/AddTrack/verify all use the handle captured before the
        # (minutes-long) transcription. If the user switched timelines in
        # the meantime we would add a track to one timeline and append the
        # captions to another, stacking onto someone else's subtitles.
        same_timeline = True
        if imported_to_pool:
            current_tl = _safe(project.GetCurrentTimeline)
            if current_tl is not None:
                our_id = _safe(timeline.GetUniqueId)
                cur_id = _safe(current_tl.GetUniqueId)
                if our_id and cur_id and our_id != cur_id:
                    # Try to put the user's original timeline back; if that
                    # fails, fall through to the Media Pool message rather
                    # than writing to the wrong timeline.
                    same_timeline = bool(_safe(project.SetCurrentTimeline, timeline))
                    if not same_timeline:
                        log.info("Timeline changed during the run; leaving the "
                                 "SRT in the Media Pool instead of guessing.")
            if not same_timeline:
                imported_to_pool = True  # keep the Media Pool fallback path

        def _subtitle_item_count():
            """Total items across every subtitle track.

            Re-reads the track count each call because AddTrack changes it,
            and counts all tracks rather than assuming track 1 so the
            placement check holds however Resolve numbers them.
            """
            tracks = _safe(timeline.GetTrackCount, "subtitle", _default=0) or 0
            total = 0
            for idx in range(1, tracks + 1):
                items = _safe(timeline.GetItemListInTrack, "subtitle", idx,
                              _default=[]) or []
                total += len(items)
            return total

        if imported_to_pool and same_timeline:
            existing_sub_tracks = _safe(timeline.GetTrackCount, "subtitle", _default=0) or 0
            existing_items = _subtitle_item_count()
            if existing_items == 0:
                # Reuse an existing empty track: AppendToTimeline targets
                # subtitle track 1, so adding a second track here would
                # place onto a track the user isn't looking at.
                if existing_sub_tracks == 0:
                    _safe(timeline.AddTrack, "subtitle")
                _safe(media_pool.AppendToTimeline, list(pool_items))
                placed = _subtitle_item_count()
                auto_placed = placed > 0
                if auto_placed:
                    log.info(f"Captions placed on subtitle track ({placed} items).")
                else:
                    log.info("AppendToTimeline placed nothing (timeline locked by "
                             "a collaborator?); falling back to the Media Pool.")
            else:
                log.info(f"Timeline already has {existing_items} subtitle item(s); "
                         "not auto-placing to avoid mixing with existing captions.")

        log.info("")
        log.info(f"SRT saved to: {srt_path}")
        if auto_placed:
            _write_status("done", "Captions are on your timeline.", progress=100)
        elif imported_to_pool:
            # ImportMedia returns a truthy handle, but on Resolve 21.0.4 the
            # imported .srt has been observed NOT to show up when walking the
            # pool, so "find it in the Media Pool" can be a dead end. Give the
            # route that always works (the file is on disk) first.
            log.info("SRT imported into the Media Pool.")
            log.info("If you can see it there: right-click > Insert Selected "
                     "Subtitles to Timeline.")
            log.info("If you can't find it: File > Import > Subtitle, and pick "
                     f"{srt_path}")
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
        # Put the user back where they started. Unconditional: the whole
        # point is the failure paths, where Resolve has parked them on
        # Deliver and nothing else would move them off it.
        _restore_page(resolve, saved_page)
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

    # Write the check results to a dedicated file directly from Python. The
    # Lua preset also redirects stdout to a log, but relying on that is
    # fragile (Windows `start /b` redirection is unreliable), and installer
    # agents are told to read resolve_whisper_check.log specifically. Owning
    # the file here makes that instruction work on every platform.
    import platforminfo
    check_log_path = os.path.join(platforminfo.log_dir(),
                                  "resolve_whisper_check.log")
    check_handler = None
    try:
        check_handler = logging.FileHandler(check_log_path, mode="w",
                                            encoding="utf-8")
        check_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                              datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(check_handler)
    except OSError:
        check_handler = None  # non-fatal; stdout redirect still captures it

    def _row(name, ok, detail=""):
        prefix = "PASS" if ok else "FAIL"
        log.info(f"  [{prefix}] {name}{(' - ' + detail) if detail else ''}")
        if not ok:
            failures.append(name)

    log.info("=== Resolve Whisper pre-flight check ===")
    try:
        import version
        log.info(f"  [INFO] build {version.version_string()}")
        # Costs one network round trip, capped at 15s and silent
        # offline. Without the fetch this is always 0.
        behind = version.behind_by(fetch=True)
        if behind > 0:
            warnings.append(
                f"{behind} update(s) behind origin -- run ./update.sh "
                "(Mac) or update.ps1 (Windows)")
    except Exception as e:
        log.debug(f"version lookup failed: {e}")

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
        import platforminfo
        platforminfo.bootstrap_resolve_env()
        import DaVinciResolveScript as _bmd  # noqa: F401
        _row("DaVinciResolveScript importable", True)
    except Exception as e:
        _row("DaVinciResolveScript importable", False, str(e))

    # 5. Connect to Resolve (only if module imported)
    resolve = None
    saved_page = None
    if "DaVinciResolveScript importable" not in failures:
        resolve = get_resolve()
        _row("Connect to Resolve", bool(resolve), "is Resolve running?")
        # The preset probe below queues render jobs, which switches Resolve
        # to the Deliver page. A health check must not relocate the user.
        if resolve:
            saved_page = _safe(resolve.GetCurrentPage)

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
        or _default_output_dir()
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
        # Snapshot the user's output path/filename before the probe below
        # overwrites TargetDir -- a health check must not edit their project.
        saved_target_dir = None
        saved_custom_name = None
        _probe_existing = _safe(project.GetRenderJobList, _default=[]) or []
        _probe_ids = {j.get("JobId") for j in _probe_existing
                      if isinstance(j, dict) and j.get("JobId")}
        _probe_id = _safe(project.AddRenderJob)
        if _probe_id:
            try:
                for j in (_safe(project.GetRenderJobList, _default=[]) or []):
                    if isinstance(j, dict) and j.get("JobId") == _probe_id:
                        saved_target_dir = j.get("TargetDir")
                        saved_custom_name = j.get("CustomName")
                        break
            finally:
                # Delete the probe even if the readback raises, or it leaks
                # into the user's render queue.
                _delete_job_if_ours(project, _probe_id, _probe_ids)
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
                            codec_note = _audio_codec_warning(job_settings)
                            if codec_note:
                                log.info(f"  [WARN] 'Audio Only' preset codec - {codec_note}")
                                warnings.append(
                                    "For faster transcription, recreate the preset as WAV: "
                                    "./.venv/bin/python create_audio_only_preset.py --force"
                                )
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
            _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name)
            _restore_page(resolve, saved_page)

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


def _attach_windows_run_log():
    """On Windows, own the run log from Python.

    The Lua launcher discards shell output to NUL there (`start /b`
    redirection to a real file is unreliable for a detached child), so we
    write the run log ourselves. On Mac the launcher's stdout redirect owns
    the log and this is a no-op, keeping the proven Mac path unchanged.
    """
    import platforminfo
    if not platforminfo.IS_WIN:
        return
    try:
        path = os.path.join(platforminfo.log_dir(), "resolve_whisper.log")
        handler = logging.FileHandler(path, mode="w", encoding="utf-8")
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                              datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(handler)
    except OSError:
        pass  # non-fatal; status JSON still drives the UI


def main():
    _attach_windows_run_log()
    # Stamp the build into every run log. Bug reports arrive as a log file
    # from someone else's machine; without this the first question is always
    # "which version are you on?" and the answer is usually a guess.
    try:
        import version
        log.info(f"resolve-whisper build {version.version_string()}")
    except Exception as e:
        log.debug(f"version stamp unavailable: {e}")
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
    parser.add_argument(
        "--dialog",
        action="store_true",
        help="Open the settings dialog first (the LAB37 Custom preset). "
             "Chosen settings override config and are remembered.",
    )

    args = parser.parse_args()

    if args.dialog:
        from config import cfg
        settings = _run_settings_dialog()
        if settings is None:
            return 0
        args.language = str(settings.get("language", "auto"))
        args.max_words = int(settings.get("max_words", 0))
        args.max_chars = int(settings.get("max_chars", 42))
        if settings.get("strip_punctuation"):
            args.strip_punctuation = True
        cfg["uppercase"] = bool(settings.get("uppercase", True))
        log.info(f"Custom settings: {settings}")

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
