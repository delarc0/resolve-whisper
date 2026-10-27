"""Everything in the pipeline that talks to DaVinci Resolve.

Kept apart from caption.py so the Resolve-shaped code - the compat shims
that keep 18 through 21 working, and the borrowed Deliver page with its
restore paths - can be read on its own. This module owns no UI: the one
function that needs to report progress, render_audio, takes an on_status
callback and the caller decides what that means.
"""
import logging
import os
import time

log = logging.getLogger(__name__)


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


def _resolve_major(resolve) -> int:
    """Major Resolve version, or 0 if it can't be determined."""
    try:
        v = resolve.GetVersionString()
        return int(str(v).split(".")[0])
    except Exception:
        return 0


def _audio_codec_warning(settings: dict, resolve=None) -> str:
    """Return a warning string when the render audio codec isn't PCM/WAV, else "".

    Non-PCM (e.g. Resolve 21's factory AAC) is not fatal - duration, VAD, and
    transcription all decode via ffmpeg/ffprobe - but WAV skips those fallback
    decodes, so it's worth surfacing instead of burying at info level.

    On Resolve 21 it is not worth surfacing as something to FIX. 'Audio Only'
    is a factory preset there, so DeleteRenderPreset refuses it, and
    SetCurrentRenderFormatAndCodec no longer accepts audio formats at all
    (both verified against 21.0.3.7). The old text told users to recreate the
    preset as WAV, which cannot succeed on 21 -- an instruction that only
    wastes the time of whoever follows it.
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
        if resolve is not None and _resolve_major(resolve) >= 21:
            return (
                f"audio codec is {audio_codec!r} - this is Resolve 21's "
                "factory 'Audio Only' preset, decoded via ffmpeg. Expected; "
                "nothing to fix."
            )
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
                           saved_target_dir=None, saved_custom_name=None,
                           clear_unknown_output: bool = False):
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
    # thing being cleared.
    if clear_unknown_output:
        # Set ONLY when we actually overwrote TargetDir and the probe never
        # told us the old value. Leaving OURS in place would silently
        # redirect the user's next export into the captions folder, which is
        # worse than an empty field they get prompted to fill. Clearing when
        # we never touched the fields would be destroying settings for
        # nothing, so both conditions have to hold.
        log.warning("Could not read your Deliver output path before changing "
                    "it; clearing it rather than leaving the tool's path there.")
        _safe(project.SetRenderSettings, {"TargetDir": "", "CustomName": ""})
    elif saved_target_dir is not None or saved_custom_name is not None:
        _safe(project.SetRenderSettings, {
            "TargetDir": saved_target_dir or "",
            "CustomName": saved_custom_name or "",
        })


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
        import platforminfo
        audio_dur = platforminfo.audio_duration_s(audio_path)
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


def render_audio(project, timeline, output_dir: str, fps: float = 0.0,
                 on_status=None):
    """Render timeline audio via the regular render queue.

    `on_status(stage, message, progress=...)` is how the caller keeps its
    progress window alive during a long render; this module never touches
    the status file itself. Omit it and the render runs silent.

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
    status = on_status or (lambda *a, **k: None)
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
    snapshot_failed = True
    output_mutated = False   # did we actually overwrite TargetDir/CustomName?
    probe_id = _safe(project.AddRenderJob)
    if probe_id:
        try:
            for j in (_safe(project.GetRenderJobList, _default=[]) or []):
                if isinstance(j, dict) and j.get("JobId") == probe_id:
                    saved_target_dir = j.get("TargetDir")
                    saved_custom_name = j.get("CustomName")
                    snapshot_failed = False
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
                               saved_target_dir, saved_custom_name,
                               snapshot_failed and output_mutated)
        return None, 0.0

    output_mutated = True
    if not _safe(project.SetRenderSettings, {
        "TargetDir": output_dir,
        "CustomName": wav_name,
    }):
        log.warning("SetRenderSettings returned falsy; continuing.")

    job_id = _safe(project.AddRenderJob)
    if not job_id:
        log.error("AddRenderJob returned no id.")
        _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name,
                               snapshot_failed and output_mutated)
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
                               saved_target_dir, saved_custom_name,
                               snapshot_failed and output_mutated)
        return None, 0.0
    codec_note = _audio_codec_warning(job_settings, _safe(get_resolve))
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
                        status("error", "Lost connection to Resolve.")
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
            status("rendering_audio", f"{pct_int}% complete", progress=pct_int)

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
                               saved_target_dir, saved_custom_name,
                               snapshot_failed and output_mutated)
