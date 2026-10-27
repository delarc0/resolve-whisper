"""The pre-flight health check (`caption.py --check`).

Run once after install or after a Resolve update. It reports every step
PASS or FAIL with the remedy, writes its own log file that install agents
are told to read, and creates the "Audio Only" render preset if it is
missing. It never leaves the user's Deliver page or render queue changed.
"""
import logging
import os
import sys
import tempfile

import platforminfo
from resolve_api import (
    _AUDIO_PRESET_NAME,
    _audio_codec_warning,
    _delete_job_if_ours,
    _resolve_major,
    _restore_deliver_state,
    _restore_page,
    _safe,
    _set_wav_format_compat,
    _validate_audio_only_settings,
    get_resolve,
)

log = logging.getLogger(__name__)


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

    # 3. ffmpeg. Resolve launches us with a bare PATH, so look where the
    # installers actually put it rather than only on PATH -- plain `which`
    # failed people whose ffmpeg sits in Homebrew or C:\ffmpeg\bin.
    ffmpeg_path = platforminfo.find_tool("ffmpeg")
    if ffmpeg_path:
        _row("ffmpeg available", True, ffmpeg_path)
    elif platforminfo.IS_WIN:
        # Windows transcribes through faster-whisper, which decodes audio
        # itself, and setup.ps1 deliberately treats ffmpeg as optional.
        # Failing the whole check here told those users their install was
        # broken when it was working.
        log.info("  [WARN] ffmpeg not found - optional on Windows")
        warnings.append(
            "ffmpeg not found. Optional on Windows, but it is the fallback "
            "decoder when Resolve renders a format the built-in one cannot "
            "open. Install: winget install Gyan.FFmpeg")
    else:
        _row("ffmpeg available", False, "install: brew install ffmpeg")

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
    output_dir = args.output_dir or platforminfo.default_output_dir()
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
        _snapshot_failed = True
        _output_mutated = False
        _probe_id = _safe(project.AddRenderJob)
        if _probe_id:
            try:
                for j in (_safe(project.GetRenderJobList, _default=[]) or []):
                    if isinstance(j, dict) and j.get("JobId") == _probe_id:
                        saved_target_dir = j.get("TargetDir")
                        saved_custom_name = j.get("CustomName")
                        _snapshot_failed = False
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
                _output_mutated = True
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
                            codec_note = _audio_codec_warning(job_settings, _safe(get_resolve))
                            if codec_note:
                                log.info(f"  [WARN] 'Audio Only' preset codec - {codec_note}")
                                if _resolve_major(resolve) < 21:
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
                log.info("  [....] 'Audio Only' preset not found, creating it...")
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
                                   saved_target_dir, saved_custom_name,
                                   _snapshot_failed and _output_mutated)
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
