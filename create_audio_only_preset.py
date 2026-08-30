"""
One-shot helper: creates / verifies the 'Audio Only' render preset in the
currently-open Resolve project so caption.py's render path can use it.

Usage:
    python create_audio_only_preset.py            # create if missing, verify if present
    python create_audio_only_preset.py --force    # delete and recreate

Snapshots existing render settings, switches to WAV/PCM audio-only,
SaveAsNewRenderPreset('Audio Only'), then validates by adding a probe job
and inspecting its settings.
"""
import argparse
import logging
import os
import sys
import tempfile

# Shared helpers live in caption.py (same dir). Resolve's Lua launcher cds
# here before invoking us, but belt-and-braces for manual runs from elsewhere.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from caption import (  # noqa: E402
    _safe,
    _set_wav_format_compat,
    _validate_audio_only_settings,
    _restore_deliver_state,
    _restore_page,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

PRESET_NAME = "Audio Only"


def _validate(project) -> list:
    """Add a probe job, inspect its rendered settings, delete the probe.

    Returns a list of human-readable problems (empty list = preset is valid).
    """
    existing = project.GetRenderJobList() or []
    pre_ids = {j.get("JobId") for j in existing if isinstance(j, dict) and j.get("JobId")}

    probe_id = None
    try:
        if not project.LoadRenderPreset(PRESET_NAME):
            return [f"LoadRenderPreset('{PRESET_NAME}') failed"]
        # Resolve 21: AddRenderJob returns no id unless a target dir is set.
        _safe(project.SetRenderSettings, {"TargetDir": tempfile.gettempdir()})
        probe_id = project.AddRenderJob()
        if not probe_id:
            return ["AddRenderJob (probe) returned no id"]

        job_settings = {}
        for j in (project.GetRenderJobList() or []):
            if isinstance(j, dict) and j.get("JobId") == probe_id:
                job_settings = j
                break
        if not isinstance(job_settings, dict) or not job_settings:
            return ["could not read probe job settings"]
        return _validate_audio_only_settings(job_settings)
    finally:
        if probe_id and probe_id not in pre_ids:
            try:
                project.DeleteRenderJob(probe_id)
            except Exception:
                pass


def _create(project) -> int:
    """Create the preset. Caller must have ensured it doesn't already exist."""
    log.info("Configuring audio-only WAV/PCM...")
    pair = _set_wav_format_compat(project)
    if not pair:
        log.error("Could not switch to WAV: all known format/codec ids failed.")
        log.error("Note: Resolve 21 removed audio formats from SetCurrentRenderFormatAndCodec")
        log.error("but ships 'Audio Only' as a factory preset, so creation shouldn't be")
        log.error("needed there. Check the Deliver page's preset list.")
        return 1
    log.info(f"WAV format set via format/codec: {pair[0]}/{pair[1] or '(default)'}")

    settings = {
        "ExportVideo": False,
        "ExportAudio": True,
        "AudioBitDepth": 24,
        "AudioSampleRate": 48000,
    }
    if pair[1]:
        settings["AudioCodec"] = pair[1]
    if not project.SetRenderSettings(settings):
        log.warning("SetRenderSettings returned falsy; continuing anyway.")

    log.info(f"Saving preset '{PRESET_NAME}'...")
    if not project.SaveAsNewRenderPreset(PRESET_NAME):
        log.error("SaveAsNewRenderPreset failed.")
        return 1
    log.info(f"Preset '{PRESET_NAME}' created.")
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true",
                        help="Delete existing preset and recreate it.")
    args = parser.parse_args()

    # Resolve's scripting module is not on the default path. caption.py has
    # bootstrapped it for years; this script never did, so the remedy the
    # health check prints -- "recreate it with create_audio_only_preset.py
    # --force" -- failed immediately for anyone who followed it.
    try:
        import platforminfo
        platforminfo.bootstrap_resolve_env()
    except Exception as e:
        log.debug(f"resolve env bootstrap failed: {e}")

    try:
        import DaVinciResolveScript as bmd
    except ImportError:
        log.error("Could not import DaVinciResolveScript (is Resolve installed?).")
        return 1

    resolve = bmd.scriptapp("Resolve")
    if not resolve:
        log.error("Could not connect to Resolve.")
        return 1

    pm = resolve.GetProjectManager()
    if not pm:
        log.error("Could not get ProjectManager.")
        return 1
    project = pm.GetCurrentProject()
    if not project:
        log.error("No project open.")
        return 1

    # Snapshot the Deliver page so the user isn't left stuck on Audio Only
    # (or our temp TargetDir) after we probe/create.
    saved_fmt = _safe(project.GetCurrentRenderFormatAndCodec, _default={}) or {}
    saved_mode = _safe(project.GetCurrentRenderMode)
    saved_page = _safe(resolve.GetCurrentPage)

    # There is no getter for the Deliver page's output path/filename, so read
    # them off a throwaway probe job. _validate() below overwrites TargetDir
    # with a temp dir to make AddRenderJob work on Resolve 21, and without
    # this snapshot that temp path is what the user finds in their Deliver
    # page afterwards -- their own output location, silently replaced.
    saved_target_dir = None
    saved_custom_name = None
    _pre = _safe(project.GetRenderJobList, _default=[]) or []
    _pre_ids = {j.get("JobId") for j in _pre if isinstance(j, dict) and j.get("JobId")}
    snapshot_failed = True
    output_mutated = False   # _validate() is what overwrites TargetDir
    _probe_id = _safe(project.AddRenderJob)
    if _probe_id:
        try:
            for j in (_safe(project.GetRenderJobList, _default=[]) or []):
                if isinstance(j, dict) and j.get("JobId") == _probe_id:
                    saved_target_dir = j.get("TargetDir")
                    saved_custom_name = j.get("CustomName")
                    snapshot_failed = False
                    break
        finally:
            if _probe_id not in _pre_ids:
                _safe(project.DeleteRenderJob, _probe_id)
    try:
        existing = project.GetRenderPresetList() or []
        already_exists = PRESET_NAME in existing

        if already_exists and args.force:
            log.info(f"--force given, deleting '{PRESET_NAME}'...")
            try:
                ok = project.DeleteRenderPreset(PRESET_NAME)
                if not ok:
                    # On Resolve 21 'Audio Only' is a FACTORY preset: it
                    # cannot be deleted, and SetCurrentRenderFormatAndCodec
                    # no longer accepts audio formats, so there is no way to
                    # build a WAV replacement through the API. Say that,
                    # instead of reporting a falsy return nobody can act on.
                    try:
                        major = int(str(resolve.GetVersionString()).split(".")[0])
                    except Exception:
                        major = 0
                    if major >= 21:
                        log.error(
                            f"'{PRESET_NAME}' is a factory preset on Resolve "
                            f"{major} and cannot be deleted or replaced.")
                        log.error(
                            "Nothing to do: the tool renders through it and "
                            "decodes the AAC with ffmpeg. This is expected.")
                    else:
                        log.error(f"DeleteRenderPreset('{PRESET_NAME}') returned falsy.")
                    return 1
            except Exception as e:
                log.error(f"DeleteRenderPreset failed: {e}")
                return 1
            already_exists = False

        if already_exists:
            log.info(f"'{PRESET_NAME}' exists; validating...")
            output_mutated = True
            problems = _validate(project)
            if problems:
                log.error(f"Existing preset is misconfigured: {'; '.join(problems)}")
                log.error("Re-run with --force to recreate it.")
                return 1
            log.info(f"Validated: '{PRESET_NAME}' is correctly configured.")
            return 0

        rc = _create(project)
        if rc != 0:
            return rc

        log.info("Validating new preset...")
        output_mutated = True
        problems = _validate(project)
        if problems:
            log.error(f"Newly-created preset is misconfigured: {'; '.join(problems)}")
            return 1
        log.info("Done. Preset created and validated.")
        return 0
    finally:
        _restore_deliver_state(project, saved_fmt, saved_mode,
                               saved_target_dir, saved_custom_name,
                               snapshot_failed and output_mutated)
        _restore_page(resolve, saved_page)


if __name__ == "__main__":
    sys.exit(main())
