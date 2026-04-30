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
import sys

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
        probe_id = project.AddRenderJob()
        if not probe_id:
            return ["AddRenderJob (probe) returned no id"]

        job_settings = {}
        for j in (project.GetRenderJobList() or []):
            if isinstance(j, dict) and j.get("JobId") == probe_id:
                job_settings = j
                break

        problems = []
        if not isinstance(job_settings, dict) or not job_settings:
            return ["could not read probe job settings"]
        # ExportVideo MUST be falsey
        ev = job_settings.get("ExportVideo")
        if ev not in (False, 0, "0", "false", "False", None):
            problems.append(f"ExportVideo is {ev!r}, want false")
        ea = job_settings.get("ExportAudio")
        if ea not in (True, 1, "1", "true", "True", None):
            problems.append(f"ExportAudio is {ea!r}, want true")
        return problems
    finally:
        if probe_id and probe_id not in pre_ids:
            try:
                project.DeleteRenderJob(probe_id)
            except Exception:
                pass


def _create(project) -> int:
    """Create the preset. Caller must have ensured it doesn't already exist."""
    log.info("Snapshotting current render format/codec...")
    saved_fmt = project.GetCurrentRenderFormatAndCodec() or {}

    try:
        log.info("Configuring audio-only WAV/PCM...")
        if not project.SetCurrentRenderFormatAndCodec("wav", "LinearPCM"):
            log.error("SetCurrentRenderFormatAndCodec failed (wav/LinearPCM).")
            return 1

        ok = project.SetRenderSettings({
            "ExportVideo": False,
            "ExportAudio": True,
            "AudioCodec": "LinearPCM",
            "AudioBitDepth": 24,
            "AudioSampleRate": 48000,
        })
        if not ok:
            log.warning("SetRenderSettings returned falsy; continuing anyway.")

        log.info(f"Saving preset '{PRESET_NAME}'...")
        if not project.SaveAsNewRenderPreset(PRESET_NAME):
            log.error("SaveAsNewRenderPreset failed.")
            return 1
        log.info(f"Preset '{PRESET_NAME}' created.")
    finally:
        log.info("Restoring previous render format/codec...")
        if isinstance(saved_fmt, dict):
            fmt = saved_fmt.get("format")
            codec = saved_fmt.get("codec")
            if fmt and codec:
                try:
                    project.SetCurrentRenderFormatAndCodec(fmt, codec)
                except Exception:
                    pass
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true",
                        help="Delete existing preset and recreate it.")
    args = parser.parse_args()

    try:
        import DaVinciResolveScript as bmd
    except ImportError:
        log.error("DaVinciResolveScript not on PYTHONPATH. Launch via the Lua wrapper.")
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

    existing = project.GetRenderPresetList() or []
    already_exists = PRESET_NAME in existing

    if already_exists and args.force:
        log.info(f"--force given, deleting '{PRESET_NAME}'...")
        try:
            ok = project.DeleteRenderPreset(PRESET_NAME)
            if not ok:
                log.error(f"DeleteRenderPreset('{PRESET_NAME}') returned falsy.")
                return 1
        except Exception as e:
            log.error(f"DeleteRenderPreset failed: {e}")
            return 1
        already_exists = False

    if already_exists:
        log.info(f"'{PRESET_NAME}' exists; validating...")
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
    problems = _validate(project)
    if problems:
        log.error(f"Newly-created preset is misconfigured: {'; '.join(problems)}")
        return 1
    log.info("Done. Preset created and validated.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
