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
    import platforminfo
    if os.path.normpath(output_dir) != os.path.normpath(
            platforminfo.default_output_dir()):
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
                # No detail line: the title already says "Transcribing
                # audio", and with no duration there is no percentage to
                # add. Repeating the title reads as a stuck window.
                _write_status("transcribing", "", progress=-1)
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


def _apply_cli_overrides(args):
    """Let the preset's flags win over caption_config.json for this run.

    "--language auto" means explicit auto-detect: cfg["language"] defaults
    to "sv", so the Auto preset has to be able to push it back to None.
    """
    from config import cfg
    if args.language:
        cfg["language"] = (None if args.language.lower() in ("auto", "none")
                           else args.language)
    if args.max_words is not None:
        cfg["max_words_per_caption"] = args.max_words
    if args.max_chars is not None:
        cfg["max_chars_per_line"] = args.max_chars
    if args.max_lines is not None:
        cfg["max_lines"] = args.max_lines


def _transcribe_to_srt(audio_path, srt_path, fps, strip_punct,
                       offset_s=0.0, empty_message="No speech detected."):
    """Load the model, transcribe, write the SRT. Returns 0 or 1.

    The half of the pipeline both run modes share. It writes its own error
    status on every way out, so the caller only has to decide what to do
    once the file exists.
    """
    from transcribe import Transcriber, TranscriptionError
    from srt import write_srt

    _write_status("loading_model", "Warming up Whisper...")
    log.info("Loading Whisper model...")
    t0 = time.time()
    # Heartbeat: the first-ever run downloads ~3GB inside Transcriber(),
    # far beyond the UI's 90s stale-file auto-close.
    hb_stop, hb_th = _start_heartbeat(
        "loading_model", "Loading AI model (first run downloads ~3 GB)...")
    try:
        transcriber = Transcriber()
    except Exception as e:
        msg = str(e).lower()
        if "huggingface" in msg or "connection" in msg or "no such file" in msg:
            _write_status("error", "Model download failed. Check internet "
                                   "connection and retry.")
        else:
            _write_status("error", f"Whisper model failed to load: {e}")
        return 1
    finally:
        hb_stop.set()
        hb_th.join(timeout=1.0)
    log.info(f"Model loaded in {time.time() - t0:.1f}s")

    _write_status("transcribing", "Listening to audio...", progress=0)
    log.info("Transcribing...")

    # Background progress estimator: mlx_whisper.transcribe blocks without
    # callbacks on Mac, so we fake-advance the bar based on elapsed time.
    # On Windows faster_whisper still calls _on_progress for real updates.
    prog_stop, prog_th = _start_progress_estimator(
        transcriber.get_audio_duration(audio_path))

    def _on_progress(pct):
        # Real progress from faster_whisper -- stop the estimator, use this.
        if not prog_stop.is_set():
            prog_stop.set()
        try:
            pct_int = max(0, min(int(pct), 100))
        except (ValueError, TypeError):
            return
        _write_status("transcribing", f"{pct_int}% complete", progress=pct_int)

    t0 = time.time()
    try:
        segments = transcriber.transcribe(audio_path, on_progress=_on_progress)
    except TranscriptionError as e:
        # Whisper itself failed: no GPU memory, a half-downloaded model,
        # audio it could not decode. This used to fall through as an empty
        # result, which sent people off to check their in/out points.
        log.error(f"Transcription failed: {e}")
        _write_status("error", f"Transcription failed: {e}")
        return 1
    finally:
        prog_stop.set()
        prog_th.join(timeout=1.0)

    log.info(f"Transcribed {sum(len(s.words) for s in segments)} words "
             f"in {time.time() - t0:.1f}s")
    if not segments:
        log.warning(empty_message)
        _write_status("error", empty_message)
        return 1

    _write_status("writing_srt", "Writing captions to disk...")
    if not write_srt(segments, srt_path, fps, strip_punctuation=strip_punct,
                     offset_s=offset_s):
        _write_status("error", "SRT write failed.")
        return 1
    return 0


def _ensure_terminal_status(ui_proc):
    """Never leave the progress window spinning after an early return.

    The UI closes itself once it sees 'done' or 'error', so a `return 1`
    that wrote no terminal status would hang it until the 90s stale timeout.
    """
    if ui_proc is None or ui_proc.poll() is not None:
        return
    try:
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            cur = json.load(f)
        if cur.get("stage") not in ("done", "error"):
            _write_status("error", "Pipeline ended unexpectedly.")
    except Exception:
        pass


def run_resolve_mode(args):
    """Full pipeline: Resolve timeline -> audio -> transcribe -> SRT."""
    import shutil
    from config import cfg
    from resolve_api import (get_resolve, get_timeline_info, render_audio,
                             _restore_page, _safe)

    _apply_cli_overrides(args)

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
        wav_path, timeline_offset_s = render_audio(
            project, timeline, tmp_dir, fps, on_status=_write_status)
        if not wav_path:
            _write_status(
                "error",
                "Audio render failed. Check in/out points and the 'Audio Only' preset (run LAB37 Check).",
            )
            return 1

        strip_punct = getattr(args, "strip_punctuation", False)

        output_dir = args.output_dir or cfg["output_dir"]
        if not output_dir:
            import platforminfo
            output_dir = platforminfo.default_output_dir()
        os.makedirs(output_dir, exist_ok=True)
        _cleanup_old_srts(output_dir, int(cfg.get("keep_srt_days", 30) or 0))

        timeline_name = timeline.GetName()
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in timeline_name)
        # Per-run timestamp suffix so Resolve treats each import as a fresh
        # media item -- importing an SRT with a name Resolve has seen before
        # will reuse the cached pool item and ignore the file changes.
        run_stamp = time.strftime("%Y%m%d-%H%M%S")
        srt_path = os.path.join(output_dir, f"{safe_name} {run_stamp}.srt")

        # offset_s puts the captions where the speech actually is when the
        # user rendered an in/out range instead of the whole timeline.
        rc = _transcribe_to_srt(
            wav_path, srt_path, fps, strip_punct,
            offset_s=timeline_offset_s,
            empty_message=("No speech detected. Check in/out points cover "
                           "the spoken section, and audio is unmuted."),
        )
        if rc != 0:
            return rc

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
        _ensure_terminal_status(ui_proc)


def run_file_mode(args):
    """Transcribe a file directly without Resolve."""
    # Apply the device override BEFORE importing Transcriber: it reads
    # _config.DEVICE when it creates the model.
    if getattr(args, "device", None):
        import config as _cfg_mod
        _cfg_mod.DEVICE = args.device
        _cfg_mod.COMPUTE_TYPE = "int8" if args.device == "cpu" else "float16"

    _apply_cli_overrides(args)

    input_path = os.path.abspath(args.file)
    if not os.path.exists(input_path):
        log.error(f"File not found: {input_path}")
        return 1
    log.info(f"Input: {input_path}")

    if args.output:
        srt_path = os.path.abspath(args.output)
    else:
        srt_path = os.path.splitext(input_path)[0] + ".srt"

    # Spawn UI by default; disable with --no-ui.
    ui_proc = None if getattr(args, "no_ui", False) else _spawn_progress_ui()
    _install_cancel_signal()

    # Same try/except/finally discipline as run_resolve_mode: a crash or a
    # Cancel-induced KeyboardInterrupt must write a terminal status, or the
    # progress window hangs until its 90s stale timeout.
    try:
        rc = _transcribe_to_srt(
            input_path, srt_path,
            args.fps or 24.0,                      # standalone files have no fps
            getattr(args, "strip_punctuation", False),
        )
        if rc != 0:
            return rc
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
        _ensure_terminal_status(ui_proc)


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
            from check import run_check_mode
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
