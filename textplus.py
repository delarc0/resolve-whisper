"""
Insert Whisper captions onto a Resolve timeline as native Text+ clips on a
new video track. Each caption gets its own Fusion comp with a styled
TextPlus node wired to MediaOut.

PLATFORM NOTE: this works on Windows Resolve. On Mac Resolve 20, the
TimelineItem.AddFusionComp() API silently returns None on PNG clips, which
makes the Text+ insertion fail. caption.py's run_resolve_mode falls back to
plain SRT subtitle import in that case. Don't expect this to "just work" on
Mac without rethinking the clip-creation step (likely needs a different
clip type that ships with a Fusion comp by default).

Ported from resolve_script.py's Windows-tested implementation.
"""
import logging
import os
import struct
import time
import zlib

log = logging.getLogger(__name__)

_TEMPLATE_NAME = "_WhisperTemplate"
_MEDIA_TYPE_VIDEO = 1


def _create_transparent_png(path: str, width: int = 1920, height: int = 1080) -> None:
    """Write an all-transparent RGBA PNG. No PIL needed."""
    header = b"\x89PNG\r\n\x1a\n"

    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    ihdr_crc = zlib.crc32(b"IHDR" + ihdr_data) & 0xFFFFFFFF
    ihdr = struct.pack(">I", 13) + b"IHDR" + ihdr_data + struct.pack(">I", ihdr_crc)

    row = b"\x00" + b"\x00" * (width * 4)
    raw = row * height
    compressed = zlib.compress(raw)
    idat_crc = zlib.crc32(b"IDAT" + compressed) & 0xFFFFFFFF
    idat = struct.pack(">I", len(compressed)) + b"IDAT" + compressed + struct.pack(">I", idat_crc)

    iend_crc = zlib.crc32(b"IEND") & 0xFFFFFFFF
    iend = struct.pack(">I", 0) + b"IEND" + struct.pack(">I", iend_crc)

    with open(path, "wb") as f:
        f.write(header + ihdr + idat + iend)


def _find_clip_in_pool(folder, name):
    try:
        for clip in (folder.GetClipList() or []):
            try:
                if clip.GetClipProperty("Clip Name") == name:
                    return clip
            except Exception:
                pass
        for sub in (folder.GetSubFolderList() or []):
            r = _find_clip_in_pool(sub, name)
            if r:
                return r
    except Exception as e:
        log.warning(f"_find_clip_in_pool: {e}")
    return None


def _get_or_create_template(media_pool, app_dir: str, width: int, height: int):
    root = media_pool.GetRootFolder()
    if not root:
        log.error("Media pool root folder unavailable")
        return None

    png_path = os.path.join(app_dir, f"{_TEMPLATE_NAME}.png")
    _create_transparent_png(png_path, width, height)

    # Recreate so the cached resolution updates if the timeline changed size
    existing = _find_clip_in_pool(root, _TEMPLATE_NAME)
    if existing:
        try:
            media_pool.DeleteClips([existing])
        except Exception as e:
            log.warning(f"DeleteClips(template): {e}")

    prev_folder = media_pool.GetCurrentFolder()
    media_pool.SetCurrentFolder(root)
    try:
        imported = media_pool.ImportMedia([png_path])
        if not imported:
            log.error("ImportMedia returned empty for template")
            return None
        try:
            imported[0].SetClipProperty("Clip Name", _TEMPLATE_NAME)
        except Exception:
            pass
        return imported[0]
    finally:
        if prev_folder:
            media_pool.SetCurrentFolder(prev_folder)


def _style_textplus(tp, text: str, video_height: int) -> None:
    """Apply Reels-friendly styling: big bold white type with black outline."""
    tp.SetInput("StyledText", text)
    tp.SetInput("Font", "Open Sans")
    tp.SetInput("Style", "Bold")
    # Size is normalized to vertical resolution. 0.07 ≈ ~134px on 1920 vertical
    tp.SetInput("Size", 0.07)
    # Center horizontally, low-third vertical (good for 9x16 Reels)
    tp.SetInput("Center", {1: 0.5, 2: 0.18})
    # White fill
    tp.SetInput("Red1", 1.0)
    tp.SetInput("Green1", 1.0)
    tp.SetInput("Blue1", 1.0)
    # Black outline
    tp.SetInput("Enabled2", 1)
    tp.SetInput("Red2", 0.0)
    tp.SetInput("Green2", 0.0)
    tp.SetInput("Blue2", 0.0)
    tp.SetInput("Thickness2", 0.2)


def insert_textplus_captions(project, timeline, captions, fps: float, app_dir: str) -> bool:
    """Top-level entry point. Returns True if any clip got Text+."""
    media_pool = project.GetMediaPool()
    if not media_pool:
        log.error("Could not access media pool")
        return False

    tl_w = int(timeline.GetSetting("timelineResolutionWidth") or 1920)
    tl_h = int(timeline.GetSetting("timelineResolutionHeight") or 1080)

    template = _get_or_create_template(media_pool, app_dir, tl_w, tl_h)
    if not template:
        return False

    prev_track_count = timeline.GetTrackCount("video")
    try:
        timeline.AddTrack("video")
    except Exception as e:
        log.warning(f"AddTrack(video): {e}")

    track_index = timeline.GetTrackCount("video")
    if track_index <= prev_track_count:
        log.error("Failed to add new caption video track")
        return False

    tl_start = timeline.GetStartFrame() or 0
    clip_infos = []
    for cap in captions:
        start_frame = round(cap["start"] * fps) + tl_start
        duration_frames = max(1, round((cap["end"] - cap["start"]) * fps))
        clip_infos.append({
            "mediaPoolItem": template,
            "startFrame": 0,
            "endFrame": duration_frames,
            "trackIndex": track_index,
            "recordFrame": start_frame,
            "mediaType": _MEDIA_TYPE_VIDEO,
        })

    items = None
    try:
        items = media_pool.AppendToTimeline(clip_infos)
    except Exception as e:
        log.warning(f"Batch AppendToTimeline: {e}")

    if not items:
        log.info("Falling back to one-by-one clip insertion")
        items = []
        for ci in clip_infos:
            try:
                r = media_pool.AppendToTimeline([ci])
                if r:
                    items.extend(r)
            except Exception as e:
                log.warning(f"Single AppendToTimeline: {e}")

    if not items:
        log.error("No clips inserted")
        return False

    # Resolve needs a moment to settle the batch insert before Fusion access
    time.sleep(0.5)

    fresh = timeline.GetItemListInTrack("video", track_index)
    if fresh:
        items = fresh

    success = 0
    for i, item in enumerate(items):
        if i >= len(captions):
            break
        text = captions[i]["text"]
        try:
            comp = item.GetFusionCompByIndex(1) or item.AddFusionComp()
            if not comp:
                log.warning(f"Caption {i+1}: no Fusion comp available")
                continue

            existing = comp.GetToolList(False, "TextPlus")
            if existing:
                for tool in existing.values():
                    tool.SetInput("StyledText", text)
                    success += 1
                    break
                continue

            comp.Lock()
            try:
                tp = comp.AddTool("TextPlus", -32768, -32768)
                if not tp:
                    log.warning(f"Caption {i+1}: AddTool returned None")
                    continue
                _style_textplus(tp, text, tl_h)
                media_out = comp.FindTool("MediaOut1")
                if media_out:
                    media_out.FindMainInput(1).ConnectTo(tp.FindMainOutput(1))
                    success += 1
                else:
                    log.warning(f"Caption {i+1}: no MediaOut1")
            finally:
                comp.Unlock()
        except Exception as e:
            log.warning(f"Caption {i+1}: {e}")

    try:
        timeline.SetTrackName("video", track_index, "Captions")
    except Exception:
        pass

    log.info(f"Text+ applied to {success}/{len(items)} clips on track {track_index}")
    return success > 0
