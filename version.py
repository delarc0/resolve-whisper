"""Report which build of the tool is running.

Bug reports from other machines are undiagnosable without this: "captions
land at 00:00" means one thing on last month's build and another on today's.
The version is stamped into every run log and the pre-flight check so the
answer travels with the report instead of having to be asked for.

Source of truth is the git checkout (that is how the tool is installed).
A ZIP download has no .git, so that case is reported honestly as unknown
rather than guessed at.
"""
import os
import subprocess

APP_DIR = os.path.dirname(os.path.abspath(__file__))

_cache = None


def _git(*args):
    """Run a git command in the app dir, or return None if it can't."""
    try:
        out = subprocess.run(
            ("git", "-C", APP_DIR) + args,
            capture_output=True, text=True, timeout=5,
        )
    except (OSError, subprocess.SubprocessError):
        return None  # git not installed, or hung
    if out.returncode != 0:
        return None
    return out.stdout.strip()


def _compute():
    # rev-parse doubles as the "is this a git checkout" probe: a ZIP install
    # has no .git and every git call below would fail the same way.
    sha = _git("log", "-1", "--format=%h")
    if not sha:
        stamp = ""
        try:
            import datetime
            mtime = os.path.getmtime(os.path.join(APP_DIR, "caption.py"))
            stamp = datetime.date.fromtimestamp(mtime).isoformat()
        except OSError:
            pass
        return f"unknown (no git checkout{', files from ' + stamp if stamp else ''})"

    date = _git("log", "-1", "--format=%cs") or "?"
    # Uncommitted edits mean the sha alone is a lie about what is running.
    dirty = _git("status", "--porcelain") or ""
    suffix = " +local-changes" if dirty.strip() else ""
    return f"{sha} ({date}){suffix}"


def version_string() -> str:
    """Short human-readable build id, e.g. 'eac3c49 (2026-08-10)'.

    Cached: this runs on every launch and shells out to git.
    """
    global _cache
    if _cache is None:
        _cache = _compute()
    return _cache


def behind_by() -> int:
    """How many commits behind origin this checkout is, or -1 if unknown.

    Deliberately does NOT fetch: this is called from the pre-flight check,
    which must stay fast and work offline. It compares against the last
    known origin state, so it under-reports rather than blocking.
    """
    count = _git("rev-list", "--count", "HEAD..@{upstream}")
    if count is None or not count.isdigit():
        return -1
    return int(count)


if __name__ == "__main__":
    print(version_string())
