"""Report which build of the tool is running.

Bug reports from other machines are undiagnosable without this: "captions
land at 00:00" means one thing on last month's build and another on today's.
The version is stamped into every run log and the pre-flight check so the
answer travels with the report instead of having to be asked for.

Source of truth is the git checkout (that is how the tool is installed).
A ZIP download has no .git, so that case is reported honestly as unknown
rather than guessed at.
"""
import json
import os
import subprocess
import time

APP_DIR = os.path.dirname(os.path.abspath(__file__))

_cache = None


def _git(*args, timeout=5):
    """Run a git command in the app dir, or return None if it can't."""
    env = dict(os.environ)
    # Never let git stop for a credential prompt: this runs inside a detached
    # subprocess with no one to type into it, and it would hang until the
    # timeout instead of failing fast.
    env["GIT_TERMINAL_PROMPT"] = "0"
    try:
        out = subprocess.run(
            ("git", "-C", APP_DIR) + args,
            capture_output=True, text=True, timeout=timeout,
            stdin=subprocess.DEVNULL, env=env,
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


def behind_by(fetch: bool = False) -> int:
    """How many commits behind origin this checkout is, or -1 if unknown.

    `HEAD..@{upstream}` compares against the remote-tracking ref, which only
    moves on fetch. Without one this answers 0 on exactly the machine the
    check exists for: someone who cloned weeks ago and never fetched is
    behind, and their tracking ref still says they are not. So callers that
    want a real answer pass fetch=True.

    Offline stays a silent 0/-1 rather than an error: a health check must
    still pass on a machine with no network.
    """
    if fetch:
        _git("fetch", "--quiet", "origin", timeout=15)
    count = _git("rev-list", "--count", "HEAD..@{upstream}")
    if count is None or not count.isdigit():
        return -1
    return int(count)


def update_available(max_age_h: float = 24.0) -> int:
    """Commits behind origin, answered from a cache at most max_age_h old.

    Called on every caption run, so it must not cost a network round trip
    every time and must never delay or break a run: any failure answers 0.
    The cache lives outside the checkout (see platforminfo.cache_dir) so it
    cannot make the build look locally modified.
    """
    try:
        import platforminfo
        cache = platforminfo.cache_dir()
        path = os.path.join(cache, "update_check.json") if cache else ""
        now = time.time()

        if path and os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                if (now - float(data.get("checked", 0))) < max_age_h * 3600:
                    return max(int(data.get("behind", 0)), 0)
            except (OSError, ValueError, TypeError):
                pass  # unreadable cache just means we re-check

        behind = behind_by(fetch=True)
        behind = behind if behind > 0 else 0
        if path:
            tmp = f"{path}.{os.getpid()}.tmp"
            try:
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump({"checked": now, "behind": behind}, f)
                os.replace(tmp, path)
            except OSError:
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
        return behind
    except Exception:
        return 0  # an update check must never be able to break a caption run


if __name__ == "__main__":
    print(version_string())
