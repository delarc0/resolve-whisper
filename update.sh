#!/bin/bash
# Update an installed copy of the tool: pull the latest code, then re-run
# setup so the Resolve menu entries and dependencies match it.
#
# Pull ALONE is not enough and is actively dangerous: the presets are COPIES
# living in Resolve's Scripts folder, so a pull updates the launcher while
# leaving stale presets behind. Updating always means pull + setup, which is
# exactly why this is one script and not two commands in a README.
set -e

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$APP_DIR"

echo ""
echo "  LAB37 TOOLS: Resolve Whisper - update"
echo "  ========================================"
echo ""

if [ ! -d "$APP_DIR/.git" ]; then
    echo "  This install is not a git checkout (downloaded as a ZIP?), so it"
    echo "  can't update itself. Re-install with:"
    echo ""
    echo "    git clone https://github.com/delarc0/resolve-whisper.git \"$APP_DIR.new\""
    echo "    cd \"$APP_DIR.new\" && ./setup.sh"
    echo ""
    exit 1
fi

if ! command -v git &>/dev/null; then
    echo "  ERROR: git not found. Install it with: xcode-select --install"
    exit 1
fi

# Refuse to update underneath a running caption job: setup rewrites the
# menu presets and reinstalls dependencies, and the job is holding a kernel
# lock we can test for by trying to take it ourselves.
if [ -x "$APP_DIR/.venv/bin/python3" ]; then
    if ! "$APP_DIR/.venv/bin/python3" - <<'PYLOCK'
import fcntl, os, sys, tempfile
lock = os.path.join(tempfile.gettempdir(), "resolve_whisper.lock")
try:
    fd = os.open(lock, os.O_CREAT | os.O_RDWR, 0o644)
    fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
except OSError:
    sys.exit(1)          # someone is captioning right now
except Exception:
    sys.exit(0)          # can't tell -> don't block the user
sys.exit(0)
PYLOCK
    then
        echo "  A caption job is running right now."
        echo "  Let it finish (or cancel it), then run this again."
        exit 1
    fi
fi

BEFORE="$(git log -1 --format='%h (%cs)' 2>/dev/null || echo unknown)"
echo "  Current build: $BEFORE"

# Local edits are stashed, never discarded. Someone hand-patching a preset to
# work around a bug is exactly the person who runs this, and silently
# throwing their fix away would be a poor trade for an unattended update.
STASHED=""
if [ -n "$(git status --porcelain)" ]; then
    STAMP="$(date '+%Y-%m-%d %H:%M')"
    echo "  Local changes found; setting them aside as 'pre-update $STAMP'."
    if git stash push -u -m "pre-update $STAMP" >/dev/null; then
        STASHED="yes"
    else
        echo ""
        echo "  ERROR: could not set your local changes aside, so the update"
        echo "  stopped rather than touching them. Send this to Erik:"
        echo ""
        git status --short --branch | head -10
        exit 1
    fi
fi

echo "  Fetching..."
git fetch --quiet origin

# A checkout that tracks nothing fails the merge below with a message about
# divergence, which sends the reader looking for the wrong problem.
if ! git rev-parse --abbrev-ref --symbolic-full-name "@{upstream}" >/dev/null 2>&1; then
    echo ""
    echo "  ERROR: this copy isn't tracking a branch on GitHub, so there's"
    echo "  nothing to update from. Send this to Erik:"
    echo ""
    git status --short --branch | head -5
    exit 1
fi

# --ff-only: refuse to invent a merge commit on a user's machine. If the
# checkout has diverged, that needs a human, not a script guessing.
if ! git merge --ff-only "@{upstream}" --quiet 2>/dev/null; then
    echo ""
    echo "  ERROR: can't fast-forward (this checkout has diverged from origin)."
    echo "  Send this to Erik:"
    echo ""
    git status --short --branch | head -20
    exit 1
fi

AFTER="$(git log -1 --format='%h (%cs)')"
if [ "$BEFORE" = "$AFTER" ]; then
    echo "  Already up to date ($AFTER)."
else
    echo "  Updated: $BEFORE -> $AFTER"
fi

echo ""
echo "  Re-running setup (dependencies + Resolve menu entries)..."
echo ""
"$APP_DIR/setup.sh"

echo "  Update complete. Build: $AFTER"
echo "  Restart DaVinci Resolve, then run Workspace > Scripts > LAB37 Check."
if [ -n "$STASHED" ]; then
    # Setup prints ~40 lines, so the earlier notice has scrolled off. Anyone
    # who had work in progress needs to see this at the point they stop
    # reading, not before a wall of pip output.
    echo ""
    echo "  ----------------------------------------"
    echo "   NOTE: you had local changes. They were set aside, not deleted."
    echo "   Get them back with:  git stash pop"
    echo "  ----------------------------------------"
fi
echo ""
