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

BEFORE="$(git log -1 --format='%h (%cs)' 2>/dev/null || echo unknown)"
echo "  Current build: $BEFORE"

# Local edits are stashed, never discarded. Someone hand-patching a preset to
# work around a bug is exactly the person who runs this, and silently
# throwing their fix away would be a poor trade for an unattended update.
if [ -n "$(git status --porcelain)" ]; then
    STAMP="$(date '+%Y-%m-%d %H:%M')"
    echo "  Local changes found; stashing them as 'pre-update $STAMP'."
    echo "  (Recover with: git stash list / git stash pop)"
    git stash push -u -m "pre-update $STAMP" >/dev/null
fi

echo "  Fetching..."
git fetch --quiet origin

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
echo ""
