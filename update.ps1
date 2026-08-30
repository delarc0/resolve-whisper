# Update an installed copy of the tool: pull the latest code, then re-run
# setup so the Resolve menu entries and dependencies match it.
#
# Pull ALONE is not enough and is actively dangerous: the presets are COPIES
# living in Resolve's Scripts folder, so a pull updates the launcher while
# leaving stale presets behind. Updating always means pull + setup.
$ErrorActionPreference = "Stop"

$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $AppDir

Write-Host ""
Write-Host "  LAB37 TOOLS: Resolve Whisper - update"
Write-Host "  ========================================"
Write-Host ""

if (-not (Test-Path (Join-Path $AppDir ".git"))) {
    Write-Host "  This install is not a git checkout (downloaded as a ZIP?), so it"
    Write-Host "  can't update itself. Re-install with:"
    Write-Host ""
    Write-Host "    git clone https://github.com/delarc0/resolve-whisper.git `"$AppDir.new`""
    Write-Host "    cd `"$AppDir.new`"; .\setup.ps1"
    Write-Host ""
    exit 1
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "  ERROR: git not found. Install it with: winget install --id Git.Git -e"
    exit 1
}

$before = (git log -1 --format='%h (%cs)')
Write-Host "  Current build: $before"

# Local edits are stashed, never discarded (see update.sh for the reasoning).
if (git status --porcelain) {
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
    Write-Host "  Local changes found; stashing them as 'pre-update $stamp'."
    Write-Host "  (Recover with: git stash list / git stash pop)"
    git stash push -u -m "pre-update $stamp" | Out-Null
}

Write-Host "  Fetching..."
git fetch --quiet origin

# A checkout that tracks nothing fails the merge below with a message about
# divergence, which sends the reader looking for the wrong problem.
git rev-parse --abbrev-ref --symbolic-full-name "@{upstream}" 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  ERROR: this copy isn't tracking a branch on GitHub, so there's"
    Write-Host "  nothing to update from. Send this to Erik:"
    Write-Host ""
    git status --short --branch
    exit 1
}

# --ff-only: refuse to invent a merge commit on a user's machine.
git merge --ff-only "@{upstream}" --quiet 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  ERROR: can't fast-forward (this checkout has diverged from origin)."
    Write-Host "  Send this to Erik:"
    Write-Host ""
    git status --short --branch
    exit 1
}

$after = (git log -1 --format='%h (%cs)')
if ($before -eq $after) {
    Write-Host "  Already up to date ($after)."
} else {
    Write-Host "  Updated: $before -> $after"
}

Write-Host ""
Write-Host "  Re-running setup (dependencies + Resolve menu entries)..."
Write-Host ""
& (Join-Path $AppDir "setup.ps1")

Write-Host "  Update complete. Build: $after"
Write-Host "  Restart DaVinci Resolve, then run Workspace > Scripts > LAB37 Check."
Write-Host ""
