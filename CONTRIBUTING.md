# Contributing

Thanks for looking. This is a small tool built for one studio's workflow and
then opened up, so the bar here is "does it keep working on other people's
machines", not "does it cover every case".

## Before you file an issue

Run the health check and paste its output. It catches almost every install
problem on its own, and it prints the build id, which tells me which code
you are running:

```
./.venv/bin/python caption.py --check     # Mac
.\.venv\Scripts\python.exe caption.py --check   # Windows
```

The run log is `/tmp/resolve_whisper.log` (Mac) or
`%TEMP%\resolve_whisper.log` (Windows). Its first line carries the same
build id.

Please include your DaVinci Resolve version. Resolve's scripting API changes
between releases, and most reports that turn out to be real bugs are a
version difference.

## Running the tests

```
./.venv/bin/python -W error::ResourceWarning -m unittest discover -s tests
```

No Resolve, no Whisper model and no display required: the suite covers the
pure logic (SRT shaping, caption chunking, config validation, the
Deliver-state restore paths, the Lua preset contract). A few tests also
drive Resolve's own `fuscript` when it is installed, and skip when it isn't.

`-W error::ResourceWarning` is not optional in CI. A leaked file handle or
subprocess pipe is invisible in a passing run and is what exhausts handles
during a long session.

Lint, same rules as CI:

```
python -m flake8 --select=F *.py tests/*.py
```

F rules only, on purpose. There is no formatter, and please don't add one:
most of the odd-looking code in here is a documented workaround for a
specific Resolve version, and a reformatting pass tends to separate those
lines from the comments that explain them.

## Changing the Resolve-facing code

`resolve_api.py` owns everything that talks to Resolve. Two rules live there
and are enforced by tests:

- **The Deliver page is borrowed, not owned.** Rendering mutates the render
  preset, the format and codec, the output path, the render queue and the
  page the user is looking at. Every one of those is snapshotted and restored
  in a `finally`, including on the failure paths. Only jobs the tool added
  are ever deleted.
- **Resolve 18 through 21 all have to work.** The compat shims handle the
  21.x job-key renames and the factory "Audio Only" preset that writes AAC
  rather than WAV. Don't delete a shim because the version you have doesn't
  need it.

`CLAUDE.md` in this repo lists the rest of the invariants, each of which was
a real bug once. It is written for AI assistants working on the code, and it
is the fastest way for a human to get oriented too.

## Pull requests

Small and focused is easier to take than large and sweeping. Say which
Resolve version and OS you tested on; I cannot test every combination.

Answers are best-effort. This is maintained because it is in daily use at
LAB37, not on a support commitment.
