# -*- coding: utf-8 -*-
"""Automatically commit and push project changes after a quiet period.

Usage: python autosave.py [quiet-period-in-seconds, default 60]
Exit: Ctrl+C

Only Git-tracked or non-ignored files are monitored. Uploads, the database,
secrets, and the migration package are excluded by .gitignore.
"""

import os
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
DEBOUNCE = int(sys.argv[1]) if len(sys.argv) > 1 else 60
IGNORE_DIRS = {".git", "uploads", "__pycache__", "打印站迁移包"}


def snapshot():
    """Return modification-time signatures for project files, excluding ignored files."""
    sig = []
    for root, dirs, files in os.walk(BASE):
        dirs[:] = [d for d in dirs if d not in IGNORE_DIRS]
        for f in files:
            if f.endswith(".pyc"):
                continue
            p = Path(root) / f
            try:
                sig.append((str(p.relative_to(BASE)), p.stat().st_mtime_ns, p.stat().st_size))
            except OSError:
                pass
    return tuple(sorted(sig))


def run(cmd):
    r = subprocess.run(cmd, cwd=str(BASE), capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.returncode, ((r.stdout or "") + (r.stderr or "")).strip()


def git_save():
    code, _ = run(["git", "add", "-A"])
    if code != 0:
        return False, "git add failed"
    code, out = run(["git", "status", "--porcelain"])
    if code != 0:
        return False, "git status failed"
    if not out.strip():
        return True, "No changes"
    # Rebase from the remote to avoid diverging branches; continue if the pull fails.
    run(["git", "pull", "--rebase", "origin", "main"])
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    code, _ = run(["git", "commit", "-m", f"Autosave {ts}"])
    if code != 0:
        return False, "Commit failed (possibly because there were no changes)"
    code, _ = run(["git", "push", "origin", "main"])
    if code != 0:
        return False, "Push failed (check your network connection and SSH access)"
    return True, f"Committed and pushed ({ts})"


def main():
    print(f"[autosave] Watching directory: {BASE}")
    print(f"[autosave] Quiet period: {DEBOUNCE} seconds; changes are committed and pushed after this period")
    print("[autosave] Press Ctrl+C to exit")
    last = snapshot()
    pending_since = None
    while True:
        time.sleep(5)
        cur = snapshot()
        if cur != last:
            last = cur
            pending_since = time.time()
            print(f"[autosave] Changes detected; committing in {DEBOUNCE} seconds ...")
        elif pending_since and time.time() - pending_since >= DEBOUNCE:
            pending_since = None
            ok, msg = git_save()
            print(f"[autosave] {'OK: ' if ok else 'FAIL: '}{msg}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[autosave] Exited")
