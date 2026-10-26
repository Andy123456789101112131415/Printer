# -*- coding: utf-8 -*-
"""自动保存与推送：监控项目文件变化，静默期过后自动 commit + push。

用法: python autosave.py [静默秒数，默认 60]
退出: Ctrl+C

只监控 git 会跟踪/未忽略的文件；uploads、data.db、密钥、迁移包等已被 .gitignore 排除。
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
    """返回项目文件（排除 .gitignore 项）的修改时间签名。"""
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
        return False, "git add 失败"
    code, out = run(["git", "status", "--porcelain"])
    if code != 0:
        return False, "git status 失败"
    if not out.strip():
        return True, "无变更"
    # 拉取远端（--rebase 避免分叉），失败不阻塞提交
    run(["git", "pull", "--rebase", "origin", "main"])
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    code, _ = run(["git", "commit", "-m", f"自动保存 {ts}"])
    if code != 0:
        return False, "commit 失败（可能无变更）"
    code, _ = run(["git", "push", "origin", "main"])
    if code != 0:
        return False, "push 失败（检查网络/SSH）"
    return True, f"已提交并推送 ({ts})"


def main():
    print(f"[autosave] 监控目录: {BASE}")
    print(f"[autosave] 静默期: {DEBOUNCE} 秒，修改后 {DEBOUNCE} 秒内无新改动则自动提交推送")
    print("[autosave] Ctrl+C 退出")
    last = snapshot()
    pending_since = None
    while True:
        time.sleep(5)
        cur = snapshot()
        if cur != last:
            last = cur
            pending_since = time.time()
            print(f"[autosave] 检测到修改，{DEBOUNCE} 秒后提交 ...")
        elif pending_since and time.time() - pending_since >= DEBOUNCE:
            pending_since = None
            ok, msg = git_save()
            print(f"[autosave] {'OK: ' if ok else 'FAIL: '}{msg}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[autosave] 已退出")
