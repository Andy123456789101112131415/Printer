# -*- coding: utf-8 -*-
"""Print Station launcher: installs missing dependencies, then runs app.py and
restarts it if it exits unexpectedly."""

import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
REQUIRED = (("flask", "flask"), ("waitress", "waitress"), ("win32print", "pywin32"))


def missing_dependencies():
    missing = []
    for module, package in REQUIRED:
        try:
            __import__(module)
        except ImportError:
            missing.append(package)
    return missing


def install_dependencies():
    missing = missing_dependencies()
    if not missing:
        return True
    print(f"[run] Missing packages: {', '.join(missing)}")
    print("[run] Installing dependencies, this may take a few minutes ...")
    code = subprocess.call(
        [sys.executable, "-m", "pip", "install", "-r", str(BASE / "requirements.txt")],
        cwd=str(BASE),
    )
    if code != 0:
        print("[run] Installation failed. Check the internet connection and start again.")
        return False
    remaining = missing_dependencies()
    if remaining:
        print(f"[run] Still missing: {', '.join(remaining)}")
        print("[run] Install them manually with: pip install -r requirements.txt")
        return False
    print("[run] Dependencies installed")
    return True


def main():
    if not install_dependencies():
        return 1
    while True:
        print("[run] starting app.py ...")
        code = subprocess.call([sys.executable, str(BASE / "app.py")], cwd=str(BASE))
        print(f"[run] app.py exited with code {code}, restart in 3s ...")
        time.sleep(3)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(0)

