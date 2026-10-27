# -*- coding: utf-8 -*-
"""Watchdog that automatically restarts app.py if it exits unexpectedly."""

import subprocess
import sys
import time

while True:
    print("[run] starting app.py ...")
    proc = subprocess.run([sys.executable, "app.py"])
    print(f"[run] app.py exited with code {proc.returncode}, restart in 3s ...")
    time.sleep(3)
