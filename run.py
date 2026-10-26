# -*- coding: utf-8 -*-
"""服务器看门狗：app.py 意外退出后自动重启，保证服务不中断。"""

import subprocess
import sys
import time

while True:
    print("[run] starting app.py ...")
    proc = subprocess.run([sys.executable, "app.py"])
    print(f"[run] app.py exited with code {proc.returncode}, restart in 3s ...")
    time.sleep(3)
