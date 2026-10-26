# -*- coding: utf-8 -*-
"""查询 cpolar 当前公网地址。

用法: python show_url.py
前提: cpolar 隧道正在运行 (启动打印站.bat 或 cpolar http 5000)
"""

import json
import re
import sys
import urllib.request


def main():
    try:
        resp = urllib.request.urlopen(
            "http://127.0.0.1:4040/http/in", timeout=5).read().decode("utf-8", "replace")
    except Exception:
        print("查不到隧道：请先运行「启动打印站.bat」")
        return 1

    m = re.search(r'window\.data = JSON\.parse\("((?:[^"\\]|\\.)*)"\)', resp)
    if not m:
        print("未找到隧道信息，请确认 cpolar 正在运行")
        return 1

    try:
        text = json.loads('"' + m.group(1) + '"')   # 先还原被 JSON 转义的字符串
        data = json.loads(text)
    except Exception:
        print("解析失败，请稍后重试")
        return 1

    tunnels = data.get("UiState", {}).get("Tunnels", [])
    if not tunnels:
        # 兜底：从请求历史里找
        tunnels = [txn.get("ConnCtx", {}).get("Tunnel", {})
                   for txn in data.get("Txns", [])]

    urls = set()
    for tunnel in tunnels:
        url = tunnel.get("PublicUrl")
        if url:
            urls.add(url)

    if not urls:
        print("隧道尚未建立，请稍等几秒后重试")
        return 1

    print("当前公网地址：")
    for u in sorted(urls):
        print("  " + u)
    return 0


if __name__ == "__main__":
    sys.exit(main())
