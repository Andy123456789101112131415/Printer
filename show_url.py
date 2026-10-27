# -*- coding: utf-8 -*-
"""Show the current public cpolar URL.

Usage: python show_url.py
Prerequisite: the cpolar tunnel must be running (start the startup batch file or run cpolar http 5000).
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
        print("Tunnel not found. Start the startup batch file first.")
        return 1

    m = re.search(r'window\.data = JSON\.parse\("((?:[^"\\]|\\.)*)"\)', resp)
    if not m:
        print("Tunnel information not found. Check that cpolar is running.")
        return 1

    try:
        text = json.loads('"' + m.group(1) + '"')   # Restore the JSON-escaped string.
        data = json.loads(text)
    except Exception:
        print("Failed to parse tunnel information. Please try again.")
        return 1

    tunnels = data.get("UiState", {}).get("Tunnels", [])
    if not tunnels:
        # Fall back to the request history.
        tunnels = [txn.get("ConnCtx", {}).get("Tunnel", {})
                   for txn in data.get("Txns", [])]

    urls = set()
    for tunnel in tunnels:
        url = tunnel.get("PublicUrl")
        if url:
            urls.add(url)

    if not urls:
        print("The tunnel is not ready yet. Wait a few seconds and try again.")
        return 1

    print("Current public URL(s):")
    for u in sorted(urls):
        print("  " + u)
    return 0


if __name__ == "__main__":
    sys.exit(main())
