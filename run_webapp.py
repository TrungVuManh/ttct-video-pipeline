# -*- coding: utf-8 -*-
"""
run_webapp.py — Khoi dong bang dieu khien web local cho TTCT Video Pipeline.

Dung:
  python run_webapp.py                  # mo http://127.0.0.1:8787 va tu mo trinh duyet
  python run_webapp.py --port 9000       # doi cong
  python run_webapp.py --no-browser      # khong tu mo trinh duyet

Chi chay tren may ban (khong deploy len GitHub Pages duoc, vi Pages chi phuc vu
file tinh - xem README muc "Chay giao dien web (local)").
"""
import argparse
import threading
import webbrowser

import uvicorn


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8787)
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()

    url = f"http://{args.host}:{args.port}/"
    if not args.no_browser:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    print(f"TTCT Video Pipeline - bang dieu khien web: {url}")
    uvicorn.run("webapp.app:app", host=args.host, port=args.port, reload=False)


if __name__ == "__main__":
    main()
