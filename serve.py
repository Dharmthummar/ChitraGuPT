from __future__ import annotations

import argparse
import threading
import webbrowser

from waitress import serve

from app import app, get_lan_ip


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Serve ChitraGuPT with Waitress")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=5055, type=int)
    parser.add_argument("--open", action="store_true")
    return parser.parse_args()


def open_browser_after_start(host: str, port: int) -> None:
    browser_host = "127.0.0.1" if host in {"0.0.0.0", "::"} else host
    webbrowser.open(f"http://{browser_host}:{port}")


if __name__ == "__main__":
    args = parse_args()
    if args.open:
        threading.Timer(1.2, open_browser_after_start, args=(args.host, args.port)).start()

    print("ChitraGuPT")
    print(f"Local URL: http://127.0.0.1:{args.port}")
    if args.host == "0.0.0.0":
        print(f"Phone/LAN URL: http://{get_lan_ip()}:{args.port}")
    print("Serving with Waitress. Press Ctrl+C to stop.")
    serve(app, host=args.host, port=args.port, threads=8)
