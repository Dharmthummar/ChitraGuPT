from __future__ import annotations

import argparse
import atexit
import os
import platform
import re
import secrets
import shutil
import subprocess
import threading
import urllib.parse
import urllib.request
import webbrowser

from waitress import serve

from app import DATA_DIR, app, get_lan_ip, set_public_share, set_public_share_error


PUBLIC_SHARE_FILE = DATA_DIR / "public_share.json"
TRYCLOUDFLARE_URL_RE = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Serve ChitraGuPT with Waitress")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=5055, type=int)
    parser.add_argument("--open", action="store_true")
    parser.add_argument("--public", action="store_true", help="Create a temporary public trycloudflare.com share link")
    return parser.parse_args()


def open_browser_after_start(host: str, port: int) -> None:
    browser_host = "127.0.0.1" if host in {"0.0.0.0", "::"} else host
    webbrowser.open(f"http://{browser_host}:{port}")


def phone_share_url(base_url: str, token: str) -> str:
    parts = urllib.parse.urlsplit(base_url.rstrip("/"))
    query = urllib.parse.urlencode({"phone": "1", "share": token})
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path or "/", query, parts.fragment))


def cloudflared_download_url() -> str:
    if os.name != "nt":
        raise RuntimeError("Install cloudflared first, then rerun with --public.")

    machine = platform.machine().lower()
    if machine in {"x86", "i386", "i686"}:
        return "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-386.exe"
    return "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"


def ensure_cloudflared() -> str:
    existing = shutil.which("cloudflared")
    if existing:
        return existing

    tools_dir = DATA_DIR / "tools"
    local_exe = tools_dir / "cloudflared.exe"
    if local_exe.exists():
        return str(local_exe)

    tools_dir.mkdir(parents=True, exist_ok=True)
    download_url = cloudflared_download_url()
    temp_exe = local_exe.with_suffix(".download")

    print("cloudflared was not found. Downloading the tunnel helper...")
    request = urllib.request.Request(download_url, headers={"User-Agent": "ChitraGuPT"})
    with urllib.request.urlopen(request, timeout=120) as response, temp_exe.open("wb") as output:
        shutil.copyfileobj(response, output)
    temp_exe.replace(local_exe)
    return str(local_exe)


class PublicTunnel:
    def __init__(self, port: int, token: str) -> None:
        self.port = port
        self.token = token
        self.public_url = ""
        self.ready = threading.Event()
        self.process: subprocess.Popen[str] | None = None

    def start(self) -> None:
        executable = ensure_cloudflared()
        local_url = f"http://127.0.0.1:{self.port}"
        command = [executable, "tunnel", "--url", local_url, "--no-autoupdate"]
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

        self.process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )
        atexit.register(self.stop)
        threading.Thread(target=self._watch_output, daemon=True).start()

        if self.ready.wait(timeout=20):
            return
        if self.process.poll() is not None:
            raise RuntimeError("cloudflared stopped before creating a public URL.")
        print("Public tunnel is still starting. Click Host, then Refresh, after the app opens.")

    def stop(self) -> None:
        if self.process and self.process.poll() is None:
            self.process.terminate()

    def _watch_output(self) -> None:
        if not self.process or not self.process.stdout:
            return

        for line in self.process.stdout:
            match = TRYCLOUDFLARE_URL_RE.search(line)
            if match and not self.public_url:
                self.public_url = match.group(0).rstrip("/")
                set_public_share(self.public_url, self.token)
                os.environ["CHITRAGUPT_PUBLIC_URL"] = self.public_url
                os.environ["CHITRAGUPT_SHARE_TOKEN"] = self.token
                share_url = phone_share_url(self.public_url, self.token)
                PUBLIC_SHARE_FILE.write_text(
                    f'{{"publicUrl": "{self.public_url}", "phoneUrl": "{share_url}"}}\n',
                    encoding="utf-8",
                )
                print(f"Public URL: {share_url}")
                self.ready.set()
                continue

            text = line.strip()
            if text and any(word in text.lower() for word in ("error", "failed", "unable")):
                print(f"[cloudflared] {text}")


if __name__ == "__main__":
    args = parse_args()
    if args.public:
        tunnel = PublicTunnel(args.port, secrets.token_urlsafe(18))
        try:
            tunnel.start()
        except Exception as error:
            set_public_share_error(str(error))
            print(f"Public tunnel unavailable: {error}")

    if args.open:
        threading.Timer(1.2, open_browser_after_start, args=(args.host, args.port)).start()

    print("ChitraGuPT")
    print(f"Local URL: http://127.0.0.1:{args.port}")
    if args.host == "0.0.0.0":
        print(f"Phone/LAN URL: http://{get_lan_ip()}:{args.port}")
    print("Serving with Waitress. Press Ctrl+C to stop.")
    serve(app, host=args.host, port=args.port, threads=8)
