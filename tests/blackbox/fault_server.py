from __future__ import annotations

from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import time
from urllib.parse import unquote, urlsplit


ROOT = Path(os.environ.get("DIST_DIR", "apps/demo/dist")).resolve()
PORT = int(os.environ.get("PORT", "4181"))
DELAY_FILE = os.environ.get("DELAY_FILE", "")
DELAY_MS = max(0, int(os.environ.get("DELAY_MS", "1500")))
DELAY_HITS = max(0, int(os.environ.get("DELAY_HITS", "50")))
EVENTS_PATH = Path(os.environ.get("EVENTS_PATH", "fault-events.jsonl")).resolve()
COUNTS: dict[str, int] = {}
EVENTS: list[dict[str, object]] = []


def record(event: dict[str, object]) -> None:
    value = {"time": time.time(), **event}
    EVENTS.append(value)
    EVENTS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with EVENTS_PATH.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(value, ensure_ascii=False) + "\n")


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, fmt: str, *args: object) -> None:
        record({"event": "http", "message": fmt % args})

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def guess_type(self, path: str) -> str:
        if path.endswith(".wasm"):
            return "application/wasm"
        return super().guess_type(path)

    def do_GET(self) -> None:
        path = unquote(urlsplit(self.path).path)
        if path == "/__health":
            body = b"ok"
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/__events":
            body = json.dumps(EVENTS, ensure_ascii=False).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        name = Path(path).name
        COUNTS[name] = COUNTS.get(name, 0) + 1
        if name == DELAY_FILE and COUNTS[name] <= DELAY_HITS:
            record(
                {
                    "event": "delay-start",
                    "file": name,
                    "hit": COUNTS[name],
                    "delayMs": DELAY_MS,
                }
            )
            time.sleep(DELAY_MS / 1000)
            record(
                {
                    "event": "delay-finished",
                    "file": name,
                    "hit": COUNTS[name],
                    "delayMs": DELAY_MS,
                }
            )
        super().do_GET()


ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
