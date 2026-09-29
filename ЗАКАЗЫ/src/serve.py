"""Раздаёт страницу и хранит рабочие данные в одном файле SQLite."""

import json
import sqlite3
import threading
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "data" / "workspace.sqlite"
LOCK = threading.Lock()


def connect():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB)
    con.execute(
        "CREATE TABLE IF NOT EXISTS workspace ("
        "id INTEGER PRIMARY KEY CHECK (id = 1), body TEXT NOT NULL)"
    )
    return con


def read_body():
    with LOCK:
        con = connect()
        try:
            row = con.execute("SELECT body FROM workspace WHERE id = 1").fetchone()
        finally:
            con.close()
    if not row:
        return {}
    try:
        data = json.loads(row[0])
    except json.JSONDecodeError:
        return {}
    return data if isinstance(data, dict) else {}


def keeps_catalog(current, data):
    old = len((current or {}).get("catalog") or [])
    new = len((data or {}).get("catalog") or [])
    return not (old > 0 and new == 0)


def write_body(data):
    if not isinstance(data, dict):
        return False
    with LOCK:
        current = {}
        con = connect()
        try:
            row = con.execute("SELECT body FROM workspace WHERE id = 1").fetchone()
            if row:
                try:
                    loaded = json.loads(row[0])
                    if isinstance(loaded, dict):
                        current = loaded
                except json.JSONDecodeError:
                    current = {}
            if not keeps_catalog(current, data):
                return False
            con.execute(
                "INSERT INTO workspace (id, body) VALUES (1, ?) "
                "ON CONFLICT(id) DO UPDATE SET body = excluded.body",
                (json.dumps(data, ensure_ascii=False),),
            )
            con.commit()
        finally:
            con.close()
    return True


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        if self.path.split("?", 1)[0] != "/api/workspace":
            return super().do_GET()
        raw = json.dumps(read_body(), ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self):
        if self.path.split("?", 1)[0] != "/api/workspace":
            self.send_error(404)
            return
        length = int(self.headers.get("Content-Length") or "0")
        raw = self.rfile.read(length) if length else b""
        try:
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self.send_error(400)
            return
        saved = write_body(data)
        body = json.dumps({"ok": saved}, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    server = ThreadingHTTPServer(("127.0.0.1", 8765), Handler)
    threading.Timer(0.4, lambda: webbrowser.open("http://127.0.0.1:8765/zakazy.html")).start()
    print("http://127.0.0.1:8765/zakazy.html")
    server.serve_forever()


if __name__ == "__main__":
    main()
