#!/usr/bin/env python3
"""
网络设备库存 — 同步服务端
纯标准库实现：http.server + sqlite3，零外部依赖

启动方式：
    python server.py                 # 默认 0.0.0.0:8080
    python server.py --port 9000     # 指定端口
    python server.py --token abc123  # 设置 API Token（或通过环境变量 STOCK_API_TOKEN）

部署到 Render（免费）：
    1. 创建 Web Service，选 Python
    2. Start Command: python server.py --port $PORT
    3. 添加环境变量 STOCK_API_TOKEN
"""

import argparse
import json
import os
import sqlite3
import sys
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote

# Windows 终端 UTF-8 支持
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "stock.db")
API_TOKEN = os.environ.get("STOCK_API_TOKEN", "")


# ============================================================
# 数据库操作
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS items (
            model   TEXT PRIMARY KEY,
            info    TEXT NOT NULL DEFAULT '',
            num     INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        )
    """)
    conn.commit()
    return conn


def db_get_all():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT model, info, num FROM items ORDER BY model").fetchall()
    conn.close()
    result = {}
    for r in rows:
        result[r["model"]] = {"info": r["info"], "num": r["num"]}
    return result


def db_upsert(model, info, num):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        INSERT INTO items (model, info, num, updated_at)
        VALUES (?, ?, ?, datetime('now','localtime'))
        ON CONFLICT(model) DO UPDATE SET
            info = excluded.info,
            num = excluded.num,
            updated_at = excluded.updated_at
    """, (model, info, num))
    conn.commit()
    conn.close()


def db_delete(model):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM items WHERE model = ?", (model,))
    conn.commit()
    conn.close()


def db_count():
    conn = sqlite3.connect(DB_FILE)
    count = conn.execute("SELECT COUNT(*) FROM items").fetchone()[0]
    conn.close()
    return count


# ============================================================
# HTTP Handler
# ============================================================
class APIHandler(BaseHTTPRequestHandler):

    def _check_auth(self):
        if not API_TOKEN:
            return True
        auth = self.headers.get("Authorization", "")
        return auth == f"Bearer {API_TOKEN}"

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length == 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def _set_cors(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def do_OPTIONS(self):
        self._set_cors()

    def do_GET(self):
        path = urlparse(self.path).path.rstrip("/")

        if path == "/api/health":
            self._send_json({"status": "ok", "count": db_count()})

        elif path == "/api/items":
            if not self._check_auth():
                self._send_json({"error": "unauthorized"}, 401)
                return
            self._send_json(db_get_all())

        else:
            self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        path = urlparse(self.path).path.rstrip("/")

        if path == "/api/items":
            if not self._check_auth():
                self._send_json({"error": "unauthorized"}, 401)
                return
            try:
                body = self._read_body()
                model = body.get("model", "").strip()
                if not model:
                    self._send_json({"error": "model is required"}, 400)
                    return
                info = body.get("info", "")
                num = body.get("num", 0)
                if not isinstance(num, int):
                    num = int(num)
                db_upsert(model, info, num)
                self._send_json({"ok": True, "model": model})
            except (ValueError, TypeError) as e:
                self._send_json({"error": str(e)}, 400)

        else:
            self._send_json({"error": "not found"}, 404)

    def do_DELETE(self):
        path = urlparse(self.path).path.rstrip("/")

        if path.startswith("/api/items/"):
            if not self._check_auth():
                self._send_json({"error": "unauthorized"}, 401)
                return
            model = unquote(path[len("/api/items/"):])
            db_delete(model)
            self._send_json({"ok": True, "model": model})

        else:
            self._send_json({"error": "not found"}, 404)

    def log_message(self, format, *args):
        # 简洁日志格式
        print(f"[{self.log_date_time_string()}] {args[0]}")


# ============================================================
# 主入口
# ============================================================
def main():
    global API_TOKEN

    parser = argparse.ArgumentParser(description="库存同步服务端")
    parser.add_argument("--port", type=int, default=8080, help="监听端口（默认 8080）")
    parser.add_argument("--host", default="0.0.0.0", help="监听地址（默认 0.0.0.0）")
    parser.add_argument("--token", default="", help="API 认证 Token（也可通过环境变量 STOCK_API_TOKEN 设置）")
    args = parser.parse_args()

    if args.token:
        API_TOKEN = args.token

    # 初始化数据库
    init_db()
    count = db_count()
    print(f"📦 库存同步服务端启动")
    print(f"   地址: http://{args.host}:{args.port}")
    print(f"   数据库: {DB_FILE} ({count} 条记录)")
    print(f"   Token: {'已设置' if API_TOKEN else '(未设置，允许无认证访问)'}")

    server = HTTPServer((args.host, args.port), APIHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n⏹  服务已停止")
        server.shutdown()


if __name__ == "__main__":
    main()
