"""
库存同步服务端 — Flask 版本（适配 PythonAnywhere 等 WSGI 平台）

部署到 PythonAnywhere 免费版：
    1. 上传此文件到 PythonAnywhere
    2. 在 Web 面板中配置 WSGI 指向 flask_app.app
    3. 设置环境变量 STOCK_API_TOKEN（可选）
"""

import os
import sqlite3
from flask import Flask, request, jsonify

app = Flask(__name__)

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "stock.db")
API_TOKEN = os.environ.get("STOCK_API_TOKEN", "")


# ============================================================
# 数据库
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
    conn.close()


def db_get_all():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT model, info, num FROM items ORDER BY model").fetchall()
    conn.close()
    return {r["model"]: {"info": r["info"], "num": r["num"]} for r in rows}


def db_upsert(model, info, num):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        INSERT INTO items (model, info, num, updated_at)
        VALUES (?, ?, ?, datetime('now','localtime'))
        ON CONFLICT(model) DO UPDATE SET
            info = excluded.info, num = excluded.num, updated_at = excluded.updated_at
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
# 认证
# ============================================================
def check_auth():
    if not API_TOKEN:
        return True
    auth = request.headers.get("Authorization", "")
    return auth == f"Bearer {API_TOKEN}"


# ============================================================
# API 路由
# ============================================================
@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "count": db_count()})


@app.route("/api/items", methods=["GET"])
def get_items():
    if not check_auth():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(db_get_all())


@app.route("/api/items", methods=["POST"])
def upsert_item():
    if not check_auth():
        return jsonify({"error": "unauthorized"}), 401
    body = request.get_json(silent=True) or {}
    model = body.get("model", "").strip()
    if not model:
        return jsonify({"error": "model is required"}), 400
    info = body.get("info", "")
    try:
        num = int(body.get("num", 0))
    except (ValueError, TypeError):
        return jsonify({"error": "num must be an integer"}), 400
    db_upsert(model, info, num)
    return jsonify({"ok": True, "model": model})


@app.route("/api/items/<path:model>", methods=["DELETE"])
def delete_item(model):
    if not check_auth():
        return jsonify({"error": "unauthorized"}), 401
    db_delete(model)
    return jsonify({"ok": True, "model": model})


# 启动时初始化数据库
init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
