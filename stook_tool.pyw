import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import io
import json
import os
import threading
import urllib.request
import urllib.error

# ============================================================
# 路径配置
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

# ============================================================
# 配色方案
# ============================================================
COLORS = {
    "bg":            "#ECF0F4",
    "card_bg":       "#FFFFFF",
    "header_bg":     "#1A2332",
    "header_fg":     "#FFFFFF",
    "primary":       "#1565C0",
    "primary_hover": "#0D47A1",
    "success":       "#2E7D32",
    "success_hover": "#1B5E20",
    "danger":        "#C62828",
    "danger_hover":  "#8E0000",
    "warning":       "#E65100",
    "warning_hover": "#BF360C",
    "text_primary":  "#1A1A2E",
    "text_secondary":"#5A6B7F",
    "border":        "#DDE4EE",
    "tree_even":     "#F7F9FC",
    "tree_odd":      "#FFFFFF",
    "tree_selected": "#D4E4F7",
    "online":        "#2E7D32",
    "offline":       "#C62828",
    "footer_bg":     "#E2E8F0",
}

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

FONT_FAMILY = "Microsoft YaHei UI"


# ============================================================
# 本地数据读写
# ============================================================
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ============================================================
# 服务器配置读写
# ============================================================
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"server_url": "", "api_token": ""}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ============================================================
# 同步管理器
# ============================================================
class SyncManager:
    """通过 HTTP REST API 与服务端同步数据（纯标准库 urllib）"""

    def __init__(self, server_url="", api_token=""):
        self.server_url = server_url.rstrip("/") if server_url else ""
        self.api_token = api_token
        self._online = False

    @property
    def configured(self):
        return bool(self.server_url)

    @property
    def is_online(self):
        return self._online

    def _request(self, method, path, body=None):
        """发送 HTTP 请求，返回 (ok, data)"""
        url = f"{self.server_url}{path}"
        data_bytes = None
        if body is not None:
            data_bytes = json.dumps(body, ensure_ascii=False).encode("utf-8")

        req = urllib.request.Request(url, data=data_bytes, method=method)
        req.add_header("Content-Type", "application/json; charset=utf-8")
        if self.api_token:
            req.add_header("Authorization", f"Bearer {self.api_token}")

        try:
            with urllib.request.urlopen(req, timeout=8) as resp:
                raw = resp.read().decode("utf-8")
                return True, json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            body_text = e.read().decode("utf-8", errors="replace")
            return False, {"error": f"HTTP {e.code}", "detail": body_text[:200]}
        except Exception as e:
            return False, {"error": str(e)}

    def check_health(self):
        """检测服务端是否可达"""
        if not self.server_url:
            self._online = False
            return False
        ok, data = self._request("GET", "/api/health")
        self._online = ok and data.get("status") == "ok"
        return self._online

    def pull_all(self):
        """拉取服务端全部数据"""
        if not self.server_url:
            return None
        ok, data = self._request("GET", "/api/items")
        if ok and isinstance(data, dict) and "error" not in data:
            self._online = True
            return data
        self._online = False
        return None

    def push_item(self, model, info, num):
        """推送单条型号到服务端"""
        if not self.server_url:
            return False
        ok, _ = self._request("POST", "/api/items",
                              {"model": model, "info": info, "num": num})
        if ok:
            self._online = True
        else:
            self._online = False
        return ok

    def delete_item(self, model):
        """从服务端删除型号"""
        if not self.server_url:
            return False
        from urllib.parse import quote
        ok, _ = self._request("DELETE", f"/api/items/{quote(model, safe='')}")
        if ok:
            self._online = True
        else:
            self._online = False
        return ok


# ============================================================
# 服务器配置对话框
# ============================================================
class ConfigDialog:
    """服务器连接配置窗口"""

    def __init__(self, parent, config, sync_mgr):
        self.result_config = None
        self.config = config
        self.sync_mgr = sync_mgr

        self.top = tk.Toplevel()
        self.top.title("服务器配置")
        self.top.geometry("480x340")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.configure(bg=COLORS["bg"])
        self._center_on_parent(parent)

        self._build_ui()
        self.top.wait_window()

    def _build_ui(self):
        # 标题
        header = tk.Frame(self.top, bg=COLORS["header_bg"], height=40)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="⚙️  服务器配置",
                 font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w", padx=16, pady=8)

        # 表单卡片
        card = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="both", expand=True, padx=12, pady=(10, 0))
        inner = tk.Frame(card, bg=COLORS["card_bg"])
        inner.pack(fill="both", expand=True, padx=16, pady=14)

        # 服务器地址
        tk.Label(inner, text="服务器地址", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(anchor="w")
        tk.Label(inner, text="例如: http://192.168.1.100:8080 或 https://your-app.onrender.com",
                 font=(FONT_FAMILY, 9), bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 4))
        self.url_var = tk.StringVar(value=self.config.get("server_url", ""))
        tk.Entry(inner, textvariable=self.url_var, font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(fill="x", ipady=4, pady=(0, 10))

        # API Token
        tk.Label(inner, text="API Token（可选）", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(anchor="w")
        tk.Label(inner, text="留空则不进行认证。服务端未设置 Token 时此处也留空。",
                 font=(FONT_FAMILY, 9), bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 4))
        self.token_var = tk.StringVar(value=self.config.get("api_token", ""))
        tk.Entry(inner, textvariable=self.token_var, font=(FONT_FAMILY, 10), show="•",
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(fill="x", ipady=4, pady=(0, 10))

        # 连接状态
        self.status_frame = tk.Frame(inner, bg=COLORS["card_bg"])
        self.status_frame.pack(fill="x", pady=(0, 8))
        self.status_label = tk.Label(self.status_frame, text="",
                                     font=(FONT_FAMILY, 10),
                                     bg=COLORS["card_bg"], fg=COLORS["text_secondary"])
        self.status_label.pack(side="left")

        # 提示
        tip = tk.Frame(inner, bg="#FFF8E1", bd=0, highlightthickness=1,
                       highlightbackground="#FFE082")
        tip.pack(fill="x", pady=(0, 12))
        tk.Label(tip, text="💡 服务端地址由管理员提供。本地测试可使用 python server.py 启动服务，\n"
                           "   地址填 http://localhost:8080，Token 留空即可。",
                 font=(FONT_FAMILY, 9), bg="#FFF8E1", fg="#795548",
                 justify="left").pack(padx=10, pady=8)

        # 底部按钮
        btn_bar = tk.Frame(self.top, bg=COLORS["bg"])
        btn_bar.pack(fill="x", padx=12, pady=12)

        ttk.Button(btn_bar, text="🔗 测试连接", style="Outline.TButton",
                   command=self._test_connection).pack(side="left")
        ttk.Button(btn_bar, text="保存", style="Primary.TButton",
                   command=self._save).pack(side="right", padx=(8, 0))
        ttk.Button(btn_bar, text="取消", style="Outline.TButton",
                   command=self._cancel).pack(side="right")

    def _test_connection(self):
        url = self.url_var.get().strip()
        token = self.token_var.get().strip()
        if not url:
            self.status_label.config(text="⚠️ 请先输入服务器地址", fg="#E65100")
            return

        self.status_label.config(text="⏳ 正在连接…", fg=COLORS["text_secondary"])
        self.top.update()

        # 临时 sync 测试
        tmp = SyncManager(url, token)
        if tmp.check_health():
            self.status_label.config(text="🟢 连接成功！服务端正常运行", fg=COLORS["online"])
        else:
            self.status_label.config(text="🔴 无法连接到服务器，请检查地址", fg=COLORS["offline"])

    def _save(self):
        self.result_config = {
            "server_url": self.url_var.get().strip(),
            "api_token": self.token_var.get().strip(),
        }
        save_config(self.result_config)
        self.top.destroy()

    def _cancel(self):
        self.top.destroy()

    def _center_on_parent(self, parent):
        self.top.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w, h = 480, 340
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.top.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")


# ============================================================
# 导入预览对话框
# ============================================================
class ImportDialog:
    """列映射与数据预览对话框 —— 用户确认后返回导入数据"""

    MODEL_KEYWORDS = ["型号", "设备型号", "设备", "名称", "model", "name", "产品型号", "物料"]
    INFO_KEYWORDS  = ["详情", "产品详情", "描述", "项目", "项目名称", "info", "description", "备注", "说明", "规格"]
    NUM_KEYWORDS   = ["数量", "库存数量", "库存", "当前库存", "num", "quantity", "qty", "台", "个", "件"]

    def __init__(self, parent, headers, raw_rows, existing_models):
        self.result = None
        self.headers = list(headers)
        self.raw_rows = raw_rows
        self.existing_models = existing_models

        self.model_col = self._guess(self.MODEL_KEYWORDS)
        self.info_col  = self._guess(self.INFO_KEYWORDS)
        self.num_col   = self._guess(self.NUM_KEYWORDS)
        if len(self.headers) >= 2 and self.info_col == self.model_col:
            self.info_col = 1 if self.model_col == 0 else 0
        if len(self.headers) >= 3 and (self.num_col == self.model_col or self.num_col == self.info_col):
            for i in range(len(self.headers)):
                if i not in (self.model_col, self.info_col):
                    self.num_col = i
                    break

        self._build_ui()
        self._refresh_preview()

    def _guess(self, keywords):
        for kw in keywords:
            for i, h in enumerate(self.headers):
                if kw.lower() in str(h).lower().strip():
                    return i
        return 0

    def _build_ui(self):
        self.top = tk.Toplevel()
        self.top.title("导入预览")
        self.top.geometry("780x560")
        self.top.minsize(640, 420)
        self.top.transient(self.top.master)
        self.top.grab_set()
        self.top.configure(bg=COLORS["bg"])
        self._center_on_parent()

        header = tk.Frame(self.top, bg=COLORS["header_bg"], height=42)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="📥  导入预览 — 确认列映射与数据",
                 font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w", padx=16, pady=9)

        self._build_mapping_card()
        self._build_preview_table()
        self._build_bottom_bar()

    def _build_mapping_card(self):
        card = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="x", padx=12, pady=(10, 0))
        inner = tk.Frame(card, bg=COLORS["card_bg"])
        inner.pack(fill="x", padx=14, pady=12)

        tk.Label(inner, text="列映射", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(anchor="w", pady=(0, 8))

        r1 = tk.Frame(inner, bg=COLORS["card_bg"])
        r1.pack(fill="x", pady=(0, 6))

        tk.Label(r1, text="设备型号 →", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self.model_combo = ttk.Combobox(r1, values=self.headers, state="readonly", width=22)
        self.model_combo.current(self.model_col)
        self.model_combo.pack(side="left", padx=(0, 14))
        self.model_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_preview())

        tk.Label(r1, text="产品详情 →", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self.info_combo = ttk.Combobox(r1, values=self.headers, state="readonly", width=28)
        self.info_combo.current(self.info_col)
        self.info_combo.pack(side="left", padx=(0, 14))
        self.info_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_preview())

        tk.Label(r1, text="库存数量 →", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self.num_combo = ttk.Combobox(r1, values=self.headers, state="readonly", width=14)
        self.num_combo.current(self.num_col)
        self.num_combo.pack(side="left", padx=(0, 10))
        self.num_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_preview())

        tk.Label(r1, text="默认数量:", font=(FONT_FAMILY, 9),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self.default_num = tk.StringVar(value="0")
        tk.Entry(r1, textvariable=self.default_num, width=6,
                 font=(FONT_FAMILY, 10), justify="center",
                 bg=COLORS["card_bg"], relief="solid", bd=1,
                 highlightthickness=1,
                 highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(side="left")

        r2 = tk.Frame(inner, bg=COLORS["card_bg"])
        r2.pack(fill="x")
        tk.Label(r2, text="重复型号:", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 8))
        self.dup_strategy = tk.StringVar(value="skip")
        ttk.Radiobutton(r2, text="跳过（保留已有数据）", variable=self.dup_strategy,
                        value="skip").pack(side="left", padx=(0, 14))
        ttk.Radiobutton(r2, text="覆盖（用新数据更新）", variable=self.dup_strategy,
                        value="overwrite").pack(side="left")

    def _build_preview_table(self):
        card = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="both", expand=True, padx=12, pady=(8, 0))

        title_bar = tk.Frame(card, bg=COLORS["card_bg"])
        title_bar.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(title_bar, text="数据预览", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(side="left")
        self.preview_count = tk.Label(title_bar, text="", font=(FONT_FAMILY, 9),
                                      bg=COLORS["card_bg"], fg=COLORS["text_secondary"])
        self.preview_count.pack(side="right")

        container = tk.Frame(card, bg=COLORS["card_bg"])
        container.pack(fill="both", expand=True, padx=(14, 4), pady=(0, 10))

        self.preview_tree = ttk.Treeview(
            container,
            columns=("row", "model", "info", "num"),
            show="headings", selectmode="none", height=10
        )
        self.preview_tree.heading("row", text="行", anchor="center")
        self.preview_tree.heading("model", text="设备型号")
        self.preview_tree.heading("info", text="产品详情")
        self.preview_tree.heading("num", text="数量", anchor="center")
        self.preview_tree.column("row", width=44, anchor="center", stretch=False)
        self.preview_tree.column("model", width=190, minwidth=100)
        self.preview_tree.column("info", width=380, minwidth=150)
        self.preview_tree.column("num", width=64, anchor="center", stretch=False)

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=vsb.set)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.preview_tree.tag_configure("empty", foreground="#B0BEC5")
        self.preview_tree.tag_configure("dup", foreground="#E65100")
        self.preview_tree.tag_configure("even", background=COLORS["tree_even"])
        self.preview_tree.tag_configure("odd", background=COLORS["tree_odd"])

    def _build_bottom_bar(self):
        bar = tk.Frame(self.top, bg=COLORS["bg"])
        bar.pack(fill="x", padx=12, pady=(8, 12))
        self.stats_label = tk.Label(bar, text="", font=(FONT_FAMILY, 10),
                                    bg=COLORS["bg"], fg=COLORS["text_secondary"])
        self.stats_label.pack(side="left", pady=6)
        ttk.Button(bar, text="确认导入", style="Primary.TButton",
                   command=self._confirm).pack(side="right", padx=(8, 0))
        ttk.Button(bar, text="取消", style="Outline.TButton",
                   command=self._cancel).pack(side="right")

    def _refresh_preview(self):
        for row in self.preview_tree.get_children():
            self.preview_tree.delete(row)
        mi = self.model_combo.current()
        ii = self.info_combo.current()
        ni = self.num_combo.current()
        valid = empty = dup = over = 0
        limit = min(len(self.raw_rows), 100)
        for idx, row in enumerate(self.raw_rows[:limit]):
            model = str(self._safe_get(row, mi)).strip()
            info  = str(self._safe_get(row, ii)).strip()
            num_raw = str(self._safe_get(row, ni)).strip()
            if not model:
                empty += 1
                self.preview_tree.insert("", "end",
                    values=(idx + 1, "(空型号，跳过)", info or "—", "—"),
                    tags=("empty", "even" if idx % 2 == 0 else "odd"))
                continue
            try:
                num = int(num_raw)
            except ValueError:
                try:
                    num = int(self.default_num.get().strip())
                except ValueError:
                    num = 0
            tags = ["even"] if idx % 2 == 0 else ["odd"]
            if model in self.existing_models:
                tags.append("dup")
                if self.dup_strategy.get() == "skip":
                    dup += 1
                else:
                    over += 1
            valid += 1
            self.preview_tree.insert("", "end",
                values=(idx + 1, model, info or "(空)", num), tags=tuple(tags))
        total = len(self.raw_rows)
        self.preview_count.config(text=f"显示前 {limit} 行" if total > limit else f"共 {total} 行")
        parts = [f"共 {total} 行"]
        if valid: parts.append(f"有效 {valid} 条")
        if empty: parts.append(f"空行 {empty} 条")
        if dup:   parts.append(f"跳过重复 {dup} 条")
        if over:  parts.append(f"覆盖更新 {over} 条")
        self.stats_label.config(text=" | ".join(parts))

    @staticmethod
    def _safe_get(row, idx):
        return row[idx] if idx < len(row) else ""

    def _confirm(self):
        mi = self.model_combo.current()
        ii = self.info_combo.current()
        ni = self.num_combo.current()
        try:
            fallback = int(self.default_num.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "默认数量请输入整数！", parent=self.top)
            return
        imported = []
        skipped_empty = skipped_dup = overwritten = 0
        for row in self.raw_rows:
            model = str(self._safe_get(row, mi)).strip()
            info  = str(self._safe_get(row, ii)).strip()
            num_raw = str(self._safe_get(row, ni)).strip()
            if not model:
                skipped_empty += 1
                continue
            try:
                num = int(num_raw)
            except ValueError:
                num = fallback
            is_dup = model in self.existing_models
            if is_dup and self.dup_strategy.get() == "skip":
                skipped_dup += 1
                continue
            if is_dup:
                overwritten += 1
            imported.append((model, info, num))
        if not imported:
            messagebox.showinfo("提示", "没有可导入的有效数据！", parent=self.top)
            return
        self.result = {
            "data": imported,
            "stats": {"imported": len(imported), "skipped_empty": skipped_empty,
                      "skipped_dup": skipped_dup, "overwritten": overwritten},
        }
        self.top.destroy()

    def _cancel(self):
        self.top.destroy()

    def _center_on_parent(self):
        self.top.update_idletasks()
        pw = self.top.master.winfo_width()
        ph = self.top.master.winfo_height()
        px = self.top.master.winfo_rootx()
        py = self.top.master.winfo_rooty()
        w, h = 780, 560
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.top.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")


# ============================================================
# 主应用
# ============================================================
class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网络设备库存查询工具")
        self.root.geometry("900x600")
        self.root.configure(bg=COLORS["bg"])
        self.root.minsize(760, 480)

        # 加载配置 & 同步管理器
        self.config = load_config()
        self.sync = SyncManager(self.config.get("server_url", ""),
                                self.config.get("api_token", ""))

        # 加载本地数据
        self.data = load_data()

        # 启动时尝试从服务器拉取
        self._sync_pull_on_startup()

        self._setup_styles()
        self._build_header()
        self._build_search_bar()
        self._build_edit_panel()
        self._build_table()
        self._build_footer()
        self._refresh_table()
        self._update_status_indicator()

    # ---------- 启动同步 ----------
    def _sync_pull_on_startup(self):
        """后台线程拉取服务端数据"""
        if not self.sync.configured:
            return

        def _pull():
            data = self.sync.pull_all()
            if data:
                # 合并到本地（服务端为准）
                self.data = data
                save_data(self.data)
                # 切回主线程刷新 UI
                self.root.after(0, self._on_pull_success)
            else:
                self.root.after(0, self._update_status_indicator)

        threading.Thread(target=_pull, daemon=True).start()

    def _on_pull_success(self):
        self._refresh_table()
        self._update_status_indicator()

    # ---------- 后台同步（不阻塞 UI）----------
    def _sync_push(self, model, info, num):
        """后台推送单条数据到服务器"""
        def _push():
            self.sync.push_item(model, info, num)
            self.root.after(0, self._update_status_indicator)
        threading.Thread(target=_push, daemon=True).start()

    def _sync_delete(self, model):
        """后台从服务器删除"""
        def _del():
            self.sync.delete_item(model)
            self.root.after(0, self._update_status_indicator)
        threading.Thread(target=_del, daemon=True).start()

    def _sync_push_batch(self, items):
        """后台批量推送（导入后）"""
        def _batch():
            for model, info, num in items:
                self.sync.push_item(model, info, num)
            self.root.after(0, self._update_status_indicator)
        threading.Thread(target=_batch, daemon=True).start()

    # ---------- 样式 ----------
    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(".", font=(FONT_FAMILY, 10), background=COLORS["card_bg"])
        style.configure("TLabel", background=COLORS["card_bg"], foreground=COLORS["text_primary"])
        style.configure("Primary.TButton",
                        font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["primary"], foreground="white",
                        borderwidth=0, padding=(16, 6))
        style.map("Primary.TButton",
                  background=[("active", COLORS["primary_hover"]),
                              ("pressed", COLORS["primary_hover"])])
        style.configure("Success.TButton",
                        font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["success"], foreground="white",
                        borderwidth=0, padding=(12, 6))
        style.map("Success.TButton",
                  background=[("active", COLORS["success_hover"]),
                              ("pressed", COLORS["success_hover"])])
        style.configure("Warning.TButton",
                        font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["warning"], foreground="white",
                        borderwidth=0, padding=(12, 6))
        style.map("Warning.TButton",
                  background=[("active", COLORS["warning_hover"]),
                              ("pressed", COLORS["warning_hover"])])
        style.configure("Danger.TButton",
                        font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["danger"], foreground="white",
                        borderwidth=0, padding=(12, 6))
        style.map("Danger.TButton",
                  background=[("active", COLORS["danger_hover"]),
                              ("pressed", COLORS["danger_hover"])])
        style.configure("Outline.TButton",
                        font=(FONT_FAMILY, 10),
                        background=COLORS["card_bg"], foreground=COLORS["primary"],
                        borderwidth=1, padding=(12, 6))
        style.map("Outline.TButton",
                  background=[("active", "#E8F0FE"), ("pressed", "#D4E4F7")])
        style.configure("Import.TButton",
                        font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["card_bg"], foreground="#00897B",
                        borderwidth=1, padding=(12, 6))
        style.map("Import.TButton",
                  background=[("active", "#E0F2F1"), ("pressed", "#B2DFDB")])
        style.configure("Treeview",
                        font=(FONT_FAMILY, 10),
                        background=COLORS["card_bg"],
                        fieldbackground=COLORS["card_bg"],
                        foreground=COLORS["text_primary"],
                        rowheight=34, borderwidth=0)
        style.configure("Treeview.Heading",
                        font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["footer_bg"],
                        foreground=COLORS["text_primary"],
                        borderwidth=0, padding=(8, 6))
        style.map("Treeview",
                  background=[("selected", COLORS["tree_selected"])],
                  foreground=[("selected", COLORS["text_primary"])])

    # ---------- 顶部标题栏 ----------
    def _build_header(self):
        header = tk.Frame(self.root, bg=COLORS["header_bg"], height=56)
        header.pack(fill="x")
        header.pack_propagate(False)

        inner = tk.Frame(header, bg=COLORS["header_bg"])
        inner.pack(side="left", padx=20, pady=8)

        tk.Label(inner, text="📦 网络设备库存查询工具",
                 font=(FONT_FAMILY, 16, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w")
        tk.Label(inner, text="管理交换机、路由器等网络设备的库存信息",
                 font=(FONT_FAMILY, 9),
                 bg=COLORS["header_bg"], fg="#8FA4B8").pack(anchor="w")

        # 设置按钮
        self.settings_btn = tk.Label(header, text="⚙", font=(FONT_FAMILY, 18),
                                     bg=COLORS["header_bg"], fg="#8FA4B8",
                                     cursor="hand2")
        self.settings_btn.pack(side="right", padx=16)
        self.settings_btn.bind("<Button-1>", lambda e: self._open_config())
        self.settings_btn.bind("<Enter>", lambda e: self.settings_btn.config(fg=COLORS["header_fg"]))
        self.settings_btn.bind("<Leave>", lambda e: self.settings_btn.config(fg="#8FA4B8"))

    # ---------- 搜索栏 ----------
    def _build_search_bar(self):
        bar = tk.Frame(self.root, bg=COLORS["bg"])
        bar.pack(fill="x", padx=16, pady=(12, 0))

        tk.Label(bar, text="🔍", font=(FONT_FAMILY, 14), bg=COLORS["bg"]).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._live_search())

        entry = tk.Entry(bar, textvariable=self.search_var,
                         font=(FONT_FAMILY, 11),
                         bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                         insertbackground=COLORS["text_primary"],
                         relief="flat", bd=8,
                         highlightthickness=1,
                         highlightbackground=COLORS["border"],
                         highlightcolor=COLORS["primary"])
        entry.pack(side="left", fill="x", expand=True, ipady=3)
        entry.insert(0, "输入型号或关键词进行搜索…")
        entry.config(fg=COLORS["text_secondary"])
        entry.bind("<FocusIn>", self._on_search_focus_in)
        entry.bind("<FocusOut>", self._on_search_focus_out)
        self._search_placeholder = True

        ttk.Button(bar, text="✕ 清空", style="Outline.TButton",
                   command=self.clear_search).pack(side="left", padx=(8, 0))

    def _on_search_focus_in(self, event):
        if self._search_placeholder:
            event.widget.delete(0, "end")
            event.widget.config(fg=COLORS["text_primary"])
            self._search_placeholder = False

    def _on_search_focus_out(self, event):
        if not event.widget.get().strip():
            event.widget.insert(0, "输入型号或关键词进行搜索…")
            event.widget.config(fg=COLORS["text_secondary"])
            self._search_placeholder = True

    def _live_search(self):
        if self._search_placeholder:
            return
        self.search_item()

    # ---------- 编辑面板 ----------
    def _build_edit_panel(self):
        card = tk.Frame(self.root, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="x", padx=16, pady=(10, 0))

        title_bar = tk.Frame(card, bg=COLORS["card_bg"])
        title_bar.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(title_bar, text="📝 新增 / 修改型号信息",
                 font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(side="left")

        self.import_btn = ttk.Button(title_bar, text="📥 导入", style="Import.TButton",
                                     command=self._show_import_menu)
        self.import_btn.pack(side="right")

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill="x", padx=16, pady=(0, 6))

        tk.Label(form, text="设备型号", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=(2, 0))
        self.model_var = tk.StringVar()
        tk.Entry(form, textvariable=self.model_var, font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).grid(
            row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 8), ipady=4)

        tk.Label(form, text="产品详情", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).grid(
            row=0, column=1, sticky="w", padx=(0, 6), pady=(2, 0))
        self.info_var = tk.StringVar()
        tk.Entry(form, textvariable=self.info_var, font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).grid(
            row=1, column=1, sticky="ew", padx=(0, 10), pady=(0, 8), ipady=4)

        tk.Label(form, text="库存数量", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).grid(
            row=0, column=2, sticky="w", padx=(0, 6), pady=(2, 0))
        self.num_var = tk.StringVar(value="0")
        tk.Entry(form, textvariable=self.num_var, font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"],
                 width=8, justify="center").grid(
            row=1, column=2, sticky="w", padx=(0, 10), pady=(0, 8), ipady=4)

        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=2)

        btn_bar = tk.Frame(card, bg=COLORS["card_bg"])
        btn_bar.pack(fill="x", padx=16, pady=(0, 12))

        ttk.Button(btn_bar, text="💾 保存型号", style="Primary.TButton",
                   command=self.save_model).pack(side="left", padx=(0, 8))
        ttk.Button(btn_bar, text="＋ 数量 +1", style="Success.TButton",
                   command=lambda: self._change_num(1)).pack(side="left", padx=(0, 8))
        ttk.Button(btn_bar, text="－ 数量 -1", style="Warning.TButton",
                   command=lambda: self._change_num(-1)).pack(side="left", padx=(0, 8))
        ttk.Button(btn_bar, text="🗑 删除型号", style="Danger.TButton",
                   command=self._del_model).pack(side="right")

    # ---------- 服务器配置 ----------
    def _open_config(self):
        dlg = ConfigDialog(self.root, self.config, self.sync)
        if dlg.result_config is not None:
            self.config = dlg.result_config
            self.sync = SyncManager(self.config.get("server_url", ""),
                                    self.config.get("api_token", ""))
            self._update_status_indicator()
            # 配置保存后尝试拉取
            self._sync_pull_on_startup()

    # ---------- 导入菜单 ----------
    def _show_import_menu(self):
        menu = tk.Menu(self.root, tearoff=0, font=(FONT_FAMILY, 10),
                       bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                       activebackground=COLORS["tree_selected"],
                       activeforeground=COLORS["text_primary"],
                       relief="flat", bd=1)
        menu.add_command(label="📄 从 CSV 文件导入…", command=self._import_csv)
        menu.add_command(label="📋 从剪贴板粘贴…", command=self._import_clipboard)
        if HAS_OPENPYXL:
            menu.add_command(label="📊 从 Excel 文件导入…", command=self._import_excel)
        else:
            menu.add_command(label="📊 从 Excel 文件导入（需安装 openpyxl）", state="disabled")
        menu.add_separator()
        menu.add_command(label="ℹ️ 支持：CSV / 剪贴板（Tab/逗号） / Excel",
                         state="disabled", font=(FONT_FAMILY, 9))
        x = self.import_btn.winfo_rootx()
        y = self.import_btn.winfo_rooty() + self.import_btn.winfo_height()
        menu.post(x, y)

    # ---------- CSV 导入 ----------
    def _import_csv(self):
        path = filedialog.askopenfilename(
            title="选择 CSV 文件",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                all_rows = list(reader)
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取 CSV 文件：\n{e}")
            return
        if not all_rows:
            messagebox.showwarning("提示", "文件为空！")
            return
        self._do_import(all_rows[0], all_rows[1:])

    # ---------- 剪贴板导入 ----------
    def _import_clipboard(self):
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("提示", "剪贴板为空或内容不可读取！\n\n请从 Excel / 表格中复制数据后重试。")
            return
        if not text.strip():
            messagebox.showwarning("提示", "剪贴板内容为空！")
            return
        lines = text.strip().splitlines()
        first_parts = lines[0].split("\t")
        if len(first_parts) >= 2:
            delimiter = "\t"
        else:
            first_parts = lines[0].split(",")
            delimiter = "," if len(first_parts) >= 2 else "\t"

        rows = []
        for line in lines:
            if delimiter == ",":
                try:
                    parts = next(csv.reader(io.StringIO(line)))
                except StopIteration:
                    continue
            else:
                parts = line.split(delimiter)
            rows.append(parts)
        if len(rows) < 1:
            messagebox.showwarning("提示", "未识别到有效数据！")
            return
        first = rows[0]
        header_keywords = ["型号", "名称", "详情", "数量", "库存", "model", "name", "info", "num", "qty"]
        looks_like_header = any(
            any(kw in str(c).lower() for kw in header_keywords) for c in first
        )
        if looks_like_header and len(rows) > 1:
            headers, data_rows = rows[0], rows[1:]
        else:
            max_cols = max(len(r) for r in rows)
            headers = [f"列 {chr(65 + i)}" if i < 26 else f"列 {i + 1}" for i in range(max_cols)]
            data_rows = rows
        self._do_import(headers, data_rows)

    # ---------- Excel 导入 ----------
    def _import_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showwarning("提示", "需要安装 openpyxl 库才能读取 Excel 文件。\n\n请在终端执行：pip install openpyxl")
            return
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            all_rows = []
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) if c is not None else "" for c in row]
                if any(c.strip() for c in cells):
                    all_rows.append(cells)
            wb.close()
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取 Excel 文件：\n{e}")
            return
        if not all_rows:
            messagebox.showwarning("提示", "文件为空或无可读数据！")
            return
        self._do_import(all_rows[0], all_rows[1:])

    # ---------- 统一导入入口 ----------
    def _do_import(self, headers, data_rows):
        if not data_rows:
            messagebox.showwarning("提示", "没有数据行可导入！（仅检测到表头）")
            return
        dlg = ImportDialog(self.root, headers, data_rows, set(self.data.keys()))
        self.root.wait_window(dlg.top)
        if dlg.result is None:
            return
        imported_data = dlg.result["data"]
        stats = dlg.result["stats"]
        for model, info, num in imported_data:
            self.data[model] = {"info": info, "num": num}
        save_data(self.data)
        self._refresh_table()
        # 后台同步到服务器
        self._sync_push_batch(imported_data)

        msg = f"成功导入 {stats['imported']} 条记录！"
        if stats["overwritten"]:
            msg += f"\n覆盖更新 {stats['overwritten']} 条已有型号。"
        if stats["skipped_dup"]:
            msg += f"\n跳过 {stats['skipped_dup']} 条重复型号。"
        if stats["skipped_empty"]:
            msg += f"\n忽略 {stats['skipped_empty']} 条空行。"
        messagebox.showinfo("导入完成", msg)

    # ---------- 表格 ----------
    def _build_table(self):
        card = tk.Frame(self.root, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="both", expand=True, padx=16, pady=(10, 0))

        title_bar = tk.Frame(card, bg=COLORS["card_bg"])
        title_bar.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(title_bar, text="📋 库存列表",
                 font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(side="left")
        self.count_label = tk.Label(title_bar, text="",
                                    font=(FONT_FAMILY, 10),
                                    bg=COLORS["card_bg"], fg=COLORS["text_secondary"])
        self.count_label.pack(side="right")

        table_container = tk.Frame(card, bg=COLORS["card_bg"])
        table_container.pack(fill="both", expand=True, padx=(16, 4), pady=(0, 12))

        self.tree = ttk.Treeview(table_container,
                                 columns=("model", "info", "num"),
                                 show="headings", selectmode="browse")
        self.tree.heading("model", text="  设备型号")
        self.tree.heading("info", text="产品详情")
        self.tree.heading("num", text="数量", anchor="center")
        self.tree.column("model", width=200, minwidth=120)
        self.tree.column("info", width=460, minwidth=200)
        self.tree.column("num", width=80, minwidth=60, anchor="center")

        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._select_row)
        self.tree.tag_configure("even", background=COLORS["tree_even"])
        self.tree.tag_configure("odd", background=COLORS["tree_odd"])
        self.tree.tag_configure("zero_stock", foreground="#B0BEC5")

    # ---------- 底部状态栏 ----------
    def _build_footer(self):
        footer = tk.Frame(self.root, bg=COLORS["footer_bg"], height=28)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        # 连接状态指示器
        self.status_dot = tk.Label(footer, text="", font=(FONT_FAMILY, 9),
                                   bg=COLORS["footer_bg"], fg=COLORS["text_secondary"])
        self.status_dot.pack(side="left", padx=16, pady=4)

        self.stats_label = tk.Label(footer, text="",
                                    font=(FONT_FAMILY, 9),
                                    bg=COLORS["footer_bg"], fg=COLORS["text_secondary"])
        self.stats_label.pack(side="right", padx=16, pady=4)

    def _update_status_indicator(self):
        """更新底部状态栏的连接状态"""
        if not self.sync.configured:
            self.status_dot.config(text="⚪ 未配置服务器", fg=COLORS["text_secondary"])
            return

        # 后台检测连接
        def _check():
            online = self.sync.check_health()
            def _update():
                if online:
                    self.status_dot.config(text="🟢 在线", fg=COLORS["online"])
                else:
                    self.status_dot.config(text="🔴 离线模式", fg=COLORS["offline"])
            self.root.after(0, _update)
        threading.Thread(target=_check, daemon=True).start()

    def _update_stats(self):
        total = len(self.data)
        total_qty = sum(item["num"] for item in self.data.values())
        self.stats_label.config(text=f"共 {total} 种型号 | 库存合计 {total_qty} 台")

    # ---------- 表格操作 ----------
    def _refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for idx, (model, item) in enumerate(self.data.items()):
            tag = "even" if idx % 2 == 0 else "odd"
            if item["num"] == 0:
                tag = ("zero_stock", tag)
            self.tree.insert("", "end", values=(model, item["info"], item["num"]), tags=tag)
        self._update_stats()
        cnt = len(self.data)
        self.count_label.config(text=f"共 {cnt} 条记录" if cnt else "暂无数据")

    def _select_row(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        item = self.tree.item(sel[0])
        model, info, num = item["values"]
        self.model_var.set(model)
        self.info_var.set(info)
        self.num_var.set(str(num))

    # ---------- 搜索 ----------
    def search_item(self):
        keyword = self.search_var.get().strip().lower()
        if not keyword:
            self._refresh_table()
            return
        for row in self.tree.get_children():
            self.tree.delete(row)
        cnt = 0
        for model, item in self.data.items():
            if keyword in model.lower() or keyword in item["info"].lower():
                tag = "even" if cnt % 2 == 0 else "odd"
                if item["num"] == 0:
                    tag = ("zero_stock", tag)
                self.tree.insert("", "end", values=(model, item["info"], item["num"]), tags=tag)
                cnt += 1
        self.count_label.config(text=f"搜索到 {cnt} 条结果" if keyword else "")

    def clear_search(self):
        self.search_var.set("")
        self._refresh_table()
        for child in self.root.winfo_children():
            if isinstance(child, tk.Frame):
                try:
                    if child.cget("bg") == COLORS["bg"]:
                        for w in child.winfo_children():
                            if isinstance(w, tk.Entry):
                                w.delete(0, "end")
                                w.insert(0, "输入型号或关键词进行搜索…")
                                w.config(fg=COLORS["text_secondary"])
                                self._search_placeholder = True
                                return
                except tk.TclError:
                    pass

    # ---------- 增删改 ----------
    def _change_num(self, delta):
        model = self.model_var.get().strip()
        if model not in self.data:
            messagebox.showwarning("提示", "请先在表格中选择一个型号，或先保存该型号！")
            return
        try:
            current = int(self.data[model]["num"])
            new_num = current + delta
            if new_num < 0:
                messagebox.showerror("错误", "库存数量不能小于 0 ！")
                return
            self.data[model]["num"] = new_num
            self.num_var.set(str(new_num))
            save_data(self.data)
            self._refresh_table()
            # 后台同步
            self._sync_push(model, self.data[model]["info"], new_num)
        except ValueError:
            messagebox.showerror("错误", "数量必须是数字！")

    def save_model(self):
        model = self.model_var.get().strip()
        info = self.info_var.get().strip()
        num_text = self.num_var.get().strip()
        if not model:
            messagebox.showerror("错误", "设备型号不能为空！")
            return
        try:
            num = int(num_text)
            if num < 0:
                messagebox.showerror("错误", "库存数量不能为负数！")
                return
        except ValueError:
            messagebox.showerror("错误", "数量请输入整数数字！")
            return
        is_new = model not in self.data
        self.data[model] = {"info": info, "num": num}
        save_data(self.data)
        self._refresh_table()
        # 后台同步
        self._sync_push(model, info, num)
        action = "新增" if is_new else "更新"
        messagebox.showinfo("成功", f"型号「{model}」已{action}！")

    def _del_model(self):
        model = self.model_var.get().strip()
        if model not in self.data:
            messagebox.showwarning("提示", f"型号「{model}」不存在！")
            return
        if messagebox.askyesno("确认删除", f"确定要删除型号「{model}」吗？\n\n此操作不可恢复。"):
            del self.data[model]
            save_data(self.data)
            self._refresh_table()
            self.model_var.set("")
            self.info_var.set("")
            self.num_var.set("0")
            # 后台同步删除
            self._sync_delete(model)


if __name__ == "__main__":
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()
