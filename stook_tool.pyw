import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import csv
import io
import json
import os
import subprocess
import threading
import urllib.request

# ============================================================
# 路径配置
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

# GitHub 仓库配置
GITHUB_REPO = "g850043976-ux/stock-sync"
GITHUB_BRANCH = "master"
GITHUB_RAW_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/data.json"

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


def migrate_data(data):
    """将旧格式 {model: {info,unit,num}} 迁移为新格式 {id: {model,info,unit,num}}"""
    if not data:
        return data, 1
    # 判断是否已是新格式（第一个 key 是否为纯数字字符串）
    first_key = next(iter(data))
    if first_key.isdigit():
        next_id = max(int(k) for k in data) + 1
        return data, next_id
    # 迁移旧格式
    new_data = {}
    for i, (model, item) in enumerate(data.items(), 1):
        new_data[str(i)] = {
            "tax": item.get("tax", ""),
            "model": model,
            "info": item.get("info", ""),
            "unit": item.get("unit", ""),
            "num": item.get("num", 0),
        }
    return new_data, len(new_data) + 1


def get_next_id(data):
    if not data:
        return 1
    return max(int(k) for k in data) + 1


# ============================================================
# 配置读写
# ============================================================
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ============================================================
# GitHub 数据操作
# ============================================================
class GitHubManager:
    """通过 Git 推送数据 + raw URL 拉取数据"""

    def __init__(self):
        self.repo_dir = BASE_DIR

    def push_data(self, callback=None):
        """后台 git add + commit + push data.json"""
        def _do():
            # 隐藏 Windows 命令行窗口
            si = None
            if os.name == "nt":
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE
            git_env = {**os.environ, "GIT_SSH_COMMAND": "ssh -o StrictHostKeyChecking=accept-new -o BatchMode=yes"}
            try:
                subprocess.run(
                    ["git", "add", "data.json"],
                    cwd=self.repo_dir, capture_output=True, text=True,
                    timeout=15, env=git_env, startupinfo=si
                )
                subprocess.run(
                    ["git", "commit", "-m", "update data"],
                    cwd=self.repo_dir, capture_output=True, text=True,
                    timeout=15, startupinfo=si
                )
                result = subprocess.run(
                    ["git", "push", "origin", GITHUB_BRANCH],
                    cwd=self.repo_dir, capture_output=True, text=True,
                    timeout=30, env=git_env, startupinfo=si
                )
                ok = result.returncode == 0
            except Exception:
                ok = False
            if callback:
                callback(ok)

        threading.Thread(target=_do, daemon=True).start()

    def pull_data(self):
        """从 GitHub raw URL 拉取 data.json，返回 dict 或 None"""
        try:
            req = urllib.request.Request(GITHUB_RAW_URL)
            req.add_header("Cache-Control", "no-cache")
            with urllib.request.urlopen(req, timeout=15) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception:
            return None


# ============================================================
# 首次运行模式选择对话框
# ============================================================
class ModeSelectDialog:
    def __init__(self, parent):
        self.mode = None
        self.top = tk.Toplevel()
        self.top.title("选择运行模式")
        self.top.geometry("440x300")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.configure(bg=COLORS["bg"])
        self._center(parent)
        self._build()
        self.top.wait_window()

    def _build(self):
        header = tk.Frame(self.top, bg=COLORS["header_bg"], height=42)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="📦  选择运行模式", font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w", padx=16, pady=9)

        card = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="both", expand=True, padx=12, pady=10)

        inner = tk.Frame(card, bg=COLORS["card_bg"])
        inner.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(inner, text="请选择本设备的运行模式：", font=(FONT_FAMILY, 11),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(anchor="w", pady=(0, 14))

        # 管理模式
        mgr = tk.Frame(inner, bg=COLORS["card_bg"], bd=1, relief="solid",
                       highlightbackground=COLORS["primary"], highlightthickness=2)
        mgr.pack(fill="x", pady=(0, 10))
        mgr.bind("<Button-1>", lambda e: self._choose("manage"))
        for child in [mgr] + list(mgr.winfo_children()):
            try: child.bind("<Button-1>", lambda e: self._choose("manage"))
            except: pass

        tk.Label(mgr, text="✏️  管理模式", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["primary"]).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(mgr, text="可以新增、修改、删除型号。数据自动同步到 GitHub。",
                 font=(FONT_FAMILY, 9), bg=COLORS["card_bg"],
                 fg=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(0, 10))

        # 查看模式
        viewer = tk.Frame(inner, bg=COLORS["card_bg"], bd=1, relief="solid",
                          highlightbackground=COLORS["border"], highlightthickness=1)
        viewer.pack(fill="x")
        viewer.bind("<Button-1>", lambda e: self._choose("view"))
        for child in [viewer] + list(viewer.winfo_children()):
            try: child.bind("<Button-1>", lambda e: self._choose("view"))
            except: pass

        tk.Label(viewer, text="👁️  查看模式", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(viewer, text="只能查看库存数据，不可修改。数据自动从 GitHub 拉取。",
                 font=(FONT_FAMILY, 9), bg=COLORS["card_bg"],
                 fg=COLORS["text_secondary"]).pack(anchor="w", padx=12, pady=(0, 10))

        # 底部
        tk.Label(inner, text="💡 之后可以在标题栏 ⚙ 中切换模式",
                 font=(FONT_FAMILY, 9), bg=COLORS["card_bg"],
                 fg=COLORS["text_secondary"]).pack(pady=(8, 0))

    def _choose(self, mode):
        if mode == "manage":
            pwd = simpledialog.askstring("验证密码", "请输入管理模式密码：", show="*", parent=self.top)
            if pwd != "000000":
                if pwd is not None:
                    messagebox.showerror("密码错误", "密码不正确！", parent=self.top)
                return
        self.mode = mode
        self.top.destroy()

    def _center(self, parent):
        self.top.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w, h = 440, 300
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.top.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")


# ============================================================
# 导入预览对话框
# ============================================================
class ImportDialog:
    MODEL_KEYWORDS = ["型号", "设备型号", "设备", "名称", "model", "name", "产品型号", "物料"]
    TAX_KEYWORDS   = ["税收分类", "税收", "税率", "tax", "分类"]
    INFO_KEYWORDS  = ["详情", "产品详情", "描述", "项目", "项目名称", "info", "description", "备注", "说明", "规格"]
    UNIT_KEYWORDS  = ["单位", "unit", "计量单位"]
    NUM_KEYWORDS   = ["数量", "库存数量", "库存", "当前库存", "num", "quantity", "qty"]

    def __init__(self, parent, headers, raw_rows, existing_models):
        self.result = None
        self.headers = list(headers)
        self.raw_rows = raw_rows
        self.existing_models = existing_models
        self.tax_col = self._guess(self.TAX_KEYWORDS)
        self.unit_col = self._guess(self.UNIT_KEYWORDS)
        self.model_col = self._guess(self.MODEL_KEYWORDS)
        self.info_col  = self._guess(self.INFO_KEYWORDS)
        self.num_col   = self._guess(self.NUM_KEYWORDS)
        if len(self.headers) >= 2 and self.info_col == self.model_col:
            self.info_col = 1 if self.model_col == 0 else 0
        if len(self.headers) >= 3 and (self.num_col == self.model_col or self.num_col == self.info_col):
            for i in range(len(self.headers)):
                if i not in (self.model_col, self.info_col):
                    self.num_col = i; break
        self._build(); self._refresh()

    def _guess(self, keywords):
        for kw in keywords:
            for i, h in enumerate(self.headers):
                if kw.lower() in str(h).lower().strip(): return i
        return 0

    def _build(self):
        self.top = tk.Toplevel()
        self.top.title("导入预览")
        self.top.geometry("780x620")
        self.top.minsize(640, 500)
        self.top.transient(self.top.master)
        self.top.grab_set()
        self.top.configure(bg=COLORS["bg"])
        self.top.update_idletasks()
        m = self.top.master
        pw, ph = m.winfo_width(), m.winfo_height()
        px, py = m.winfo_rootx(), m.winfo_rooty()
        self.top.geometry(f"780x620+{max(0,px+(pw-780)//2)}+{max(0,py+(ph-620)//2)}")

        hdr = tk.Frame(self.top, bg=COLORS["header_bg"], height=42)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="📥  导入预览 — 确认列映射与数据", font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w", padx=16, pady=9)

        card = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="x", padx=12, pady=(10, 0))
        inner = tk.Frame(card, bg=COLORS["card_bg"]); inner.pack(fill="x", padx=14, pady=12)
        tk.Label(inner, text="列映射", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(anchor="w", pady=(0, 8))

        r1 = tk.Frame(inner, bg=COLORS["card_bg"]); r1.pack(fill="x", pady=(0, 6))
        for text, attr, w in [("税收分类 →", "tax_combo", 12), ("产品详情 →", "info_combo", 22),
                               ("设备型号 →", "model_combo", 16), ("单位 →", "unit_combo", 8),
                               ("库存数量 →", "num_combo", 10)]:
            tk.Label(r1, text=text, font=(FONT_FAMILY, 10),
                     bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
            cb = ttk.Combobox(r1, values=self.headers, state="readonly", width=w)
            cb.pack(side="left", padx=(0, 10))
            setattr(self, attr, cb)
            cb.bind("<<ComboboxSelected>>", lambda e: self._refresh())
        self.tax_combo.current(self.tax_col)
        self.model_combo.current(self.model_col)
        self.info_combo.current(self.info_col)
        self.unit_combo.current(self.unit_col)
        self.num_combo.current(self.num_col)

        tk.Label(r1, text="默认数量:", font=(FONT_FAMILY, 9),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self.default_num = tk.StringVar(value="0")
        tk.Entry(r1, textvariable=self.default_num, width=6, font=(FONT_FAMILY, 10), justify="center",
                 bg=COLORS["card_bg"], relief="solid", bd=1).pack(side="left")

        r2 = tk.Frame(inner, bg=COLORS["card_bg"]); r2.pack(fill="x")
        tk.Label(r2, text="重复型号:", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=(0, 8))
        self.dup_strategy = tk.StringVar(value="skip")
        ttk.Radiobutton(r2, text="跳过", variable=self.dup_strategy, value="skip").pack(side="left", padx=(0, 14))
        ttk.Radiobutton(r2, text="覆盖", variable=self.dup_strategy, value="overwrite").pack(side="left")

        # Preview table
        c2 = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                      highlightthickness=1, highlightbackground=COLORS["border"])
        c2.pack(fill="both", expand=True, padx=12, pady=(8, 0))
        tb = tk.Frame(c2, bg=COLORS["card_bg"]); tb.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(tb, text="数据预览", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(side="left")
        self.pcnt = tk.Label(tb, text="", font=(FONT_FAMILY, 9), bg=COLORS["card_bg"], fg=COLORS["text_secondary"])
        self.pcnt.pack(side="right")

        ct = tk.Frame(c2, bg=COLORS["card_bg"]); ct.pack(fill="both", expand=True, padx=(14,4), pady=(0,10))
        self.pt = ttk.Treeview(ct, columns=("row","tax","info","model","unit","num"), show="headings", selectmode="none", height=8)
        for c, t, w in [("row","行",36), ("tax","税收分类",80), ("info","产品详情",260),
                         ("model","设备型号",140), ("unit","单位",44), ("num","数量",54)]:
            self.pt.heading(c, text=t, anchor="center" if c in ("row","unit","num") else "w")
            self.pt.column(c, width=w, anchor="center" if c in ("row","unit","num") else "w", stretch=c=="info")
        vsb = ttk.Scrollbar(ct, orient="vertical", command=self.pt.yview)
        self.pt.configure(yscrollcommand=vsb.set)
        self.pt.pack(side="left", fill="both", expand=True); vsb.pack(side="right", fill="y")
        self.pt.tag_configure("empty", foreground="#B0BEC5")
        self.pt.tag_configure("dup", foreground="#E65100")
        self.pt.tag_configure("even", background=COLORS["tree_even"])
        self.pt.tag_configure("odd", background=COLORS["tree_odd"])

        # Bottom bar
        bar = tk.Frame(self.top, bg=COLORS["bg"]); bar.pack(fill="x", padx=12, pady=(8,12))
        self.slbl = tk.Label(bar, text="", font=(FONT_FAMILY, 10), bg=COLORS["bg"], fg=COLORS["text_secondary"])
        self.slbl.pack(side="left", pady=6)
        ttk.Button(bar, text="确认导入", style="Primary.TButton", command=self._confirm).pack(side="right", padx=(8,0))
        ttk.Button(bar, text="取消", style="Outline.TButton", command=self._cancel).pack(side="right")

    def _refresh(self):
        for row in self.pt.get_children(): self.pt.delete(row)
        ti = self.tax_combo.current()
        mi, ii, ui, ni = (self.model_combo.current(), self.info_combo.current(),
                           self.unit_combo.current(), self.num_combo.current())
        valid = empty = dup = 0
        limit = min(len(self.raw_rows), 100)
        for idx, row in enumerate(self.raw_rows[:limit]):
            tax   = str(row[ti] if ti < len(row) else "").strip()
            model = str(row[mi] if mi < len(row) else "").strip()
            info  = str(row[ii] if ii < len(row) else "").strip()
            unit  = str(row[ui] if ui < len(row) else "").strip()
            num_raw = str(row[ni] if ni < len(row) else "").strip()
            if not model:
                empty += 1
                self.pt.insert("", "end", values=(idx+1, tax or "", "(空详情)", "(空型号，跳过)", unit, "—"),
                               tags=("empty", "even" if idx%2==0 else "odd")); continue
            try: num = int(num_raw)
            except ValueError:
                try: num = int(self.default_num.get().strip())
                except ValueError: num = 0
            tags = ["even"] if idx%2==0 else ["odd"]
            if model in self.existing_models: tags.append("dup"); dup += 1
            valid += 1
            self.pt.insert("", "end", values=(idx+1, tax or "", info or "(空)", model, unit or "", num), tags=tuple(tags))
        total = len(self.raw_rows)
        self.pcnt.config(text=f"显示前 {limit} 行" if total > limit else f"共 {total} 行")
        parts = [f"共 {total} 行", f"有效 {valid} 条"]
        if empty: parts.append(f"空行 {empty} 条")
        if dup: parts.append(f"重复 {dup} 条")
        self.slbl.config(text=" | ".join(parts))

    def _confirm(self):
        ti = self.tax_combo.current()
        mi, ii, ui, ni = (self.model_combo.current(), self.info_combo.current(),
                           self.unit_combo.current(), self.num_combo.current())
        try: fallback = int(self.default_num.get().strip())
        except ValueError: messagebox.showwarning("提示", "默认数量请输入整数！", parent=self.top); return
        imported = []; se = sd = so = 0
        for row in self.raw_rows:
            tax   = str(row[ti] if ti < len(row) else "").strip()
            model = str(row[mi] if mi < len(row) else "").strip()
            info  = str(row[ii] if ii < len(row) else "").strip()
            unit  = str(row[ui] if ui < len(row) else "").strip()
            num_raw = str(row[ni] if ni < len(row) else "").strip()
            if not model: se += 1; continue
            try: num = int(num_raw)
            except ValueError: num = fallback
            if model in self.existing_models and self.dup_strategy.get() == "skip": sd += 1; continue
            if model in self.existing_models: so += 1
            imported.append((tax, model, info, unit, num))
        if not imported: messagebox.showinfo("提示", "没有可导入的有效数据！", parent=self.top); return
        self.result = {"data": imported, "stats": {"imported": len(imported), "skipped_empty": se,
                        "skipped_dup": sd, "overwritten": so}}
        self.top.destroy()

    def _cancel(self): self.top.destroy()


# ============================================================
# 修改弹窗
# ============================================================
class EditDialog:
    """点击"修改"按钮后弹出，编辑税收分类、设备型号、产品详情、单位"""

    def __init__(self, parent, tax, info, model, unit):
        self.result = None
        self.top = tk.Toplevel()
        self.top.title("修改型号信息")
        self.top.geometry("460x300")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.configure(bg=COLORS["bg"])
        self._center(parent)
        self._build(tax, info, model, unit)
        self.top.wait_window()

    def _build(self, tax, info, model, unit):
        hdr = tk.Frame(self.top, bg=COLORS["header_bg"], height=38)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="✏️  修改型号信息", font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w", padx=16, pady=7)

        card = tk.Frame(self.top, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="both", expand=True, padx=12, pady=10)

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill="x", padx=16, pady=14)

        # 税收分类
        tk.Label(form, text="税收分类", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w")
        self.tax_var = tk.StringVar(value=tax)
        tk.Entry(form, textvariable=self.tax_var, font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1, highlightthickness=1,
                 highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(fill="x", ipady=4, pady=(0, 10))

        # 产品详情
        tk.Label(form, text="产品详情", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w")
        self.info_var = tk.StringVar(value=info)
        tk.Entry(form, textvariable=self.info_var, font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1, highlightthickness=1,
                 highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(fill="x", ipady=4, pady=(0, 10))

        # 设备型号
        tk.Label(form, text="设备型号", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w")
        self.model_var = tk.StringVar(value=model)
        tk.Entry(form, textvariable=self.model_var, font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1, highlightthickness=1,
                 highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(fill="x", ipady=4, pady=(0, 10))

        # 单位
        tk.Label(form, text="单位", font=(FONT_FAMILY, 10),
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).pack(anchor="w")
        self.unit_var = tk.StringVar(value=unit)
        tk.Entry(form, textvariable=self.unit_var, font=(FONT_FAMILY, 10), width=10,
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1, highlightthickness=1,
                 highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(anchor="w", ipady=4, pady=(0, 10))

        # 按钮
        bar = tk.Frame(self.top, bg=COLORS["bg"]); bar.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(bar, text="💾 保存修改", style="Primary.TButton",
                   command=self._confirm).pack(side="right", padx=(8, 0))
        ttk.Button(bar, text="取消", style="Outline.TButton",
                   command=self._cancel).pack(side="right")

    def _confirm(self):
        model = self.model_var.get().strip()
        if not model:
            messagebox.showwarning("提示", "设备型号不能为空！", parent=self.top)
            return
        self.result = {
            "tax": self.tax_var.get().strip(),
            "model": model,
            "info": self.info_var.get().strip(),
            "unit": self.unit_var.get().strip(),
        }
        self.top.destroy()

    def _cancel(self):
        self.top.destroy()

    def _center(self, parent):
        self.top.update_idletasks()
        pw = parent.winfo_width(); ph = parent.winfo_height()
        px = parent.winfo_rootx(); py = parent.winfo_rooty()
        x = px + (pw - 460) // 2; y = py + (ph - 300) // 2
        self.top.geometry(f"460x300+{max(0,x)}+{max(0,y)}")


# ============================================================
# 主应用
# ============================================================
class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网络设备进项查询工具")
        self.root.geometry("900x600")
        self.root.configure(bg=COLORS["bg"])
        self.root.minsize(760, 480)

        self.config = load_config()
        self.mode = self.config.get("mode", "")
        self.github = GitHubManager()
        raw = load_data()
        self.data, self._next_id = migrate_data(raw)
        self.selected_id = None  # 当前选中记录的 ID
        self.batch_mode = False   # 批量删除模式
        self.batch_checked = set()  # 批量删除勾选的 ID 集合
        self._last_push_ok = True

        # 首次运行 → 选择模式
        if not self.mode:
            dlg = ModeSelectDialog(self.root)
            self.mode = dlg.mode or "view"
            self.config["mode"] = self.mode
            save_config(self.config)

        self._setup_styles()
        self._build_header()
        self._build_search_bar()

        if self.is_manage:
            self._build_edit_panel()
            self._build_table()
            self._build_footer()
            self._refresh_table()
        else:
            self._build_table()
            self._build_footer()
            self._refresh_from_github()

    @property
    def is_manage(self):
        return self.mode == "manage"

    # ---------- 样式 ----------
    def _setup_styles(self):
        style = ttk.Style(); style.theme_use("clam")
        style.configure(".", font=(FONT_FAMILY, 10), background=COLORS["card_bg"])
        style.configure("TLabel", background=COLORS["card_bg"], foreground=COLORS["text_primary"])
        style.configure("Primary.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["primary"], foreground="white", borderwidth=0, padding=(16,6))
        style.map("Primary.TButton", background=[("active", COLORS["primary_hover"]), ("pressed", COLORS["primary_hover"])])
        style.configure("Success.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["success"], foreground="white", borderwidth=0, padding=(12,6))
        style.map("Success.TButton", background=[("active", COLORS["success_hover"]), ("pressed", COLORS["success_hover"])])
        style.configure("Warning.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["warning"], foreground="white", borderwidth=0, padding=(12,6))
        style.map("Warning.TButton", background=[("active", COLORS["warning_hover"]), ("pressed", COLORS["warning_hover"])])
        style.configure("Danger.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["danger"], foreground="white", borderwidth=0, padding=(12,6))
        style.map("Danger.TButton", background=[("active", COLORS["danger_hover"]), ("pressed", COLORS["danger_hover"])])
        style.configure("Outline.TButton", font=(FONT_FAMILY, 10),
                        background=COLORS["card_bg"], foreground=COLORS["primary"], borderwidth=1, padding=(12,6))
        style.map("Outline.TButton", background=[("active", "#E8F0FE"), ("pressed", "#D4E4F7")])
        style.configure("Import.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["card_bg"], foreground="#00897B", borderwidth=1, padding=(12,6))
        style.map("Import.TButton", background=[("active", "#E0F2F1"), ("pressed", "#B2DFDB")])
        style.configure("Edit.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background="#FDD835", foreground="#333333", borderwidth=0, padding=(12,6))
        style.map("Edit.TButton", background=[("active", "#F9A825"), ("pressed", "#FBC02D")])
        style.configure("Reset.TButton", font=(FONT_FAMILY, 10, "bold"),
                        background="#7E57C2", foreground="white", borderwidth=0, padding=(12,6))
        style.map("Reset.TButton", background=[("active", "#6A1B9A"), ("pressed", "#5E35B1")])
        style.configure("Treeview", font=(FONT_FAMILY, 10),
                        background=COLORS["card_bg"], fieldbackground=COLORS["card_bg"],
                        foreground=COLORS["text_primary"], rowheight=34, borderwidth=0)
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 10, "bold"),
                        background=COLORS["footer_bg"], foreground=COLORS["text_primary"],
                        borderwidth=0, padding=(8,6))
        style.map("Treeview", background=[("selected", COLORS["tree_selected"])],
                  foreground=[("selected", COLORS["text_primary"])])

    # ---------- 标题栏 ----------
    def _build_header(self):
        hdr = tk.Frame(self.root, bg=COLORS["header_bg"], height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        inner = tk.Frame(hdr, bg=COLORS["header_bg"]); inner.pack(side="left", padx=20, pady=8)
        title = "📦 网络设备进项查询工具" + (" — 管理模式" if self.is_manage else " — 查看模式")
        tk.Label(inner, text=title, font=(FONT_FAMILY, 16, "bold"),
                 bg=COLORS["header_bg"], fg=COLORS["header_fg"]).pack(anchor="w")

        # 切换模式按钮（右上角）
        switch_text = "🔒 切换为查看模式" if self.is_manage else "🔓 切换为管理模式"
        self.switch_btn = tk.Label(hdr, text=switch_text, font=(FONT_FAMILY, 9),
                                   bg=COLORS["header_bg"], fg="#8FA4B8", cursor="hand2")
        self.switch_btn.pack(side="right", padx=16)
        self.switch_btn.bind("<Button-1>", lambda e: self._toggle_mode())
        self.switch_btn.bind("<Enter>", lambda e: self.switch_btn.config(fg=COLORS["header_fg"]))
        self.switch_btn.bind("<Leave>", lambda e: self.switch_btn.config(fg="#8FA4B8"))

    def _toggle_mode(self):
        if self.is_manage:
            # 管理 → 查看：直接切换
            if not messagebox.askyesno("切换模式", "确定要切换到查看模式吗？\n\n应用将重新启动。"):
                return
        else:
            # 查看 → 管理：需要密码
            pwd = simpledialog.askstring("验证密码", "请输入管理模式密码：", show="*", parent=self.root)
            if pwd != "000000":
                if pwd is not None:
                    messagebox.showerror("密码错误", "密码不正确，无法切换到管理模式！")
                return

        self.mode = "view" if self.is_manage else "manage"
        self.config["mode"] = self.mode
        save_config(self.config)
        self.root.destroy()
        os.startfile(__file__)

    # ---------- 搜索栏 ----------
    def _build_search_bar(self):
        bar = tk.Frame(self.root, bg=COLORS["bg"]); bar.pack(fill="x", padx=16, pady=(12, 0))
        tk.Label(bar, text="🔍", font=(FONT_FAMILY, 14), bg=COLORS["bg"]).pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._live_search())
        entry = tk.Entry(bar, textvariable=self.search_var, font=(FONT_FAMILY, 11),
                         bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                         relief="flat", bd=8, highlightthickness=1,
                         highlightbackground=COLORS["border"], highlightcolor=COLORS["primary"])
        entry.pack(side="left", fill="x", expand=True, ipady=3)
        entry.insert(0, "输入型号或关键词进行搜索…")
        entry.config(fg=COLORS["text_secondary"])
        entry.bind("<FocusIn>", self._on_focus_in); entry.bind("<FocusOut>", self._on_focus_out)
        self._placeholder = True
        if not self.is_manage:
            ttk.Button(bar, text="🔄 刷新数据", style="Outline.TButton",
                       command=self._refresh_from_github).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="✕ 清空", style="Outline.TButton",
                   command=self.clear_search).pack(side="left", padx=(8, 0))

    def _on_focus_in(self, event):
        if self._placeholder:
            event.widget.delete(0, "end"); event.widget.config(fg=COLORS["text_primary"]); self._placeholder = False

    def _on_focus_out(self, event):
        if not event.widget.get().strip():
            event.widget.insert(0, "输入型号或关键词进行搜索…")
            event.widget.config(fg=COLORS["text_secondary"]); self._placeholder = True

    def _live_search(self):
        if self._placeholder: return
        self.search_item()

    # ---------- 编辑面板（仅管理模式）----------
    def _build_edit_panel(self):
        card = tk.Frame(self.root, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="x", padx=16, pady=(10, 0))

        tb = tk.Frame(card, bg=COLORS["card_bg"]); tb.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(tb, text="📝 新增 / 修改型号信息", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(side="left")
        self.import_btn = ttk.Button(tb, text="📥 导入", style="Import.TButton",
                                     command=self._show_import_menu)
        self.import_btn.pack(side="right")

        form = tk.Frame(card, bg=COLORS["card_bg"]); form.pack(fill="x", padx=16, pady=(0, 6))
        for c, label, var, stretch in [
            (0, "税收分类", "tax_var", True), (1, "产品详情", "info_var", True),
            (2, "设备型号", "model_var", True), (3, "单位", "unit_var", False),
            (4, "库存数量", "num_var", False)
        ]:
            tk.Label(form, text=label, font=(FONT_FAMILY, 10),
                     bg=COLORS["card_bg"], fg=COLORS["text_secondary"]).grid(
                row=0, column=c, sticky="w", padx=(0, 6), pady=(2, 0))
            sv = tk.StringVar(value="0" if "数量" in label else "")
            setattr(self, var, sv)
            if "数量" in label:
                kw = {"width": 8, "justify": "center", "font": (FONT_FAMILY, 11, "bold")}
            elif "单位" in label:
                kw = {"width": 6, "justify": "center", "font": (FONT_FAMILY, 10)}
            else:
                kw = {"font": (FONT_FAMILY, 10)}
            tk.Entry(form, textvariable=sv, bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                     relief="solid", bd=1, highlightthickness=1,
                     highlightbackground=COLORS["border"], highlightcolor=COLORS["primary"], **kw).grid(
                row=1, column=c, sticky="ew" if stretch else "w", padx=(0, 10), pady=(0, 8), ipady=4)
        form.columnconfigure(0, weight=1); form.columnconfigure(1, weight=1); form.columnconfigure(2, weight=2)

        bb = tk.Frame(card, bg=COLORS["card_bg"]); bb.pack(fill="x", padx=16, pady=(0, 12))
        ttk.Button(bb, text="➕ 新增型号", style="Primary.TButton", command=self.save_model).pack(side="left", padx=(0, 8))
        # 数量增减组： [+] [输入框] [-]
        ttk.Button(bb, text="＋", style="Success.TButton", width=3,
                   command=lambda: self._apply_delta(1)).pack(side="left")
        self.delta_var = tk.StringVar(value="1")
        tk.Entry(bb, textvariable=self.delta_var, font=(FONT_FAMILY, 11, "bold"),
                 width=5, justify="center", bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                 relief="solid", bd=1, highlightthickness=1,
                 highlightbackground=COLORS["border"],
                 highlightcolor=COLORS["primary"]).pack(side="left", ipady=3)
        ttk.Button(bb, text="－", style="Warning.TButton", width=3,
                   command=lambda: self._apply_delta(-1)).pack(side="left", padx=(0, 8))
        ttk.Button(bb, text="修改", style="Edit.TButton", command=self._edit_model).pack(side="left", padx=(0, 8))
        ttk.Button(bb, text="重置", style="Reset.TButton", command=self._reset_form).pack(side="left", padx=(0, 8))
        ttk.Button(bb, text="🗑 删除型号", style="Danger.TButton", command=self._del_model).pack(side="right")

    # ---------- 表格 ----------
    def _build_table(self):
        card = tk.Frame(self.root, bg=COLORS["card_bg"], bd=0,
                        highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="both", expand=True, padx=16, pady=(10, 0))

        tb = tk.Frame(card, bg=COLORS["card_bg"]); tb.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(tb, text="📋 库存列表", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS["card_bg"], fg=COLORS["text_primary"]).pack(side="left")
        if self.is_manage:
            self.batch_btn = ttk.Button(tb, text="🗑 批量删除", style="Outline.TButton",
                                        command=self._toggle_batch_mode)
            self.batch_btn.pack(side="right", padx=(0, 6))
            self.batch_confirm_btn = ttk.Button(tb, text="确定删除", style="Danger.TButton",
                                                command=self._batch_delete)
        ttk.Button(tb, text="📋 提取", style="Outline.TButton", command=self._copy_row).pack(side="right", padx=(0, 6))
        self.count_label = tk.Label(tb, text="", font=(FONT_FAMILY, 10),
                                    bg=COLORS["card_bg"], fg=COLORS["text_secondary"])
        self.count_label.pack(side="right")

        ct = tk.Frame(card, bg=COLORS["card_bg"]); ct.pack(fill="both", expand=True, padx=(16, 4), pady=(0, 12))
        self.tree = ttk.Treeview(ct, columns=("chk","id","tax","info","model","unit","num"), show="headings", selectmode="browse")
        self.tree.heading("chk", text="☐", anchor="center")
        self.tree.heading("id", text="编码", anchor="center")
        self.tree.heading("tax", text="税收分类"); self.tree.heading("info", text="产品详情")
        self.tree.heading("model", text="设备型号")
        self.tree.heading("unit", text="单位", anchor="center"); self.tree.heading("num", text="数量", anchor="center")
        self.tree.column("chk", width=0, minwidth=0, stretch=False)
        self.tree.column("id", width=50, minwidth=40, anchor="center")
        self.tree.column("tax", width=100, minwidth=70)
        self.tree.column("info", width=330, minwidth=140)
        self.tree.column("model", width=150, minwidth=90)
        self.tree.column("unit", width=50, minwidth=40, anchor="center")
        self.tree.column("num", width=70, minwidth=50, anchor="center")
        vsb = ttk.Scrollbar(ct, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True); vsb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._select_row)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree.tag_configure("even", background=COLORS["tree_even"])
        self.tree.tag_configure("odd", background=COLORS["tree_odd"])
        self.tree.tag_configure("zero_stock", foreground="#B0BEC5")
        if self.is_manage:
            self.batch_confirm_btn.pack_forget()  # 初始隐藏

    # ---------- 底部状态栏 ----------
    def _build_footer(self):
        footer = tk.Frame(self.root, bg=COLORS["footer_bg"], height=28)
        footer.pack(fill="x", side="bottom"); footer.pack_propagate(False)
        self.status_label = tk.Label(footer, text="", font=(FONT_FAMILY, 9),
                                     bg=COLORS["footer_bg"], fg=COLORS["text_secondary"])
        self.status_label.pack(side="left", padx=16, pady=4)
        self.stats_label = tk.Label(footer, text="", font=(FONT_FAMILY, 9),
                                    bg=COLORS["footer_bg"], fg=COLORS["text_secondary"])
        self.stats_label.pack(side="right", padx=16, pady=4)

    def _update_status(self, push_ok=None):
        total = len(self.data)
        total_qty = sum(item["num"] for item in self.data.values())
        self.stats_label.config(text=f"共 {total} 种型号 | 库存合计 {total_qty} 台")
        if self.is_manage:
            status = "🟢 管理端" if (push_ok if push_ok is not None else self._last_push_ok) else "🟠 管理端（推送失败）"
            self.status_label.config(text=status)
        else:
            self.status_label.config(text="👁️ 查看端 | 数据来自 GitHub")

    # ---------- 从 GitHub 刷新（查看模式）----------
    def _refresh_from_github(self):
        self.status_label.config(text="⏳ 正在从 GitHub 获取数据…")
        self.root.update()

        def _do():
            data = self.github.pull_data()
            def _done():
                if data:
                    migrated, nid = migrate_data(data)
                    self.data = migrated
                    self._next_id = nid
                    save_data(migrated)
                    self._refresh_table()
                    self.status_label.config(text="👁️ 查看端 | 数据来自 GitHub")
                else:
                    self.status_label.config(text="🔴 无法连接 GitHub，显示本地缓存")
                    self._refresh_table()
            self.root.after(0, _done)

        threading.Thread(target=_do, daemon=True).start()

    # ---------- Git Push（管理模式）----------
    def _git_push(self):
        self._last_push_ok = True
        self._update_status()
        self.github.push_data(callback=lambda ok: self.root.after(0, lambda: self._update_status(ok)))

    # ---------- 表格操作 ----------
    def _refresh_table(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        for idx, (rid, item) in enumerate(self.data.items()):
            tag = "even" if idx % 2 == 0 else "odd"
            if item["num"] == 0: tag = ("zero_stock", tag)
            chk = "☑" if self.batch_mode and rid in self.batch_checked else ""
            self.tree.insert("", "end", iid=rid,
                             values=(chk, rid, item.get("tax",""), item.get("info",""), item.get("model",""),
                                     item.get("unit",""), item["num"]),
                             tags=tag)
        cnt = len(self.data)
        self.count_label.config(text=f"共 {cnt} 条记录" if cnt else "暂无数据")
        total_qty = sum(item["num"] for item in self.data.values())
        self.stats_label.config(text=f"共 {cnt} 条记录 | 库存合计 {total_qty} 台")

    def _on_tree_click(self, event):
        """点击表格空白区域时取消选中；批量模式下点击行切换勾选"""
        row = self.tree.identify_row(event.y)
        if not row:
            self.tree.selection_remove(self.tree.selection())
            return
        if self.batch_mode and row:
            # 批量模式：点击行切换勾选状态
            if row in self.batch_checked:
                self.batch_checked.discard(row)
            else:
                self.batch_checked.add(row)
            self._refresh_table()

    def _toggle_batch_mode(self):
        """切换批量删除模式"""
        self.batch_mode = not self.batch_mode
        self.batch_checked.clear()
        if self.batch_mode:
            self.tree.column("chk", width=36, minwidth=36)
            self.tree.heading("chk", text="☐")
            self.batch_btn.config(text="取消")
            self.batch_confirm_btn.pack(side="right", padx=(0, 6))
        else:
            self.tree.column("chk", width=0, minwidth=0)
            self.batch_btn.config(text="🗑 批量删除")
            self.batch_confirm_btn.pack_forget()
            self.selected_id = None
        self._refresh_table()

    def _batch_delete(self):
        """执行批量删除"""
        if not self.batch_checked:
            messagebox.showwarning("提示", "请先勾选要删除的记录！")
            return
        count = len(self.batch_checked)
        if not messagebox.askyesno("确认批量删除",
                                   f"确定删除已勾选的 {count} 条记录吗？\n\n此操作不可恢复。"):
            return
        for rid in list(self.batch_checked):
            if rid in self.data:
                del self.data[rid]
        save_data(self.data)
        self.batch_checked.clear()
        self._refresh_table()
        self._git_push()
        messagebox.showinfo("完成", f"已删除 {count} 条记录。")

    def _select_row(self, event):
        if self.batch_mode:
            return  # 批量模式下不更新编辑面板
        sel = self.tree.selection()
        if not sel:
            self.selected_id = None
            if self.is_manage:
                self.tax_var.set(""); self.model_var.set(""); self.info_var.set("")
                self.unit_var.set(""); self.num_var.set("0")
            return
        self.selected_id = sel[0]

    # ---------- 搜索 ----------
    def search_item(self):
        keyword = self.search_var.get().strip().lower()
        if not keyword: self._refresh_table(); return
        for row in self.tree.get_children(): self.tree.delete(row)
        cnt = 0
        for rid, item in self.data.items():
            model = item.get("model", "")
            if keyword in model.lower() or keyword in item.get("info","").lower():
                tag = "even" if cnt % 2 == 0 else "odd"
                if item["num"] == 0: tag = ("zero_stock", tag)
                chk = "☑" if self.batch_mode and rid in self.batch_checked else ""
                self.tree.insert("", "end", iid=rid,
                                 values=(chk, rid, item.get("tax",""), model, item.get("info",""),
                                         item.get("unit",""), item["num"]),
                                 tags=tag)
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
                                w.delete(0, "end"); w.insert(0, "输入型号或关键词进行搜索…")
                                w.config(fg=COLORS["text_secondary"]); self._placeholder = True; return
                except tk.TclError: pass

    # ---------- 导入菜单 ----------
    def _show_import_menu(self):
        if not self.is_manage: return
        menu = tk.Menu(self.root, tearoff=0, font=(FONT_FAMILY, 10),
                       bg=COLORS["card_bg"], fg=COLORS["text_primary"],
                       activebackground=COLORS["tree_selected"],
                       activeforeground=COLORS["text_primary"])
        menu.add_command(label="📄 从 CSV 文件导入…", command=self._import_csv)
        menu.add_command(label="📋 从剪贴板粘贴…", command=self._import_clipboard)
        if HAS_OPENPYXL:
            menu.add_command(label="📊 从 Excel 文件导入…", command=self._import_excel)
        else:
            menu.add_command(label="📊 从 Excel 文件导入（需安装 openpyxl）", state="disabled")
        x = self.import_btn.winfo_rootx()
        y = self.import_btn.winfo_rooty() + self.import_btn.winfo_height()
        menu.post(x, y)

    def _import_csv(self):
        path = filedialog.askopenfilename(title="选择 CSV 文件",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.reader(f))
        except Exception as e: messagebox.showerror("读取失败", str(e)); return
        if not rows: messagebox.showwarning("提示", "文件为空！"); return
        self._do_import(rows[0], rows[1:])

    def _import_clipboard(self):
        try: text = self.root.clipboard_get()
        except tk.TclError: messagebox.showwarning("提示", "剪贴板为空！"); return
        if not text.strip(): messagebox.showwarning("提示", "剪贴板内容为空！"); return
        lines = text.strip().splitlines()
        delim = "\t" if len(lines[0].split("\t")) >= 2 else ("," if len(lines[0].split(",")) >= 2 else "\t")
        rows = []
        for line in lines:
            parts = next(csv.reader(io.StringIO(line))) if delim == "," else line.split(delim)
            rows.append(parts)
        first = rows[0]; hk = ["型号","名称","详情","数量","库存","model","name","info","num","qty"]
        looks_header = any(any(kw in str(c).lower() for kw in hk) for c in first)
        headers, data = (rows[0], rows[1:]) if looks_header and len(rows) > 1 else (
            [f"列 {chr(65+i)}" if i < 26 else f"列 {i+1}" for i in range(max(len(r) for r in rows))], rows)
        self._do_import(headers, data)

    def _import_excel(self):
        if not HAS_OPENPYXL: messagebox.showwarning("提示", "需要 pip install openpyxl"); return
        path = filedialog.askopenfilename(title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if not path: return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True); ws = wb.active
            rows = [[str(c) if c is not None else "" for c in r] for r in ws.iter_rows(values_only=True)
                    if any(c is not None and str(c).strip() for c in r)]
            wb.close()
        except Exception as e: messagebox.showerror("读取失败", str(e)); return
        if not rows: messagebox.showwarning("提示", "文件为空！"); return
        self._do_import(rows[0], rows[1:])

    def _do_import(self, headers, data_rows):
        if not data_rows: messagebox.showwarning("提示", "没有数据行！"); return
        # 收集现有型号名（用于预览标记重复，不再用于阻止导入）
        existing_models = {item.get("model", "") for item in self.data.values()}
        dlg = ImportDialog(self.root, headers, data_rows, existing_models)
        self.root.wait_window(dlg.top)
        if dlg.result is None: return
        for tax, model, info, unit, num in dlg.result["data"]:
            new_id = str(self._next_id)
            self.data[new_id] = {"tax": tax, "model": model, "info": info, "unit": unit, "num": num}
            self._next_id += 1
        save_data(self.data)
        self._refresh_table()
        self._git_push()
        s = dlg.result["stats"]
        msg = f"成功导入 {s['imported']} 条记录！"
        if s["skipped_dup"]: msg += f"\n跳过 {s['skipped_dup']} 条重复。"
        messagebox.showinfo("导入完成", msg)

    def _reset_form(self):
        """清空新增表单"""
        self.tax_var.set(""); self.model_var.set(""); self.info_var.set("")
        self.unit_var.set(""); self.num_var.set("0")
        self.tree.selection_remove(self.tree.selection())

    def _copy_row(self):
        """将选中行内容复制为空格分隔的文本"""
        if not self.selected_id or self.selected_id not in self.data:
            messagebox.showwarning("提示", "请先在表格中点击选择一条记录！"); return
        item = self.data[self.selected_id]
        text = f"{item.get('tax','')} {item.get('info','')} {item.get('model','')} {item.get('unit','')}"
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制到剪贴板：\n{text}")

    # ---------- 增删改（仅管理模式）----------
    def _apply_delta(self, sign):
        """解析输入框中的数量，调用 _change_num"""
        try:
            delta = int(self.delta_var.get()) * sign
        except ValueError:
            messagebox.showerror("错误", "增减数量请输入整数！")
            return
        self._change_num(delta)

    def _change_num(self, delta):
        if not self.selected_id or self.selected_id not in self.data:
            messagebox.showwarning("提示", "请先在表格中选择一条记录！"); return
        if delta == 0:
            messagebox.showwarning("提示", "增减数量不能为 0！"); return
        try:
            cur = int(self.data[self.selected_id]["num"]); new = cur + delta
            if new < 0: messagebox.showerror("错误", "库存不能小于 0！"); return
            item = self.data[self.selected_id]
            label = f"+{delta}" if delta > 0 else f"{delta}"
            if not messagebox.askyesno("确认操作",
                                       f"确定将「{item.get('model','')}」数量 {label} 吗？\n\n"
                                       f"当前库存：{cur} → 新库存：{new}"):
                return
            self.data[self.selected_id]["num"] = new; self.num_var.set(str(new))
            save_data(self.data); self._refresh_table(); self._git_push()
            self.tree.selection_set(self.selected_id)
        except ValueError: messagebox.showerror("错误", "数量必须是数字！")

    def save_model(self):
        """新增型号 —— 永远创建新记录，不更新已有"""
        tax = self.tax_var.get().strip(); model = self.model_var.get().strip()
        info = self.info_var.get().strip(); unit = self.unit_var.get().strip()
        num_text = self.num_var.get().strip()
        if not model: messagebox.showerror("错误", "型号不能为空！"); return
        try:
            num = int(num_text)
            if num < 0: messagebox.showerror("错误", "库存不能为负数！"); return
        except ValueError: messagebox.showerror("错误", "数量请输入整数！"); return

        if not messagebox.askyesno("确认新增",
                                   f"确定新增型号「{model}」吗？\n\n"
                                   f"税收分类：{tax or '(空)'}\n"
                                   f"产品详情：{info or '(空)'}\n"
                                   f"单位：{unit or '(空)'}\n"
                                   f"库存数量：{num}"):
            return

        new_id = str(self._next_id)
        self.data[new_id] = {"tax": tax, "model": model, "info": info, "unit": unit, "num": num}
        self._next_id += 1
        self.selected_id = new_id
        save_data(self.data); self._refresh_table(); self._git_push()
        self.tree.selection_set(new_id)
        self.tree.see(new_id)
        messagebox.showinfo("成功", f"型号「{model}」已新增！（编码 {new_id}）")

    def _edit_model(self):
        """弹出修改窗口，只能修改已保存的记录"""
        if not self.selected_id or self.selected_id not in self.data:
            messagebox.showwarning("提示", "请先在表格中点击选择一条记录！"); return
        item = self.data[self.selected_id]
        dlg = EditDialog(self.root, item.get("tax", ""), item.get("info", ""), item.get("model", ""), item.get("unit", ""))
        if dlg.result is None:
            return
        # 更新记录
        self.data[self.selected_id]["tax"] = dlg.result["tax"]
        self.data[self.selected_id]["model"] = dlg.result["model"]
        self.data[self.selected_id]["info"] = dlg.result["info"]
        self.data[self.selected_id]["unit"] = dlg.result["unit"]
        save_data(self.data); self._refresh_table(); self._git_push()
        self.tree.selection_set(self.selected_id)
        self.tree.see(self.selected_id)
        # 同步更新编辑面板
        self.tax_var.set(dlg.result["tax"])
        self.model_var.set(dlg.result["model"])
        self.info_var.set(dlg.result["info"])
        self.unit_var.set(dlg.result["unit"])
        messagebox.showinfo("成功", f"编码 {self.selected_id} 已修改！")

    def _del_model(self):
        if not self.selected_id or self.selected_id not in self.data:
            messagebox.showwarning("提示", "请先在表格中选择一条记录！"); return
        item = self.data[self.selected_id]
        model = item.get("model", "")
        if messagebox.askyesno("确认删除", f"确定删除编码 {self.selected_id} 的「{model}」吗？\n\n此操作不可恢复。"):
            del self.data[self.selected_id]; save_data(self.data); self._refresh_table(); self._git_push()
            self.selected_id = None
            self.tax_var.set(""); self.model_var.set(""); self.info_var.set("")
            self.unit_var.set(""); self.num_var.set("0")


if __name__ == "__main__":
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()
