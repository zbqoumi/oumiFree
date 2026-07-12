"""欧米freeGPT注册机 — ChatGPT 自动注册（不加入工作区，直接导出 JSON）"""
import sys, json, threading, time, os, subprocess, traceback, logging
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

_DEVICE_CODE_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
_DEVICE_CODE_AUTHORITY = "https://login.microsoftonline.com/consumers"
_DEVICE_CODE_SCOPE = ["https://graph.microsoft.com/Mail.Read"]


def _device_code_get_token():
    try:
        from msal import PublicClientApplication
    except ImportError:
        return None, None, None, "msal 未安装，请运行: pip install msal"
    try:
        app = PublicClientApplication(
            client_id=_DEVICE_CODE_CLIENT_ID,
            authority=_DEVICE_CODE_AUTHORITY,
        )
        flow = app.initiate_device_flow(scopes=_DEVICE_CODE_SCOPE)
        if "error" in flow:
            return None, None, None, f"{flow.get('error')}: {flow.get('error_description', '')}"
        if "user_code" not in flow:
            return None, None, None, f"Device flow 返回异常: {json.dumps(flow, ensure_ascii=False)[:200]}"
        return flow.get("user_code"), flow.get("verification_uri"), flow, None
    except Exception as e:
        return None, None, None, str(e)


def _device_code_poll(flow, timeout=120):
    try:
        from msal import PublicClientApplication
    except ImportError:
        return None, None, "msal 未安装"
    app = PublicClientApplication(
        client_id=_DEVICE_CODE_CLIENT_ID,
        authority=_DEVICE_CODE_AUTHORITY,
    )
    deadline = time.time() + timeout
    while time.time() < deadline:
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            rt = result.get("refresh_token", "")
            if rt:
                return _DEVICE_CODE_CLIENT_ID, rt, None
        err = result.get("error", "")
        if err and err != "authorization_pending":
            return None, None, result.get("error_description", err)[:100]
        time.sleep(2)
    return None, None, "超时未授权"


path = str(Path(__file__).resolve().parent)
if path not in sys.path:
    sys.path.insert(0, path)

from gmailreg.blackcat_engine import BlackCatEngine

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
if not getattr(sys, "frozen", False):
    APP_DIR = Path(__file__).resolve().parent

CONFIG_FILE = APP_DIR / "blackcat_config.json"
LOGGER = logging.getLogger("omega_freegpt")

# ═══════════════════════════════════════════════════════════════════
# 配色方案 — 暗色主题 + 蓝紫渐变质感
# ═══════════════════════════════════════════════════════════════════
C_BG        = "#0d1117"   # 主窗口背景
C_FRAME     = "#161b22"   # 面板背景
C_FRAME_HD  = "#1c2333"   # 面板标题背景
C_ACCENT    = "#58a6ff"   # 主强调色（蓝）
C_ACCENT2   = "#7ee787"   # 成功色（绿）
C_DANGER    = "#f85149"   # 危险色（红）
C_WARN      = "#d29922"   # 警告色（黄）
C_TEXT      = "#e6edf3"   # 主文字
C_TEXT_DIM  = "#8b949e"   # 次要文字
C_ENTRY_BG  = "#0d1117"   # 输入框背景
C_LOG_BG    = "#010409"   # 日志背景
C_BTN_BG    = "#1f6feb"   # 普通按钮背景（蓝）
C_BTN_HOVER = "#388bfd"   # 普通按钮 hover
C_BTN_FG    = "#010409"   # 普通按钮文字（深色）


def setup_logging(log_path=None):
    logging.basicConfig(
        level=logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path or "blackcat_debug.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    LOGGER.info("欧米freeGPT注册机 启动 — " + datetime.now().isoformat())
    LOGGER.info("APP_DIR: " + str(APP_DIR))


def load_config():
    if CONFIG_FILE.exists():
        return json.load(open(CONFIG_FILE, encoding="utf-8"))
    return {}


def save_config(cfg):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    json.dump(cfg, open(CONFIG_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════════════════════════
# 自定义控件
# ═══════════════════════════════════════════════════════════════════

class DarkFrame(tk.LabelFrame):
    """暗色主题 LabelFrame"""
    def __init__(self, parent, text="", **kw):
        bg = kw.pop("bg", C_FRAME)
        super().__init__(parent, text="  " + text + "  ", bg=bg, fg=C_ACCENT,
                         font=("Microsoft YaHei UI", 10, "bold"),
                         bd=1, relief="solid",
                         highlightbackground="#30363d", highlightthickness=1,
                         labelanchor="n", **kw)
        self.configure(padx=12, pady=8)


class AccentButton(tk.Button):
    """强调色按钮"""
    def __init__(self, parent, text="", command=None, bg=C_ACCENT, fg="#010409",
                 hover_bg=None, font=("Microsoft YaHei UI", 10, "bold"), **kw):
        self._bg = bg
        self._hover_bg = hover_bg or bg
        super().__init__(parent, text=text, command=command, bg=bg, fg=fg,
                         font=font, relief="flat", bd=0, cursor="hand2",
                         activebackground=self._hover_bg, activeforeground=fg,
                         padx=16, pady=6, **kw)
        self.bind("<Enter>", lambda _: self.configure(bg=self._hover_bg))
        self.bind("<Leave>", lambda _: self.configure(bg=self._bg))


def _style_entry(entry):
    entry.configure(bg=C_ENTRY_BG, fg=C_TEXT, insertbackground=C_TEXT,
                    relief="solid", bd=1, highlightbackground="#30363d",
                    highlightthickness=1, font=("Consolas", 10))


def _make_btn(parent, text, command, width=None, bg=C_BTN_BG, hover_bg=C_BTN_HOVER, fg=C_BTN_FG):
    """创建带 hover 效果的可见按钮"""
    btn = tk.Button(parent, text=text, command=command, bg=bg, fg=fg,
                    relief="flat", bd=0, cursor="hand2",
                    font=("Microsoft YaHei UI", 9), activebackground=hover_bg, activeforeground=fg,
                    padx=10, pady=3)
    if width:
        btn.configure(width=width)
    btn.bind("<Enter>", lambda _: btn.configure(bg=hover_bg))
    btn.bind("<Leave>", lambda _: btn.configure(bg=bg))
    return btn


class BlackCatApp:
    def __init__(self, root):
        root.title("欧米freeGPT注册机")
        root.geometry("820x860")
        root.resizable(True, True)
        root.configure(bg=C_BG)
        self.root = root
        self.running = False
        self._stop_event = threading.Event()
        self._otp_event = threading.Event()
        self._otp_code = None
        self._stats_ok = 0
        self._stats_fail = 0
        self._stats_total = 0
        self._stats_json = 0
        self._stats_start = 0
        self._stats_lock = threading.Lock()

        cfg = load_config()

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TCheckbutton", background=C_FRAME, foreground=C_TEXT, font=("Microsoft YaHei UI", 9))
        style.configure("TRadiobutton", background=C_FRAME, foreground=C_TEXT, font=("Microsoft YaHei UI", 9))
        style.configure("TSpinbox", fieldbackground=C_ENTRY_BG, foreground=C_TEXT, arrowcolor=C_ACCENT)
        style.configure("TScrollbar", background=C_FRAME, troughcolor=C_BG)
        style.configure("TProgressbar", background=C_ACCENT, troughcolor=C_FRAME)

        root.columnconfigure(0, weight=1)

        # ═════════════════════ 品牌头部 ═════════════════════
        header = tk.Frame(root, bg=C_BG)
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))

        tk.Label(header, text="欧米", font=("Microsoft YaHei UI", 22, "bold"),
                 bg=C_BG, fg=C_ACCENT).pack(side="left")
        tk.Label(header, text="freeGPT 注册机", font=("Microsoft YaHei UI", 22, "bold"),
                 bg=C_BG, fg=C_TEXT).pack(side="left")
        tk.Label(header, text="  ·  ChatGPT 批量注册", font=("Microsoft YaHei UI", 10),
                 bg=C_BG, fg=C_TEXT_DIM).pack(side="left", padx=(0, 0), pady=(8, 0))

        row_num = 1  # 后续 Frame 从 row=1 开始

        # ═════════════════════ Frame 1: 邮箱配置 ═════════════════════
        f1 = DarkFrame(root, text="邮箱配置")
        f1.grid(row=row_num, column=0, sticky="ew", padx=12, pady=3); row_num += 1
        f1.columnconfigure(1, weight=1)

        mode_frame = tk.Frame(f1, bg=C_FRAME)
        mode_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self.mode_var = tk.StringVar(value=cfg.get("mode", "gmail"))
        self.mode_var.trace_add("write", lambda *a: self._on_mode_change())

        tk.Label(mode_frame, text="模式", font=("Microsoft YaHei UI", 9),
                 bg=C_FRAME, fg=C_TEXT_DIM).pack(side="left", padx=(0, 8))
        for label, val in [("Gmail", "gmail"), ("Outlook", "outlook"), ("API接码", "api"),
                           ("Outlook令牌", "outlook_token"), ("Token生成", "token_gen")]:
            ttk.Radiobutton(mode_frame, text=label, variable=self.mode_var, value=val).pack(side="left", padx=(0, 6))

        # Email
        tk.Label(f1, text="邮箱地址", bg=C_FRAME, fg=C_TEXT_DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=1, column=0, sticky="w", pady=(0, 4), padx=(4, 8))
        self.email_var = tk.StringVar(value=cfg.get("email", ""))
        self.email_entry = tk.Entry(f1, textvariable=self.email_var)
        _style_entry(self.email_entry)
        self.email_entry.grid(row=1, column=1, sticky="ew", pady=(0, 4))

        # Password
        self.pw_label = tk.Label(f1, text="应用专用密码", bg=C_FRAME, fg=C_TEXT_DIM,
                                 font=("Microsoft YaHei UI", 9))
        self.pw_label.grid(row=2, column=0, sticky="w", pady=(0, 8), padx=(4, 8))
        self.pw_frame = tk.Frame(f1, bg=C_FRAME)
        self.pw_frame.grid(row=2, column=1, sticky="ew", pady=(0, 8))
        self.pass_var = tk.StringVar(value=cfg.get("password", ""))
        self.pass_entry = tk.Entry(self.pw_frame, textvariable=self.pass_var, show="*")
        _style_entry(self.pass_entry)
        self.pass_entry.pack(side="left", fill="x", expand=True)
        self.show_pw = False
        self.pw_btn = _make_btn(self.pw_frame, "显示", self._toggle_pw, width=4)

        # Gmail hint
        self.gmail_hint = tk.Label(f1, text="不是账号密码，是 Google 应用专用密码", bg=C_FRAME, fg=C_TEXT_DIM,
                                   font=("Microsoft YaHei UI", 8))
        self.gmail_hint.grid(row=3, column=1, sticky="w", pady=(0, 8))

        # Outlook hint + code entry
        self.outlook_hint = tk.Label(f1, text="Outlook 不支持自动收信，请到网页版查验证码后填入下方",
                                     bg=C_FRAME, fg=C_WARN, font=("Microsoft YaHei UI", 8))
        self.outlook_hint.grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self.outlook_code_label = tk.Label(f1, text="验证码", bg=C_FRAME, fg=C_TEXT_DIM,
                                           font=("Microsoft YaHei UI", 9))
        self.outlook_code_label.grid(row=5, column=0, sticky="w", padx=(4, 0))
        self.outlook_code_var = tk.StringVar()
        self.outlook_code_entry = tk.Entry(f1, textvariable=self.outlook_code_var, width=12)
        _style_entry(self.outlook_code_entry)
        self.outlook_code_entry.configure(font=("Consolas", 9))
        self.outlook_code_entry.grid(row=5, column=1, sticky="w", pady=(4, 2))
        self.outlook_confirm_btn = AccentButton(f1, text="确认验证码", command=self._confirm_otp,
                                                 bg=C_ACCENT, hover_bg="#79b8ff")
        self.outlook_confirm_btn.grid(row=5, column=2, sticky="w", padx=(4, 0), pady=2)

        # API list
        self.api_list_label = tk.Label(f1, text="邮箱列表（一行一个：邮箱----API链接）",
                                       bg=C_FRAME, fg=C_TEXT_DIM, font=("Microsoft YaHei UI", 8))
        self.api_list_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 4))
        self.api_list_text = tk.Text(f1, height=4, font=("Consolas", 9), bg=C_ENTRY_BG, fg=C_TEXT,
                                      insertbackground=C_TEXT, relief="solid", bd=1,
                                      highlightbackground="#30363d", highlightthickness=1)
        self.api_list_text.grid(row=7, column=0, columnspan=3, sticky="ew")
        self.api_list_text.insert("1.0", cfg.get("api_list", ""))

        # ═════════════════════ Frame 2: Codex2API 上传配置 ═════════════════════
        f2 = DarkFrame(root, text="Codex2API 上传配置")
        f2.grid(row=row_num, column=0, sticky="ew", padx=12, pady=3); row_num += 1
        f2.columnconfigure(1, weight=1)

        tk.Label(f2, text="服务器地址", bg=C_FRAME, fg=C_TEXT_DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=0, column=0, sticky="w", padx=(4, 8))
        self.cp_url_var = tk.StringVar(value=cfg.get("cp_url", ""))
        e1 = tk.Entry(f2, textvariable=self.cp_url_var); _style_entry(e1)
        e1.grid(row=0, column=1, sticky="ew", pady=(0, 6))

        tk.Label(f2, text="Admin 密钥", bg=C_FRAME, fg=C_TEXT_DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=1, column=0, sticky="w", padx=(4, 8))
        key_frame = tk.Frame(f2, bg=C_FRAME)
        key_frame.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        self.cp_key_var = tk.StringVar(value=cfg.get("cp_key", ""))
        self.cp_key_entry = tk.Entry(key_frame, textvariable=self.cp_key_var, show="*")
        _style_entry(self.cp_key_entry)
        self.cp_key_entry.pack(side="left", fill="x", expand=True)
        self.show_key = False
        self.key_btn = _make_btn(key_frame, "显示", self._toggle_key, width=4)
        self.key_btn.pack(side="left", padx=(4, 0))

        tk.Label(f2, text="代理地址", bg=C_FRAME, fg=C_TEXT_DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=2, column=0, sticky="w", padx=(4, 8))
        self.proxy_var = tk.StringVar(value=cfg.get("proxy", "127.0.0.1:7897"))
        e2 = tk.Entry(f2, textvariable=self.proxy_var); _style_entry(e2)
        e2.grid(row=2, column=1, sticky="ew", pady=(0, 8))

        self.upload_var = tk.BooleanVar(value=cfg.get("upload", True))
        ttk.Checkbutton(f2, text="注册后自动上传到 Codex2API",
                        variable=self.upload_var, onvalue=True, offvalue=False).grid(row=3, column=0, columnspan=2, sticky="w")

        # ═════════════════════ 账号保存路径 ═════════════════════
        f_save = DarkFrame(root, text="账号保存路径")
        f_save.grid(row=row_num, column=0, sticky="ew", padx=12, pady=3); row_num += 1
        f_save.columnconfigure(0, weight=1)
        save_row = tk.Frame(f_save, bg=C_FRAME)
        save_row.grid(row=0, column=0, sticky="ew")
        save_row.columnconfigure(0, weight=1)
        self.save_dir_var = tk.StringVar(value=cfg.get("save_dir", str(APP_DIR / "codex")))
        sd = tk.Entry(save_row, textvariable=self.save_dir_var); _style_entry(sd)
        sd.grid(row=0, column=0, sticky="ew")
        b1 = _make_btn(save_row, "浏览", self._browse_save_dir)
        b1.grid(row=0, column=1, padx=(4, 0))
        b2 = _make_btn(save_row, "打开目录", self._open_save_dir)
        b2.grid(row=0, column=2, padx=(4, 0))

        # ═════════════════════ Frame 3: 注册控制（精简版） ═════════════════════
        f3 = DarkFrame(root, text="注册控制")
        f3.grid(row=row_num, column=0, sticky="ew", padx=12, pady=3); row_num += 1
        f3.columnconfigure(0, weight=1)

        ctrl = tk.Frame(f3, bg=C_FRAME)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(2, 8))
        ctrl.columnconfigure(4, weight=1)

        # 数量
        tk.Label(ctrl, text="注册数量", bg=C_FRAME, fg=C_TEXT_DIM,
                 font=("Microsoft YaHei UI", 10)).grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.count_var = tk.IntVar(value=cfg.get("count", 1))
        sp1 = ttk.Spinbox(ctrl, from_=1, to=999, textvariable=self.count_var, width=6,
                          font=("Consolas", 11))
        sp1.grid(row=0, column=1, sticky="w", padx=(0, 20))

        # 线程数
        tk.Label(ctrl, text="并发线程", bg=C_FRAME, fg=C_TEXT_DIM,
                 font=("Microsoft YaHei UI", 10)).grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.threads_var = tk.IntVar(value=cfg.get("threads", 1))
        sp2 = ttk.Spinbox(ctrl, from_=1, to=20, textvariable=self.threads_var, width=6,
                          font=("Consolas", 11))
        sp2.grid(row=0, column=3, sticky="w", padx=(0, 20))

        # 状态
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = tk.Label(ctrl, textvariable=self.status_var, bg=C_FRAME, fg=C_ACCENT2,
                                     font=("Microsoft YaHei UI", 10, "bold"))
        self.status_label.grid(row=0, column=4, sticky="w")

        # 按钮行
        btn_row = tk.Frame(f3, bg=C_FRAME)
        btn_row.grid(row=1, column=0, sticky="ew", pady=(4, 6))
        btn_row.columnconfigure(0, weight=1)

        self.start_btn = AccentButton(btn_row, text="  开始注册  ", command=self._start,
                                       bg=C_ACCENT2, hover_bg="#5fdb77", font=("Microsoft YaHei UI", 12, "bold"))
        self.start_btn.grid(row=0, column=0, padx=(0, 8))

        self.cancel_btn = AccentButton(btn_row, text="  停止  ", command=self._cancel,
                                        bg=C_DANGER, hover_bg="#ff6b63", font=("Microsoft YaHei UI", 12, "bold"),
                                        state="disabled")
        self.cancel_btn.grid(row=0, column=1)

        # 进度条
        self.progress = ttk.Progressbar(f3, mode="indeterminate")
        self.progress.grid(row=2, column=0, sticky="ew", pady=(6, 6))

        # 统计栏
        self.stats_var = tk.StringVar(value="")
        self.stats_label = tk.Label(f3, textvariable=self.stats_var, bg=C_LOG_BG, fg=C_ACCENT2,
                                    font=("Consolas", 11, "bold"), anchor="w", padx=10, pady=6)
        self.stats_label.grid(row=3, column=0, sticky="ew", pady=(4, 0))

        # ═════════════════════ Frame 4: 日志 ═════════════════════
        f4 = DarkFrame(root, text="运行日志")
        f4.grid(row=row_num, column=0, sticky="nsew", padx=12, pady=(3, 10))
        root.rowconfigure(row_num, weight=1)
        f4.columnconfigure(0, weight=1)
        f4.rowconfigure(0, weight=1)

        self.log_text = tk.Text(f4, wrap="word", font=("Consolas", 9),
                                bg=C_LOG_BG, fg=C_TEXT, padx=8, pady=8,
                                insertbackground=C_TEXT, relief="flat", bd=0,
                                highlightbackground="#30363d", highlightthickness=1)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_configure("ok", foreground=C_ACCENT2)
        self.log_text.tag_configure("warn", foreground=C_WARN)
        self.log_text.tag_configure("err", foreground=C_DANGER)
        self.log_text.tag_configure("info", foreground=C_ACCENT)
        self.log_text.tag_configure("time", foreground="#6e7681", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("step", foreground="#6e7681", font=("Consolas", 9, "bold"))

        log_scroll = ttk.Scrollbar(f4, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self._get_save_dir()
        self._on_mode_change()
        self._log("欧米freeGPT注册机 就绪")
        self._log("保存路径: " + self._get_save_dir())

    # ────────────────────────── 辅助方法 ──────────────────────────

    def _get_save_dir(self):
        d = self.save_dir_var.get().strip()
        if not d or d == "codex":
            d = str(APP_DIR / "codex")
        return d

    def _browse_save_dir(self):
        path = filedialog.askdirectory(title="选择账号保存目录", initialdir=self._get_save_dir())
        if path:
            self.save_dir_var.set(path)

    def _open_save_dir(self):
        d = Path(self._get_save_dir())
        d.mkdir(parents=True, exist_ok=True)
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(d)])
        elif sys.platform == "win32":
            subprocess.Popen(f'explorer "{d}"', shell=True)
        else:
            subprocess.Popen(["xdg-open", str(d)])

    def _toggle_key(self):
        self.show_key = not self.show_key
        self.cp_key_entry.configure(show="" if self.show_key else "*")
        self.key_btn.configure(text="隐藏" if self.show_key else "显示")

    def _toggle_pw(self):
        self.show_pw = not self.show_pw
        self.pass_entry.configure(show="" if self.show_pw else "*")
        self.pw_btn.configure(text="隐藏" if self.show_pw else "显示")

    def _on_mode_change(self):
        mode = self.mode_var.get()

        def hide_all_email_fields():
            self.email_entry.grid_remove()
            self.pw_label.grid_remove()
            self.pw_frame.grid_remove()
            self.gmail_hint.grid_remove()
            self.outlook_hint.grid_remove()
            self.outlook_code_label.grid_remove()
            self.outlook_code_entry.grid_remove()
            self.outlook_confirm_btn.grid_remove()
            self.api_list_label.grid_remove()
            self.api_list_text.grid_remove()

        if mode == "outlook":
            hide_all_email_fields()
            self.outlook_hint.grid()
            self.outlook_code_label.grid()
            self.outlook_code_entry.grid()
            self.outlook_confirm_btn.grid()
        elif mode == "api":
            hide_all_email_fields()
            self.api_list_label.configure(text="邮箱列表（一行一个：邮箱----API链接）")
            self.api_list_label.grid()
            self.api_list_text.grid()
        elif mode == "outlook_token":
            hide_all_email_fields()
            self.api_list_label.configure(text="邮箱列表（一行一个：邮箱----密码----client_id----refresh_token）")
            self.api_list_label.grid()
            self.api_list_text.grid()
        elif mode == "token_gen":
            hide_all_email_fields()
            self.api_list_label.configure(text="输入：邮箱----密码（一行一个），通过 Device Code 获取 refresh_token")
            self.api_list_label.grid()
            self.api_list_text.grid()
        else:  # gmail
            hide_all_email_fields()
            self.email_entry.grid()
            self.pw_label.grid()
            self.pw_frame.grid()
            self.gmail_hint.grid()

    def _confirm_otp(self):
        code = self.outlook_code_var.get().strip()
        if code.isdigit() and len(code) == 6:
            self._otp_code = code
            self._otp_event.set()

    def _log(self, text, level="info"):
        LOGGER.info(text)
        ts = datetime.now().strftime("%H:%M:%S")
        tag = "info"
        if text.startswith(("OK", "成功", "完成", "注册完成")):
            tag = "ok"
        elif any(k in text for k in ("失败", "错误", "[错误]", "出错", "FAIL")):
            tag = "err"
        elif any(k in text for k in ("跳过", "取消", "停止", "未获取")):
            tag = "warn"
        elif text.startswith("[") and "/9]" in text:
            tag = "step"
        self.log_text.insert(tk.END, f"[{ts}] ", "time")
        self.log_text.insert(tk.END, text + "\n", tag)
        self.log_text.see(tk.END)

    def _save_cfg(self):
        account_list = ""
        if self.mode_var.get() in ("api", "outlook_token", "token_gen"):
            account_list = self.api_list_text.get("1.0", "end").strip()
        save_config({
            "mode": self.mode_var.get(),
            "email": self.email_var.get(),
            "password": self.pass_var.get(),
            "cp_url": self.cp_url_var.get(),
            "cp_key": self.cp_key_var.get(),
            "count": self.count_var.get(),
            "threads": self.threads_var.get(),
            "proxy": self.proxy_var.get(),
            "upload": self.upload_var.get(),
            "save_dir": self._get_save_dir(),
            "api_list": account_list,
        })

    # ────────────────────────── 注册流程 ──────────────────────────

    def _start(self):
        if self.running:
            return
        self._save_cfg()
        count = self.count_var.get()
        threads = self.threads_var.get()
        do_upload = self.upload_var.get()
        mode = self.mode_var.get()

        if mode not in ("api", "outlook_token", "token_gen") and not self.email_var.get():
            messagebox.showerror("错误", "请填写邮箱")
            return

        # ── Token 生成模式：独立流程 ──
        if mode == "token_gen":
            self.running = True
            self._stop_event.clear()
            self._stats_ok = 0
            self._stats_fail = 0
            self.start_btn.configure(state="disabled")
            self.cancel_btn.configure(state="normal")
            self.progress.start()
            self.status_var.set("生成Token中...")
            self.log_text.delete("1.0", tk.END)
            threading.Thread(target=self._run_token_gen, daemon=True).start()
            return

        self.running = True
        self._stop_event.clear()
        self._otp_event.clear()
        self._otp_code = None
        self._stats_ok = 0
        self._stats_fail = 0
        self._stats_total = 0
        self._stats_json = 0
        self._stats_start = time.time()
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.progress.start()
        self.status_var.set("运行中...")
        self.log_text.delete("1.0", tk.END)
        self._log("保存路径: " + self._get_save_dir())
        self.root.after(1000, self._refresh_stats)

        cp_url = self.cp_url_var.get()
        cp_key = self.cp_key_var.get()
        account_list_modes = ("api", "outlook_token")
        imap_pass = "" if mode in account_list_modes else self.pass_var.get()

        # 解析账号列表
        api_accounts = []
        if mode in account_list_modes:
            for line in self.api_list_text.get("1.0", "end").strip().split("\n"):
                if not line.strip():
                    continue
                parts = line.strip().split("----")
                if mode == "api" and len(parts) == 2:
                    api_accounts.append({"email": parts[0].strip(), "api_url": parts[1].strip()})
                elif mode == "outlook_token" and len(parts) == 4:
                    api_accounts.append({
                        "email": parts[0].strip(),
                        "password": parts[1].strip(),
                        "client_id": parts[2].strip(),
                        "refresh_token": parts[3].strip(),
                    })
            if not api_accounts:
                fmt = "邮箱----API链接" if mode == "api" else "邮箱----密码----client_id----refresh_token"
                messagebox.showerror("错误", f"请填写至少一行 {fmt}")
                self._set_idle()
                return

        # 保存引擎参数（每个线程会独立创建引擎实例）
        self._engine_params = dict(
            mode=mode,
            imap_user=self.email_var.get() if mode not in account_list_modes else "",
            imap_pass=imap_pass,
            codexproxy_url=cp_url if do_upload else "",
            codexproxy_key=cp_key if do_upload else "",
            save_dir=self._get_save_dir(),
            manual_otp_cb=self._get_manual_otp if mode == "outlook" else None,
            accounts=api_accounts,
        )

        threading.Thread(target=self._run_batch, args=(count, threads, api_accounts), daemon=True).start()

    def _get_manual_otp(self, target_email, timeout=120):
        self._otp_event.clear()
        self._otp_code = None
        self.outlook_code_var.set("")
        self._log(f"请到 Outlook 网页查看 {target_email} 的验证码，填入后点确认")
        self._otp_event.wait(timeout=timeout)
        return self._otp_code

    def _cancel(self):
        if not self.running:
            return
        self._log("\n用户取消，正在停止...")
        self._stop_event.set()
        self._otp_event.set()
        self.stats_var.set("正在停止...")

    def _make_engine(self):
        """创建独立的引擎实例（线程安全）"""
        return BlackCatEngine(
            log_cb=lambda msg: self.root.after(0, lambda m=msg: self._log(m)),
            stop_event=self._stop_event,
            **self._engine_params,
        )

    def _run_batch(self, cnt, threads, api_accounts):
        try:
            mode = self._engine_params["mode"]
            account_list_modes = ("api", "outlook_token")

            if mode in account_list_modes:
                # 账号列表模式：按轮数 x 账号数生成任务
                tasks = []
                for rnd in range(cnt):
                    for acct in api_accounts:
                        tasks.append(acct)
                self._stats_total = len(tasks)
                mode_desc = "API接码" if mode == "api" else "Outlook令牌"
                self.root.after(0, lambda: self._log(
                    f"\n{mode_desc}模式: {len(api_accounts)} 个邮箱 x {cnt} 轮 = {len(tasks)} 次 | {threads} 线程\n"))

                def worker(task_id, acct):
                    if self._stop_event.is_set():
                        return
                    engine = self._make_engine()
                    if mode == "api":
                        engine._current_api_url = acct.get("api_url", "")
                    elif mode == "outlook_token":
                        engine.imap_user = acct["email"]
                        engine.imap_pass = acct["password"]
                        engine.outlook_client_id = acct["client_id"]
                        engine.outlook_refresh_token = acct["refresh_token"]
                        engine._token_cache = {}
                    self.root.after(0, lambda a=acct: self._log(f"[线程] 开始: {a['email']}"))
                    success = False
                    for attempt in range(3):
                        if self._stop_event.is_set():
                            break
                        ok, data = engine.register(
                            acct["email"],
                            callback=lambda msg: self.root.after(0, lambda m=msg: self._log(m)))
                        if ok:
                            success = True
                            with self._stats_lock:
                                self._stats_ok += 1
                                if isinstance(data, dict) and data.get("access_token"):
                                    self._stats_json += 1
                            break
                    if not success:
                        with self._stats_lock:
                            self._stats_fail += 1
            else:
                # 简单模式：cnt 次注册
                self._stats_total = cnt
                self.root.after(0, lambda: self._log(
                    f"\n开始注册: {cnt} 个 | {threads} 线程\n"))

                def worker(task_id, _acct=None):
                    if self._stop_event.is_set():
                        return
                    engine = self._make_engine()
                    self.root.after(0, lambda t=task_id: self._log(f"[线程-{t}] 开始注册"))
                    success = False
                    for attempt in range(3):
                        if self._stop_event.is_set():
                            break
                        ok, data = engine.register(
                            callback=lambda msg: self.root.after(0, lambda m=msg: self._log(m)))
                        if ok:
                            success = True
                            with self._stats_lock:
                                self._stats_ok += 1
                                if isinstance(data, dict) and data.get("access_token"):
                                    self._stats_json += 1
                            break
                    if not success:
                        with self._stats_lock:
                            self._stats_fail += 1

            # ── 提交线程池 ──
            with ThreadPoolExecutor(max_workers=threads) as pool:
                futures = []
                for i, task in enumerate(tasks if mode in account_list_modes else range(cnt)):
                    futures.append(pool.submit(worker, i, task))
                for future in as_completed(futures):
                    if self._stop_event.is_set():
                        # 取消尚未开始的任务
                        for f in futures:
                            f.cancel()
                        break

            self._print_summary()
        except Exception as e:
            self.root.after(0, lambda msg=str(e): self._log(f"\n[错误] {msg}", "err"))
            LOGGER.error(traceback.format_exc())
        finally:
            self.root.after(0, self._set_idle)

    # ────────────────────────── Token 生成 ──────────────────────────

    def _run_token_gen(self):
        try:
            lines = self.api_list_text.get("1.0", "end").strip().split("\n")
            accounts = []
            for line in lines:
                line = line.strip()
                if not line or "----" not in line:
                    continue
                parts = line.split("----")
                if len(parts) >= 2:
                    accounts.append((parts[0].strip(), parts[1].strip()))

            if not accounts:
                self.root.after(0, lambda: messagebox.showerror("错误", "请输入至少一行 邮箱----密码"))
                self.root.after(0, self._set_idle)
                return

            self._stats_total = len(accounts)
            self.root.after(0, lambda: self._log(f"开始生成 Token，共 {len(accounts)} 个账号\n"))

            results = []
            for i, (email, password) in enumerate(accounts):
                if self._stop_event.is_set():
                    break
                self.root.after(0, lambda e=email, n=i, t=len(accounts): self._log(f"[{n+1}/{t}] {e} — Device Code 授权中..."))
                client_id, refresh_token, error = self._device_code_for_account(email, password)

                if refresh_token:
                    token_line = f"{email}----{password}----{client_id}----{refresh_token}"
                    results.append(token_line)
                    self._stats_ok += 1
                    self.root.after(0, lambda e=email: self._log(f"  成功: {e}"))
                else:
                    self._stats_fail += 1
                    self.root.after(0, lambda e=email, err=error: self._log(f"  失败: {e} — {err}", "err"))

            if results:
                self.root.after(0, lambda: self._log("\n" + "=" * 50))
                self.root.after(0, lambda: self._log("▼ 复制以下 Token ▼", "ok"))
                self.root.after(0, lambda: self._log("=" * 50))
                for line in results:
                    self.root.after(0, lambda l=line: self._log(l, "ok"))

            self.root.after(0, lambda: self._log(
                f"\n完成: 成功 {self._stats_ok} | 失败 {self._stats_fail}"))
        except Exception as e:
            self.root.after(0, lambda msg=str(e): self._log(f"\n[错误] {msg}", "err"))
            LOGGER.error(traceback.format_exc())
        finally:
            self.root.after(0, self._set_idle)

    def _device_code_for_account(self, email, password):
        user_code, verify_uri, flow, error = _device_code_get_token()
        if error:
            return None, None, error

        self.root.after(0, lambda: self._log(
            f"\n  >>> 浏览器打开: {verify_uri}\n  >>> 输入验证码: {user_code}\n  >>> 用 {email} 登录\n", "warn"))

        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", verify_uri])
            elif sys.platform == "win32":
                subprocess.Popen(f'start "" "{verify_uri}"', shell=True)
        except Exception:
            pass

        confirmed = {"done": False}
        def _ask():
            confirmed["done"] = messagebox.askokcancel(
                "Device Code 授权",
                f"验证码: {user_code}\n\n浏览器已打开 {verify_uri}\n\n请用 {email} 登录后点确定"
            )
        self.root.after(0, _ask)
        while not confirmed["done"] and not self._stop_event.is_set():
            time.sleep(0.3)

        if self._stop_event.is_set() or not confirmed["done"]:
            return None, None, "用户取消"

        self.root.after(0, lambda: self._log("  正在获取 Token..."))
        return _device_code_poll(flow, timeout=60)

    # ────────────────────────── 统计与收尾 ──────────────────────────

    def _print_summary(self):
        self.root.after(0, lambda: self._log("\n" + "=" * 50))
        self.root.after(0, lambda: self._log(
            f"本次执行: {self._stats_total} 个 | 成功 {self._stats_ok} | 失败 {self._stats_fail} | 产出 {self._stats_json} 个 JSON | 保存至 {self._get_save_dir()}", "ok"))
        self.root.after(0, lambda: self._log("=" * 50))

    def _set_idle(self):
        self.running = False
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.progress.stop()
        self.status_var.set("就绪")

    def _refresh_stats(self):
        if not self.running:
            return
        elapsed = time.time() - self._stats_start
        self.stats_var.set(
            f"运行中... 成功 {self._stats_ok} | 失败 {self._stats_fail} | JSON {self._stats_json} | {elapsed:.0f}s")
        self.root.after(2000, self._refresh_stats)


def _run_playwright_smoke_test():
    try:
        from playwright._impl._driver import compute_driver_executable
        from playwright.sync_api import sync_playwright
        playwright_pkg = sync_playwright()
        node_path = compute_driver_executable()
        cli_path = Path(__file__).parent / "driver" / "package"
        print(f"playwright package: {playwright_pkg}")
        print(f"driver node: {node_path} exists={Path(node_path).exists()}")
        print(f"driver cli: {cli_path}")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
        print("chromium ready: True")
        print("playwright smoke ok")
        return 0
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    if "--smoke-playwright" in sys.argv:
        sys.exit(_run_playwright_smoke_test())

    log_path = "blackcat_debug.log"
    setup_logging(log_path)
    try:
        root = tk.Tk()
        app = BlackCatApp(root)
        root.mainloop()
    except Exception:
        LOGGER.critical(traceback.format_exc())
