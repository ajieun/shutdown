# -*- coding: utf-8 -*-
"""
定时关机助手 (Shutdown Timer)
--------------------------------
可视化倒计时 / 定时关机 + 办公健康提醒（Windows / macOS）

功能：
  * 倒计时关机：分/秒自定义 + 常用快捷时长
  * 定时关机：图形化日历 + 时间选择
  * 动作可选：关机 / 重启 / 注销（macOS 为退出登录）
  * 喝水提醒：定时提醒补水，记录当日饮水杯数
  * 提肛助手：盆底肌收缩/放松节奏训练与阶段提示音
  * 走动提醒：久坐定时起身提醒，并统计活动次数与时长
  * 计划持久化：关闭程序后再次打开可继续显示剩余时间
  * 运行日志：本地文本日志，便于回溯

跨平台说明：
  * Windows：使用 shutdown.exe（支持 /t 延时与 /a 取消）
  * macOS  ：使用 AppleScript 驱动 System Events；延时任务以独立后台进程运行
  * 健康提醒部分为纯界面逻辑，两端行为一致

依赖：仅使用 Python 标准库（tkinter）
运行：Windows 双击 run.bat；macOS 运行 python3 shutdown_timer.py 或双击 run.command
"""

import base64
import calendar
import ctypes
import datetime as dt
import json
import math
import os
import signal
import struct
import subprocess
import sys
import threading
import time
import zlib
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk

APP_NAME = "ShutdownTimer"
APP_TITLE = "定时关机助手"
APP_VERSION = "1.2.0"

# --------------------------------------------------------------------------
# 平台判定
# --------------------------------------------------------------------------
# SHUTDOWN_TIMER_OS=mac 可在任意系统上模拟 macOS 行为（便于自测）
_FORCED_OS = os.environ.get("SHUTDOWN_TIMER_OS", "").strip().lower()
IS_WINDOWS = sys.platform.startswith("win") and _FORCED_OS != "mac"
IS_MAC = (sys.platform == "darwin" or _FORCED_OS == "mac") and not IS_WINDOWS
PLATFORM_LABEL = "Windows" if IS_WINDOWS else ("macOS" if IS_MAC else sys.platform)

# --------------------------------------------------------------------------
# 路径与常量
# --------------------------------------------------------------------------
def _data_dir() -> str:
    if IS_MAC:
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, APP_NAME)
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        path = os.path.dirname(os.path.abspath(__file__))
    return path


DATA_DIR = _data_dir()
STATE_FILE = os.path.join(DATA_DIR, "state.json")
LOG_FILE = os.path.join(DATA_DIR, "shutdown_timer.log")
SHUTDOWN_EXE = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "shutdown.exe")
CREATE_NO_WINDOW = 0x08000000

# 动作定义：Windows 用 shutdown.exe 参数，macOS 用 AppleScript
if IS_MAC:
    ACTIONS = {
        "关机": {"label": "关机", "win": "/s", "mac": 'tell application "System Events" to shut down',
                 "delay": True},
        "重启": {"label": "重启", "win": "/r", "mac": 'tell application "System Events" to restart',
                 "delay": True},
        "退出登录": {"label": "退出登录", "win": "/l",
                     "mac": 'tell application "System Events" to log out', "delay": True},
    }
else:
    ACTIONS = {
        "关机": {"label": "关机", "win": "/s", "mac": "", "delay": True},
        "重启": {"label": "重启", "win": "/r", "mac": "", "delay": True},
        "注销": {"label": "注销", "win": "/l", "mac": "", "delay": False},
    }

# --------------------------------------------------------------------------
# 配色 / 主题
# --------------------------------------------------------------------------
C = {
    "bg": "#0d1117",
    "panel": "#111823",
    "card": "#161d29",
    "card_hi": "#1e2736",
    "border": "#26303f",
    "accent": "#38bdf8",
    "accent_dim": "#0e7490",
    "danger": "#f87171",
    "success": "#4ade80",
    "warn": "#fbbf24",
    "text": "#e6edf3",
    "muted": "#8b98a9",
    "faint": "#5c6a7d",
}


# --------------------------------------------------------------------------
# 小工具函数
# --------------------------------------------------------------------------
WEEK_CN = ["日", "一", "二", "三", "四", "五", "六"]

FONT_UI = "TkDefaultFont"
FONT_MONO = "TkFixedFont"


def fmt_clock(seconds: float) -> str:
    seconds = int(max(0, seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return "%02d:%02d:%02d" % (h, m, s)


def fmt_natural(seconds: float) -> str:
    seconds = int(max(0, seconds))
    if seconds < 60:
        return "%d 秒" % seconds
    days, rem = divmod(seconds, 86400)
    h, rem = divmod(rem, 3600)
    m, s = divmod(rem, 60)
    parts = []
    if days:
        parts.append("%d 天" % days)
    if h:
        parts.append("%d 小时" % h)
    if m:
        parts.append("%d 分钟" % m)
    if not parts:
        parts.append("%d 秒" % s)
    return " ".join(parts)


def fmt_datetime(ts: float) -> str:
    return dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def pick_font(candidates, fallback="TkDefaultFont"):
    try:
        families = set(tkfont.families())
    except Exception:
        return fallback
    for name in candidates:
        if name in families:
            return name
    return fallback


# --------------------------------------------------------------------------
# 程序图标（纯 Python 生成 PNG，避免额外依赖）
# --------------------------------------------------------------------------
def _icon_pixel(x: float, y: float, size: int):
    radius = size * 0.22
    cx = min(max(x, radius), size - radius)
    cy = min(max(y, radius), size - radius)
    if math.hypot(x - cx, y - cy) > radius:
        return (0, 0, 0, 0)

    t = y / float(size)
    bg = (
        int(37 + 8 * (1 - t)),
        int(54 + 10 * (1 - t)),
        int(78 + 12 * (1 - t)),
    )

    pcx, pcy = size / 2.0, size / 2.0 + size * 0.02
    pr = size * 0.28
    thick = size * 0.085
    dx, dy = x - pcx, y - pcy
    dist = math.hypot(dx, dy)
    ang = math.degrees(math.atan2(dy, dx))
    gap = abs(((ang + 90 + 180) % 360) - 180)   # 与正上方的夹角

    if gap > 32 and abs(dist - pr) <= thick / 2:
        return (56, 189, 248, 255)
    if abs(dx) <= thick / 2 and -pr * 1.2 <= dy <= 0:
        return (56, 189, 248, 255)
    return (bg[0], bg[1], bg[2], 255)


def make_icon_png(size: int = 64) -> bytes:
    ss = 2  # 超采样
    rows = bytearray()
    for y in range(size):
        rows.append(0)  # PNG filter type 0
        for x in range(size):
            r = g = b = a = 0
            for sy in range(ss):
                for sx in range(ss):
                    px = _icon_pixel(x + (sx + 0.5) / ss, y + (sy + 0.5) / ss, size)
                    r += px[0] * px[3]
                    g += px[1] * px[3]
                    b += px[2] * px[3]
                    a += px[3]
            n = ss * ss
            if a == 0:
                rows.extend((0, 0, 0, 0))
            else:
                rows.extend((r // a, g // a, b // a, a // n))

    def chunk(tag: bytes, data: bytes) -> bytes:
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(bytes(rows), 9))
    png += chunk(b"IEND", b"")
    return png


# --------------------------------------------------------------------------
# 日历选择器
# --------------------------------------------------------------------------
class CalendarPopup(tk.Toplevel):
    """轻量图形化日历弹窗"""

    WEEK = ["日", "一", "二", "三", "四", "五", "六"]

    def __init__(self, parent, initial: dt.date, on_select, anchor=None):
        super().__init__(parent)
        self.on_select = on_select
        self.anchor = anchor
        self.title("选择日期")
        self.resizable(False, False)
        self.configure(bg=C["card"])
        self.transient(parent)
        self.year, self.month = initial.year, initial.month
        self.selected = initial
        self._build()
        self.update_idletasks()
        self._center_over(parent)

    # -- 构建 --------------------------------------------------------------
    def _build(self):
        head = tk.Frame(self, bg=C["card"])
        head.pack(fill="x", padx=14, pady=(12, 4))

        tk.Button(head, text="◀", relief="flat", bd=0, bg=C["card_hi"],
                  fg=C["text"], activebackground=C["accent"], activeforeground="#04121c",
                  width=3, command=self._prev_year).pack(side="left")
        tk.Button(head, text="‹", relief="flat", bd=0, bg=C["card_hi"],
                  fg=C["text"], activebackground=C["accent"], activeforeground="#04121c",
                  width=3, command=self._prev_month).pack(side="left", padx=(4, 0))

        self.lbl_month = tk.Label(head, text="", bg=C["card"], fg=C["text"],
                                  font=(FONT_UI, 13, "bold"))
        self.lbl_month.pack(side="left", expand=True)

        tk.Button(head, text="›", relief="flat", bd=0, bg=C["card_hi"],
                  fg=C["text"], activebackground=C["accent"], activeforeground="#04121c",
                  width=3, command=self._next_month).pack(side="right")
        tk.Button(head, text="▶", relief="flat", bd=0, bg=C["card_hi"],
                  fg=C["text"], activebackground=C["accent"], activeforeground="#04121c",
                  width=3, command=self._next_year).pack(side="right", padx=(0, 4))

        grid = tk.Frame(self, bg=C["card"])
        grid.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        for col, name in enumerate(self.WEEK):
            tk.Label(grid, text=name, bg=C["card"], fg=C["faint"],
                     font=(FONT_UI, 10), width=4).grid(row=0, column=col, pady=(0, 4))

        self.cells = []
        self._grid = {}
        for r in range(6):
            row = []
            for col in range(7):
                lbl = tk.Label(grid, text="", bg=C["card"], fg=C["text"],
                               font=(FONT_UI, 11), width=4, height=1, bd=0,
                               cursor="hand2")
                lbl.bind("<Button-1>", lambda e, rr=r, cc=col: self._pick(rr, cc))
                lbl.grid(row=r + 1, column=col)
                row.append(lbl)
            self.cells.append(row)

        self._render()

    # -- 逻辑 --------------------------------------------------------------
    def _render(self):
        self.lbl_month.configure(text="%d 年 %d 月" % (self.year, self.month))
        today = dt.date.today()
        self._grid = {}

        days_in_month = calendar.monthrange(self.year, self.month)[1]
        first_col = (dt.date(self.year, self.month, 1).weekday() + 1) % 7   # 周日为第 0 列
        for day in range(1, days_in_month + 1):
            date = dt.date(self.year, self.month, day)
            row, col = divmod(day - 1 + first_col, 7)
            if row < 6:
                self._grid[(row, col)] = date

        for r in range(6):
            for col in range(7):
                lbl = self.cells[r][col]
                date = self._grid.get((r, col))
                if not date:
                    lbl.configure(text="", bg=C["card"])
                    lbl.bind("<Enter>", lambda e: None)
                    lbl.bind("<Leave>", lambda e: None)
                    continue
                lbl.configure(text=str(date.day))
                if date == self.selected:
                    lbl.configure(bg=C["accent"], fg="#04121c", font=(FONT_UI, 11, "bold"))
                elif date == today:
                    lbl.configure(bg=C["card_hi"], fg=C["accent"], font=(FONT_UI, 11, "bold"))
                else:
                    lbl.configure(bg=C["card"], fg=C["text"], font=(FONT_UI, 11))
                lbl.bind("<Enter>", lambda e, w=lbl: w.configure(bg=C["card_hi"]))
                lbl.bind("<Leave>", lambda e, w=lbl, d=date: w.configure(
                    bg=C["accent"] if d == self.selected else
                    (C["card_hi"] if d == today else C["card"])))

    def _pick(self, row, col):
        date = self._grid.get((row, col))
        if not date:
            return
        self.on_select(date)
        self.destroy()

    def _shift(self, months):
        m = self.month - 1 + months
        self.year += m // 12
        self.month = m % 12 + 1
        self._render()

    def _prev_month(self):
        self._shift(-1)

    def _next_month(self):
        self._shift(1)

    def _prev_year(self):
        self.year -= 1
        self._render()

    def _next_year(self):
        self.year += 1
        self._render()

    def _center_over(self, parent):
        if self.anchor is not None:
            x = self.anchor.winfo_rootx()
            y = self.anchor.winfo_rooty() + self.anchor.winfo_height() + 4
        else:
            x = parent.winfo_rootx() + 60
            y = parent.winfo_rooty() + 120
        w, h = self.winfo_width(), self.winfo_height()
        sw = parent.winfo_screenwidth()
        x = min(max(0, x), max(0, sw - w))
        self.geometry("+%d+%d" % (x, y))


# --------------------------------------------------------------------------
# 跨平台电源控制
# --------------------------------------------------------------------------
class PowerControl:
    """关机 / 重启 / 注销的跨平台实现。

    Windows：调用 shutdown.exe，支持 /t 延时与 /a 取消
    macOS  ：用 osascript 驱动 System Events；
              延时任务以独立后台进程运行（关闭本程序后依然生效）
    """

    def __init__(self):
        self.pid = None          # macOS 后台任务进程号

    # -- 内部工具 --------------------------------------------------------
    @staticmethod
    def _run(cmd, timeout=15):
        kwargs = {}
        if IS_WINDOWS:
            kwargs["creationflags"] = CREATE_NO_WINDOW
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  timeout=timeout, **kwargs)
            return proc.returncode, (proc.stdout or "") + (proc.stderr or "")
        except FileNotFoundError as exc:
            return -1, "找不到命令：%s" % exc
        except Exception as exc:
            return -1, str(exc)

    # -- 对外接口 --------------------------------------------------------
    def schedule(self, seconds, action, comment="由定时关机助手安排的自动任务"):
        """安排延时任务，返回 (是否成功, 说明)"""
        seconds = int(max(0, math.ceil(seconds)))
        if IS_MAC:
            script = ACTIONS[action]["mac"]
            kwargs = {"start_new_session": True} if not IS_WINDOWS else {}
            try:
                proc = subprocess.Popen(
                    ["osascript", "-e", "delay %d" % max(1, seconds), "-e", script],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, **kwargs)
            except Exception as exc:
                return False, "无法启动系统级任务：%s" % exc
            self.pid = proc.pid
            return True, "已启动后台系统级任务（PID %d），%d 秒后执行" % (proc.pid, seconds)

        if not ACTIONS[action]["delay"]:           # Windows 的注销不支持 /t
            code, out = self._run([SHUTDOWN_EXE, ACTIONS[action]["win"]])
        else:
            code, out = self._run([SHUTDOWN_EXE, ACTIONS[action]["win"], "/t",
                                   str(seconds), "/c", comment])
        if code == 0:
            return True, "已设置系统级任务：%d 秒后执行" % seconds
        if code == 1190:
            return True, "系统级任务已存在，已保持原有安排"
        return False, "系统级任务设置失败（代码 %s）：%s" % (code, (out or "").strip())

    def abort(self):
        """取消待执行的系统级任务，返回 (是否成功, 说明)"""
        if IS_MAC:
            if self.pid:
                try:
                    os.kill(self.pid, signal.SIGTERM)
                except Exception:
                    pass
                self.pid = None
                return True, "已取消后台系统级任务"
            return False, "系统中没有待执行的后台任务"
        code, out = self._run([SHUTDOWN_EXE, "/a"])
        if code == 0:
            return True, "已取消系统中的全部关机任务"
        return False, "系统中没有待执行的关机任务（代码 %s）" % code

    def run_now(self, action):
        """立即执行动作，返回 (是否成功, 说明)"""
        if IS_MAC:
            code, out = self._run(["osascript", "-e", ACTIONS[action]["mac"]])
            if code == 0:
                return True, "已触发：%s" % ACTIONS[action]["label"]
            return False, "执行失败：%s" % (out or "").strip()
        if ACTIONS[action]["delay"]:
            args = [SHUTDOWN_EXE, ACTIONS[action]["win"], "/t", "0", "/c", "定时关机助手触发"]
        else:
            args = [SHUTDOWN_EXE, ACTIONS[action]["win"]]
        code, out = self._run(args)
        if code == 0:
            return True, "已触发：%s" % ACTIONS[action]["label"]
        return False, "执行失败（代码 %s）：%s" % (code, (out or "").strip())


def open_path(path: str):
    """用系统默认程序打开目录/文件"""
    try:
        if IS_WINDOWS:
            os.startfile(path)
        elif IS_MAC:
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# --------------------------------------------------------------------------
# 健康提醒：图标绘制 + 置顶提醒弹窗
# --------------------------------------------------------------------------
def draw_symbol(canvas, kind: str, color: str):
    """在 Canvas 上绘制提醒图标（water / kegel / walk）"""
    canvas.delete("all")
    if kind == "water":
        canvas.create_polygon(23, 3, 35, 21, 11, 21, fill=color, outline="")
        canvas.create_oval(11, 17, 35, 43, fill=color, outline="")
        canvas.create_oval(17, 26, 24, 36, fill="#ffffff", outline="")
    elif kind == "kegel":
        canvas.create_oval(9, 9, 37, 37, outline=color, width=3)
        canvas.create_polygon(23, 13, 30, 21, 16, 21, fill=color, outline="")
        canvas.create_polygon(23, 33, 30, 25, 16, 25, fill=color, outline="")
    else:  # walk
        canvas.create_oval(19, 3, 29, 13, fill=color, outline="")
        canvas.create_line(24, 14, 24, 30, fill=color, width=3)
        canvas.create_line(24, 20, 14, 26, fill=color, width=3)
        canvas.create_line(24, 20, 34, 24, fill=color, width=3)
        canvas.create_line(24, 30, 16, 42, fill=color, width=3)
        canvas.create_line(24, 30, 32, 41, fill=color, width=3)


class ReminderToast(tk.Toplevel):
    """置顶的健康提醒弹窗：喝水 / 提肛 / 走动"""

    def __init__(self, parent, kind, title, message, color, on_ok=None,
                 on_snooze=None, snooze_text="稍后 5 分钟"):
        super().__init__(parent)
        self.on_ok = on_ok
        self.on_snooze = on_snooze
        self.title(title)
        self.configure(bg=C["border"])
        self.resizable(False, False)
        try:
            self.attributes("-topmost", True)
            self.iconphoto(True, parent._icon)
        except Exception:
            pass

        outer = tk.Frame(self, bg=C["border"])
        outer.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg=C["card"])
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        head = tk.Frame(inner, bg=C["card"])
        head.pack(fill="x", padx=18, pady=(16, 0))
        icon = tk.Canvas(head, width=46, height=46, bg=C["card"], highlightthickness=0)
        icon.pack(side="left")
        draw_symbol(icon, kind, color)
        tk.Label(head, text=title, bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 14, "bold")).pack(side="left", padx=(12, 0))

        tk.Label(inner, text=message, bg=C["card"], fg=C["muted"], font=(FONT_UI, 10),
                 justify="left", wraplength=380).pack(anchor="w", padx=18, pady=(12, 0))

        btns = tk.Frame(inner, bg=C["card"])
        btns.pack(fill="x", padx=18, pady=(16, 16))
        ttk.Button(btns, text="好的", style="Accent.TButton",
                   command=self._ok).pack(side="right")
        if on_snooze:
            ttk.Button(btns, text=snooze_text, style="TButton",
                       command=self._snooze).pack(side="right", padx=8)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        self._center_over(parent)

    def _center_over(self, parent):
        w, h = self.winfo_width(), self.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + 90
        self.geometry("+%d+%d" % (max(0, x), max(0, y)))

    def _ok(self):
        if self.on_ok:
            try:
                self.on_ok()
            except Exception:
                pass
        self.destroy()

    def _snooze(self):
        if self.on_snooze:
            try:
                self.on_snooze()
            except Exception:
                pass
        self.destroy()


# --------------------------------------------------------------------------
# 主程序
# --------------------------------------------------------------------------
class ShutdownApp(tk.Tk):
    def __init__(self):
        if not (IS_WINDOWS or IS_MAC):
            messagebox.showwarning(APP_TITLE, "本工具目前支持 Windows 与 macOS。")
            sys.exit(0)

        if IS_WINDOWS:
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                try:
                    ctypes.windll.user32.SetProcessDPIAware()
                except Exception:
                    pass

        super().__init__()
        self.title("%s  v%s" % (APP_TITLE, APP_VERSION))
        self.configure(bg=C["bg"])
        self.minsize(760, 680)
        self._center(880, 790)

        global FONT_UI, FONT_MONO
        FONT_UI = pick_font(["Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "SimHei",
                             "PingFang SC", "Hiragino Sans GB", "Heiti SC", "Helvetica Neue"])
        FONT_MONO = pick_font(["Consolas", "Cascadia Mono", "Courier New",
                               "SF Mono", "Menlo", "Monaco"], "TkFixedFont")

        self._icon = tk.PhotoImage(data=base64.b64encode(make_icon_png(64)))
        self._icon_small = tk.PhotoImage(data=base64.b64encode(make_icon_png(48)))
        try:
            self.iconphoto(True, self._icon)
        except Exception:
            pass

        self.plan = None
        self._firing = False
        self.log_lines = []
        self._toasts = {}
        self.power = PowerControl()
        self.water_day = dt.date.today().isoformat()
        self.walk_day = self.water_day
        self.kegel_paused_phase = "squeeze"
        self.kegel_remaining = 0.0

        self._load_state()
        self._init_styles()
        self._build_ui()
        self._restore_plan()
        self._restore_native_task()
        self._tick()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # 状态持久化
    # ------------------------------------------------------------------
    def _load_state(self):
        data = {}
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}

        settings = data.get("settings", {})
        action = settings.get("action", "关机")
        if action not in ACTIONS:                 # 兼容其他系统保存的状态
            action = list(ACTIONS.keys())[0]
        self.var_action = tk.StringVar(value=action)
        self.var_native = tk.BooleanVar(value=bool(settings.get("native", True)))

        plan = data.get("plan")
        self.plan = plan if isinstance(plan, dict) else None
        self._pending_settings = settings
        self._pending_native_pid = data.get("native_pid")

        # 健康提醒（喝水 / 提肛 / 走动）
        health = data.get("health")
        health = health if isinstance(health, dict) else {}
        today = dt.date.today().isoformat()
        water = health.get("water", {}) if isinstance(health.get("water"), dict) else {}
        walk = health.get("walk", {}) if isinstance(health.get("walk"), dict) else {}
        kegel = health.get("kegel", {}) if isinstance(health.get("kegel"), dict) else {}
        self.health_cfg = {
            "water": {
                "interval": int(water.get("interval", 45)),
                "target": int(water.get("target", 8)),
                "count": int(water.get("count", 0)) if water.get("day") == today else 0,
                "day": today,
                "running": bool(water.get("running", False)),
                "next": float(water.get("next", 0) or 0),
            },
            "walk": {
                "interval": int(walk.get("interval", 60)),
                "duration": int(walk.get("duration", 2)),
                "count": int(walk.get("count", 0)) if walk.get("day") == today else 0,
                "total": int(walk.get("total", 0)) if walk.get("day") == today else 0,
                "day": today,
                "running": bool(walk.get("running", False)),
                "next": float(walk.get("next", 0) or 0),
            },
            "kegel": {
                "squeeze": int(kegel.get("squeeze", 5)),
                "relax": int(kegel.get("relax", 5)),
                "rounds": int(kegel.get("rounds", 10)),
                "done": int(kegel.get("done", 0)),
                "sound": bool(kegel.get("sound", True)),
            },
        }

    def _save_state(self):
        data = {
            "settings": {
                "action": self.var_action.get(),
                "native": bool(self.var_native.get()),
            },
            "plan": self.plan,
            "health": self._health_snapshot(),
            "native_pid": getattr(getattr(self, "power", None), "pid", None),
            "updated": dt.datetime.now().isoformat(timespec="seconds"),
        }
        try:
            with open(STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            self._log("状态保存失败：%s" % exc, "error")

    def _health_snapshot(self):
        """把三个提醒器的当前状态打包，供持久化"""
        cfg = getattr(self, "health_cfg", None)
        if cfg is None:
            return {}
        snap = {k: dict(v) for k, v in cfg.items()}
        if getattr(self, "var_water_interval", None) is not None:
            snap["water"].update({
                "interval": self.var_water_interval.get(),
                "target": self.var_water_target.get(),
                "count": self.water_count,
                "running": self.water_running,
                "next": self.water_next,
                "day": dt.date.today().isoformat(),
            })
        if getattr(self, "var_walk_interval", None) is not None:
            snap["walk"].update({
                "interval": self.var_walk_interval.get(),
                "duration": self.var_walk_duration.get(),
                "count": self.walk_count,
                "total": self.walk_total,
                "running": self.walk_running,
                "next": self.walk_next,
                "day": dt.date.today().isoformat(),
            })
        if getattr(self, "var_kegel_squeeze", None) is not None:
            snap["kegel"].update({
                "squeeze": self.var_kegel_squeeze.get(),
                "relax": self.var_kegel_relax.get(),
                "rounds": self.var_kegel_rounds.get(),
                "done": self.kegel_done,
                "sound": bool(self.var_kegel_sound.get()),
            })
        return snap

    def _restore_plan(self):
        if not self.plan:
            return
        target = float(self.plan.get("target", 0))
        if target <= time.time():
            self.plan = None
            self._save_state()
            return
        self._log("已恢复上次的计划：%s（剩余 %s）" %
                  (self.plan.get("desc", "定时任务"), fmt_natural(target - time.time())))
        self._refresh_banner()
        self._update_status()

    def _restore_native_task(self):
        """恢复上次的系统级任务句柄（macOS 后台进程），以便仍可取消"""
        pid = getattr(self, "_pending_native_pid", None)
        if IS_MAC and pid and os.name != "nt":      # 仅 POSIX 可安全探测进程
            try:
                os.kill(int(pid), 0)          # 探测进程是否仍存在
                self.power.pid = int(pid)
                self._log("检测到仍在后台运行的系统级任务（PID %d）" % pid)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # 样式
    # ------------------------------------------------------------------
    def _init_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background=C["bg"])
        style.configure("TLabel", background=C["bg"], foreground=C["text"], font=(FONT_UI, 10))
        style.configure("Card.TFrame", background=C["card"])
        style.configure("Muted.TLabel", background=C["card"], foreground=C["muted"], font=(FONT_UI, 9))
        style.configure("Text.TLabel", background=C["card"], foreground=C["text"], font=(FONT_UI, 10))
        style.configure("H1.TLabel", background=C["card"], foreground=C["text"], font=(FONT_UI, 15, "bold"))
        style.configure("H2.TLabel", background=C["card"], foreground=C["muted"], font=(FONT_UI, 9))

        style.configure("TButton", background=C["card_hi"], foreground=C["text"],
                        borderwidth=0, focuscolor=C["card_hi"], padding=(14, 7),
                        font=(FONT_UI, 10))
        style.map("TButton",
                  background=[("active", C["accent_dim"]), ("pressed", C["accent_dim"])],
                  foreground=[("active", "#f0f9ff")])

        style.configure("Accent.TButton", background=C["accent"], foreground="#04121c",
                        borderwidth=0, focuscolor=C["accent"], padding=(20, 9),
                        font=(FONT_UI, 10, "bold"))
        style.map("Accent.TButton",
                  background=[("active", "#7dd3fc"), ("pressed", "#0ea5e9")],
                  foreground=[("active", "#04121c")])

        style.configure("Danger.TButton", background="#3b1d22", foreground=C["danger"],
                        borderwidth=0, focuscolor="#3b1d22", padding=(14, 6), font=(FONT_UI, 9))
        style.map("Danger.TButton",
                  background=[("active", "#7f1d1d"), ("pressed", "#991b1b")],
                  foreground=[("active", "#fee2e2")])

        style.configure("TNotebook", background=C["bg"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", background=C["panel"], foreground=C["muted"],
                        padding=(22, 9), borderwidth=0, font=(FONT_UI, 10))
        style.map("TNotebook.Tab",
                  background=[("selected", C["card"]), ("active", C["card_hi"])],
                  foreground=[("selected", C["accent"]), ("active", C["text"])])

        style.configure("Sub.TNotebook", background=C["card"], borderwidth=0)
        style.configure("Sub.TNotebook.Tab", background=C["panel"], foreground=C["muted"],
                        padding=(16, 6), borderwidth=0, font=(FONT_UI, 9))
        style.map("Sub.TNotebook.Tab",
                  background=[("selected", C["card_hi"]), ("active", C["border"])],
                  foreground=[("selected", C["accent"]), ("active", C["text"])])

        style.configure("TProgressbar", background=C["accent"], troughcolor=C["card_hi"],
                        borderwidth=0, thickness=8)
        style.configure("Vertical.TScrollbar", background=C["card_hi"],
                        troughcolor=C["panel"], bordercolor=C["panel"],
                        arrowcolor=C["muted"], darkcolor=C["card_hi"],
                        lightcolor=C["card_hi"], borderwidth=0, arrowsize=12)
        style.configure("TCheckbutton", background=C["card"], foreground=C["muted"], font=(FONT_UI, 9))
        style.map("TCheckbutton", background=[("active", C["card"])],
                  foreground=[("active", C["text"])])
        style.configure("TCombobox", fieldbackground=C["card_hi"], background=C["card_hi"],
                        foreground=C["text"], borderwidth=0, arrowcolor=C["text"],
                        padding=(8, 5), font=(FONT_UI, 10))
        style.map("TCombobox", fieldbackground=[("readonly", C["card_hi"])],
                  foreground=[("readonly", C["text"])], selectbackground=[("readonly", C["card_hi"])],
                  selectforeground=[("readonly", C["text"])])
        self.option_add("*TCombobox*Listbox.background", C["card_hi"])
        self.option_add("*TCombobox*Listbox.foreground", C["text"])
        self.option_add("*TCombobox*Listbox.selectBackground", C["accent"])
        self.option_add("*TCombobox*Listbox.selectForeground", "#04121c")

        style.configure("TSpinbox", fieldbackground=C["card_hi"], background=C["card_hi"],
                        foreground=C["text"], borderwidth=0, arrowcolor=C["text"],
                        padding=(6, 4), font=(FONT_MONO, 13))
        style.map("TSpinbox", fieldbackground=[("readonly", C["card_hi"])],
                  foreground=[("readonly", C["text"])])

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------
    def _card(self, parent, **kw):
        """带 1px 描边的卡片容器，返回 (外框, 内容框)"""
        outer = tk.Frame(parent, bg=C["border"], bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=C["card"], bd=0, highlightthickness=0, **kw)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return outer, inner

    def _chip(self, parent, text, command, width=None):
        btn = tk.Button(parent, text=text, command=command, relief="flat", bd=0,
                        bg=C["card_hi"], fg=C["text"], activebackground=C["accent"],
                        activeforeground="#04121c", font=(FONT_UI, 9),
                        cursor="hand2", padx=12, pady=6,
                        highlightthickness=0)
        btn.bind("<Enter>", lambda e: btn.configure(bg=C["border"]))
        btn.bind("<Leave>", lambda e: btn.configure(bg=C["card_hi"]))
        if width:
            btn.configure(width=width)
        return btn

    def _build_ui(self):
        self._build_header()

        self.body = tk.Frame(self, bg=C["bg"])
        self.body.pack(fill="both", expand=True, padx=16, pady=(10, 12))

        self._build_banner()
        # 底部区域先占位，选项卡再吃掉剩余空间，避免状态栏被挤出窗口
        self._build_options()
        self._build_statusbar()
        self._build_tabs()

    # -- 顶部标题栏 ------------------------------------------------------
    def _build_header(self):
        head = tk.Frame(self, bg=C["bg"])
        head.pack(fill="x", padx=16, pady=(14, 0))

        logo = tk.Label(head, image=self._icon_small, bg=C["bg"], bd=0)
        logo.image = self._icon_small
        logo.pack(side="left")

        box = tk.Frame(head, bg=C["bg"])
        box.pack(side="left", padx=(12, 0))
        tk.Label(box, text=APP_TITLE, bg=C["bg"], fg=C["text"],
                 font=(FONT_UI, 15, "bold")).pack(anchor="w")
        tk.Label(box, text="关机定时 · 喝水 · 提肛 · 走动 · %s" % PLATFORM_LABEL,
                 bg=C["bg"], fg=C["muted"], font=(FONT_UI, 9)).pack(anchor="w")

        clock = tk.Frame(head, bg=C["bg"])
        clock.pack(side="right")
        self.var_clock = tk.StringVar(value="--:--:--")
        self.var_date = tk.StringVar(value="")
        tk.Label(clock, textvariable=self.var_clock, bg=C["bg"], fg=C["accent"],
                 font=(FONT_MONO, 18, "bold")).pack(anchor="e")
        tk.Label(clock, textvariable=self.var_date, bg=C["bg"], fg=C["muted"],
                 font=(FONT_UI, 9)).pack(anchor="e")

    # -- 运行中横幅 ------------------------------------------------------
    def _build_banner(self):
        self.banner = tk.Frame(self.body, bg=C["card"])
        strip = tk.Frame(self.banner, bg=C["accent"], width=4)
        strip.pack(side="left", fill="y")

        inner = tk.Frame(self.banner, bg=C["card"])
        inner.pack(side="left", fill="both", expand=True, padx=14, pady=10)

        self.var_plan_desc = tk.StringVar(value="")
        self.var_remain = tk.StringVar(value="")
        tk.Label(inner, text="正在执行的任务", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 9)).pack(anchor="w")
        tk.Label(inner, textvariable=self.var_plan_desc, bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 11, "bold")).pack(anchor="w", pady=(2, 0))
        tk.Label(inner, textvariable=self.var_remain, bg=C["card"], fg=C["accent"],
                 font=(FONT_MONO, 26, "bold")).pack(anchor="w")

        right = tk.Frame(self.banner, bg=C["card"])
        right.pack(side="right", fill="y", padx=14)
        self.progress = ttk.Progressbar(right, maximum=1000, value=0, length=220,
                                        style="TProgressbar")
        self.progress.pack(fill="x", pady=(14, 6))
        ttk.Button(right, text="取消该任务", style="Danger.TButton",
                   command=self._cancel_plan).pack(anchor="e")

        self.banner.pack_forget()

    # -- 选项区 ----------------------------------------------------------
    def _build_options(self):
        bar = tk.Frame(self.body, bg=C["bg"])
        bar.pack(side="bottom", fill="x", pady=(10, 0))

        outer, inner = self._card(bar)
        outer.pack(fill="x")
        inner.configure(pady=8)

        tk.Label(inner, text="执行动作", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 9)).pack(side="left", padx=(14, 6))
        self.cmb_action = ttk.Combobox(inner, textvariable=self.var_action,
                                      values=list(ACTIONS.keys()), state="readonly",
                                      width=12, font=(FONT_UI, 10))
        self.cmb_action.pack(side="left")
        self.cmb_action.bind("<<ComboboxSelected>>", lambda e: self._save_state())

        chk_text = ("同时设置系统级任务（后台脚本，关闭本程序后仍会执行）" if IS_MAC
                    else "同时设置系统级任务（关闭本程序后仍会执行）")
        chk = ttk.Checkbutton(inner, text=chk_text,
                              variable=self.var_native)
        chk.pack(side="left", padx=18)
        self.var_native.trace_add("write", lambda *a: self._save_state())

        tk.Label(inner, text="日志目录：%s" % APP_NAME, bg=C["card"], fg=C["faint"],
                 font=(FONT_UI, 8)).pack(side="right", padx=14)

    # -- 状态栏 ----------------------------------------------------------
    def _build_statusbar(self):
        bar = tk.Frame(self.body, bg=C["bg"])
        bar.pack(side="bottom", fill="x", pady=(8, 0))

        self.var_status = tk.StringVar(value="状态：空闲")
        tk.Label(bar, textvariable=self.var_status, bg=C["bg"], fg=C["muted"],
                 font=(FONT_UI, 9)).pack(side="left")

        ttk.Button(bar, text="取消全部关机任务", style="Danger.TButton",
                   command=self._abort_all).pack(side="right")

    # -- 选项卡 ----------------------------------------------------------
    def _build_tabs(self):
        nb = ttk.Notebook(self.body)
        nb.pack(fill="both", expand=True, pady=(12, 0))
        self.nb = nb

        self.tab_countdown = tk.Frame(nb, bg=C["card"])
        self.tab_schedule = tk.Frame(nb, bg=C["card"])
        self.tab_health = tk.Frame(nb, bg=C["card"])
        self.tab_log = tk.Frame(nb, bg=C["card"])
        nb.add(self.tab_countdown, text="  倒计时关机  ")
        nb.add(self.tab_schedule, text="  定时关机  ")
        nb.add(self.tab_health, text="  健康提醒  ")
        nb.add(self.tab_log, text="  运行日志  ")

        self._build_countdown_tab(self.tab_countdown)
        self._build_schedule_tab(self.tab_schedule)
        self._build_health_tab(self.tab_health)
        self._build_log_tab(self.tab_log)

    # -- 健康提醒页 ------------------------------------------------------
    def _build_health_tab(self, parent):
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(wrap, text="健康提醒", bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 14, "bold")).pack(anchor="w")
        tk.Label(wrap, text="喝水、盆底肌训练、起身活动——久坐办公的三个小助手。",
                 bg=C["card"], fg=C["muted"], font=(FONT_UI, 9)).pack(anchor="w", pady=(2, 0))

        sub = ttk.Notebook(wrap, style="Sub.TNotebook")
        sub.pack(fill="both", expand=True, pady=(10, 0))
        self.sub_nb = sub

        f_water = tk.Frame(sub, bg=C["card"])
        f_kegel = tk.Frame(sub, bg=C["card"])
        f_walk = tk.Frame(sub, bg=C["card"])
        sub.add(f_water, text="  喝水提醒  ")
        sub.add(f_kegel, text="  提肛助手  ")
        sub.add(f_walk, text="  走动提醒  ")

        self._build_water_tab(f_water)
        self._build_kegel_tab(f_kegel)
        self._build_walk_tab(f_walk)

    # -- 喝水提醒 --------------------------------------------------------
    def _build_water_tab(self, parent):
        cfg = self.health_cfg["water"]
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        row = tk.Frame(wrap, bg=C["card"])
        row.pack(fill="x")
        tk.Label(row, text="提醒间隔", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(0, 10))
        self.var_water_interval = tk.IntVar(value=cfg["interval"])
        sp = ttk.Spinbox(row, from_=5, to=240, width=5, textvariable=self.var_water_interval,
                         font=(FONT_MONO, 13), justify="center", style="TSpinbox")
        sp.pack(side="left")
        tk.Label(row, text="分钟", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(6, 14))
        for text, mins in (("30 分钟", 30), ("45 分钟", 45), ("1 小时", 60), ("2 小时", 120)):
            self._chip(row, text, lambda v=mins: self.var_water_interval.set(v)).pack(
                side="left", padx=(0, 8))
        tk.Label(row, text="杯", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="right", padx=(0, 6))
        self.var_water_target = tk.IntVar(value=cfg["target"])
        ttk.Spinbox(row, from_=1, to=30, width=4, textvariable=self.var_water_target,
                    font=(FONT_MONO, 13), justify="center",
                    style="TSpinbox").pack(side="right")
        tk.Label(row, text="目标", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="right", padx=(14, 10))

        self.var_water_rem = tk.StringVar(value="--:--:--")
        tk.Label(wrap, textvariable=self.var_water_rem, bg=C["card"], fg=C["accent"],
                 font=(FONT_MONO, 30, "bold")).pack(anchor="w", pady=(14, 0))
        self.lbl_water_state = tk.Label(wrap, text="", bg=C["card"], fg=C["muted"], font=(FONT_UI, 9))
        self.lbl_water_state.pack(anchor="w", pady=(2, 0))

        self.water_progress = ttk.Progressbar(wrap, maximum=1000, style="TProgressbar")
        self.water_progress.pack(fill="x", pady=(12, 0))

        stat = tk.Frame(wrap, bg=C["card"])
        stat.pack(fill="x", pady=(12, 0))
        self.var_water_stat = tk.StringVar(value="今日 0 / 8 杯")
        tk.Label(stat, textvariable=self.var_water_stat, bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 10, "bold")).pack(side="left")
        self.water_stat_bar = ttk.Progressbar(stat, maximum=1000, length=220,
                                               style="TProgressbar")
        self.water_stat_bar.pack(side="left", padx=12)
        tk.Label(stat, text="每次完成后点「好的」即可计数", bg=C["card"], fg=C["faint"],
                 font=(FONT_UI, 8)).pack(side="right")

        btns = tk.Frame(wrap, bg=C["card"])
        btns.pack(fill="x", pady=(16, 0))
        ttk.Button(btns, text="开始提醒", style="Accent.TButton",
                   command=self._water_start).pack(side="left")
        ttk.Button(btns, text="暂停", style="TButton",
                   command=self._water_pause).pack(side="left", padx=8)
        ttk.Button(btns, text="立即提醒", style="TButton",
                   command=self._water_fire_now).pack(side="left")
        ttk.Button(btns, text="重置", style="TButton",
                   command=self._water_reset).pack(side="left", padx=8)

        tk.Label(wrap, text="提示：少量多次补水更利于吸收；咖啡与浓茶会加重脱水，"
                            "建议以白水或淡茶为主。",
                 bg=C["card"], fg=C["faint"], font=(FONT_UI, 8),
                 justify="left").pack(anchor="w", pady=(16, 0))

        self.water_count = cfg["count"]
        self.water_running = cfg["running"]
        self.water_next = cfg["next"]
        if self.water_running and self.water_next <= time.time():
            self.water_next = time.time() + cfg["interval"] * 60
        self.var_water_interval.trace_add("write", lambda *a: self._save_state())
        self.var_water_target.trace_add("write", lambda *a: self._save_state())
        self._refresh_water()

    # -- 提肛助手 --------------------------------------------------------
    def _build_kegel_tab(self, parent):
        cfg = self.health_cfg["kegel"]
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        row = tk.Frame(wrap, bg=C["card"])
        row.pack(fill="x")
        self.var_kegel_squeeze = tk.IntVar(value=cfg["squeeze"])
        self.var_kegel_relax = tk.IntVar(value=cfg["relax"])
        self.var_kegel_rounds = tk.IntVar(value=cfg["rounds"])
        for text, var, rng in (("收缩", self.var_kegel_squeeze, (1, 30)),
                               ("放松", self.var_kegel_relax, (1, 60)),
                               ("组数", self.var_kegel_rounds, (1, 50))):
            tk.Label(row, text=text, bg=C["card"], fg=C["muted"],
                     font=(FONT_UI, 10)).pack(side="left", padx=(0, 8))
            ttk.Spinbox(row, from_=rng[0], to=rng[1], width=4, textvariable=var,
                        font=(FONT_MONO, 13), justify="center",
                        style="TSpinbox").pack(side="left", padx=(0, 8))
            tk.Label(row, text="秒" if var is not self.var_kegel_rounds else "组",
                     bg=C["card"], fg=C["muted"], font=(FONT_UI, 10)).pack(
                side="left", padx=(0, 16))

        self.var_kegel_phase = tk.StringVar(value="准备开始")
        self.var_kegel_phase_label = tk.Label(wrap, textvariable=self.var_kegel_phase,
                                              bg=C["card"], fg=C["accent"],
                                              font=(FONT_UI, 13, "bold"))
        self.var_kegel_phase_label.pack(anchor="w", pady=(14, 0))
        self.var_kegel_timer = tk.StringVar(value="--")
        self.lbl_kegel_timer = tk.Label(wrap, textvariable=self.var_kegel_timer, bg=C["card"],
                                        fg=C["text"], font=(FONT_MONO, 40, "bold"))
        self.lbl_kegel_timer.pack(anchor="w")
        self.kegel_progress = ttk.Progressbar(wrap, maximum=1000, style="TProgressbar")
        self.kegel_progress.pack(fill="x", pady=(10, 0))
        self.var_kegel_stat = tk.StringVar(value="已完成 0 组训练")
        tk.Label(wrap, textvariable=self.var_kegel_stat, bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 9)).pack(anchor="w", pady=(4, 0))

        btns = tk.Frame(wrap, bg=C["card"])
        btns.pack(fill="x", pady=(16, 0))
        ttk.Button(btns, text="开始训练", style="Accent.TButton",
                   command=self._kegel_start).pack(side="left")
        ttk.Button(btns, text="暂停", style="TButton",
                   command=self._kegel_pause).pack(side="left", padx=8)
        ttk.Button(btns, text="结束", style="TButton",
                   command=self._kegel_stop).pack(side="left")
        ttk.Button(btns, text="重置", style="TButton",
                   command=self._kegel_reset).pack(side="left", padx=8)
        self.var_kegel_sound = tk.BooleanVar(value=cfg["sound"])
        ttk.Checkbutton(btns, text="阶段提示音", variable=self.var_kegel_sound).pack(
            side="left", padx=20)
        self.var_kegel_sound.trace_add("write", lambda *a: self._save_state())

        tk.Label(wrap, text="提示：收缩时缓慢吸气、放松时呼气，力度以能感受到收缩即可；"
                            "出现坠胀或疼痛请立即停止，孕期或产后请遵医嘱。",
                 bg=C["card"], fg=C["faint"], font=(FONT_UI, 8),
                 justify="left").pack(anchor="w", pady=(16, 0))

        self.kegel_phase = "idle"          # idle / squeeze / relax / paused / done
        self.kegel_phase_end = 0.0
        self.kegel_phase_total = 0.0
        self.kegel_round = 0
        self.kegel_done = cfg["done"]
        for var in (self.var_kegel_squeeze, self.var_kegel_relax, self.var_kegel_rounds):
            var.trace_add("write", lambda *a: self._save_state())

    # -- 走动提醒 --------------------------------------------------------
    def _build_walk_tab(self, parent):
        cfg = self.health_cfg["walk"]
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        row = tk.Frame(wrap, bg=C["card"])
        row.pack(fill="x")
        self.var_walk_interval = tk.IntVar(value=cfg["interval"])
        self.var_walk_duration = tk.IntVar(value=cfg["duration"])
        tk.Label(row, text="久坐提醒", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(0, 10))
        ttk.Spinbox(row, from_=10, to=240, width=5, textvariable=self.var_walk_interval,
                    font=(FONT_MONO, 13), justify="center", style="TSpinbox").pack(side="left")
        tk.Label(row, text="分钟", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(6, 14))
        for text, mins in (("30 分钟", 30), ("45 分钟", 45), ("1 小时", 60), ("2 小时", 120)):
            self._chip(row, text, lambda v=mins: self.var_walk_interval.set(v)).pack(
                side="left", padx=(0, 8))

        row2 = tk.Frame(wrap, bg=C["card"])
        row2.pack(fill="x", pady=(10, 0))
        tk.Label(row2, text="活动时长", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(0, 10))
        ttk.Spinbox(row2, from_=1, to=30, width=5, textvariable=self.var_walk_duration,
                    font=(FONT_MONO, 13), justify="center", style="TSpinbox").pack(side="left")
        tk.Label(row2, text="分钟", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(6, 14))
        for text, mins in (("1 分钟", 1), ("2 分钟", 2), ("3 分钟", 3), ("5 分钟", 5)):
            self._chip(row2, text, lambda v=mins: self.var_walk_duration.set(v)).pack(
                side="left", padx=(0, 8))

        self.var_walk_rem = tk.StringVar(value="--:--:--")
        self.lbl_walk_rem = tk.Label(wrap, textvariable=self.var_walk_rem, bg=C["card"],
                                     fg=C["accent"], font=(FONT_MONO, 30, "bold"))
        self.lbl_walk_rem.pack(anchor="w", pady=(14, 0))
        self.lbl_walk_state = tk.Label(wrap, text="", bg=C["card"], fg=C["muted"], font=(FONT_UI, 9))
        self.lbl_walk_state.pack(anchor="w", pady=(2, 0))

        self.walk_progress = ttk.Progressbar(wrap, maximum=1000, style="TProgressbar")
        self.walk_progress.pack(fill="x", pady=(12, 0))

        self.var_walk_stat = tk.StringVar(value="今日活动 0 次 · 累计 0 分钟")
        tk.Label(wrap, textvariable=self.var_walk_stat, bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 10, "bold")).pack(anchor="w", pady=(12, 0))

        btns = tk.Frame(wrap, bg=C["card"])
        btns.pack(fill="x", pady=(16, 0))
        ttk.Button(btns, text="开始提醒", style="Accent.TButton",
                   command=self._walk_start).pack(side="left")
        ttk.Button(btns, text="暂停", style="TButton",
                   command=self._walk_pause).pack(side="left", padx=8)
        ttk.Button(btns, text="立即活动", style="TButton",
                   command=self._walk_fire_now).pack(side="left")
        ttk.Button(btns, text="结束活动", style="TButton",
                   command=self._walk_finish).pack(side="left", padx=8)

        tk.Label(wrap, text="提示：久坐易引发腰颈不适与下肢循环问题；"
                            "建议起身接杯水、走动 2 分钟并做肩颈伸展。",
                 bg=C["card"], fg=C["faint"], font=(FONT_UI, 8),
                 justify="left").pack(anchor="w", pady=(16, 0))

        self.walk_count = cfg["count"]
        self.walk_total = cfg["total"]
        self.walk_running = cfg["running"]
        self.walk_next = cfg["next"]
        self.walk_active_until = 0.0
        if self.walk_running and self.walk_next <= time.time():
            self.walk_next = time.time() + cfg["interval"] * 60
        self.var_walk_interval.trace_add("write", lambda *a: self._save_state())
        self.var_walk_duration.trace_add("write", lambda *a: self._save_state())
        self._refresh_walk()

    # -- 倒计时页 --------------------------------------------------------
    def _build_countdown_tab(self, parent):
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        tk.Label(wrap, text="倒计时关机", bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 14, "bold")).pack(anchor="w")
        tk.Label(wrap, text="选择常用的快捷时长，或自定义小时 / 分钟 / 秒。",
                 bg=C["card"], fg=C["muted"], font=(FONT_UI, 9)).pack(anchor="w", pady=(2, 0))

        chips = tk.Frame(wrap, bg=C["card"])
        chips.pack(fill="x", pady=(10, 0))
        presets = [("5 分钟", 0, 5, 0), ("10 分钟", 0, 10, 0), ("15 分钟", 0, 15, 0),
                   ("30 分钟", 0, 30, 0), ("45 分钟", 0, 45, 0), ("1 小时", 1, 0, 0),
                   ("2 小时", 2, 0, 0), ("3 小时", 3, 0, 0)]
        for text, h, m, s in presets:
            self._chip(chips, text, lambda h=h, m=m, s=s: self._set_countdown(h, m, s)).pack(
                side="left", padx=(0, 8))

        custom = tk.Frame(wrap, bg=C["card"])
        custom.pack(fill="x", pady=(12, 0))

        self.var_h = tk.StringVar(value="0")
        self.var_m = tk.StringVar(value="30")
        self.var_s = tk.StringVar(value="0")

        tk.Label(custom, text="自定义", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 9)).pack(side="left", padx=(0, 14))
        for var, text, to in ((self.var_h, "时", 23), (self.var_m, "分", 59), (self.var_s, "秒", 59)):
            tk.Label(custom, text=text, bg=C["card"], fg=C["muted"],
                     font=(FONT_UI, 10)).pack(side="left", padx=(0, 4))
            sp = ttk.Spinbox(custom, from_=0, to=to, width=4, textvariable=var,
                             font=(FONT_MONO, 13), justify="center", style="TSpinbox")
            sp.pack(side="left", padx=(0, 10))
            sp.bind("<KeyRelease>", lambda e: self._refresh_countdown_preview())
            sp.bind("<FocusOut>", lambda e: self._refresh_countdown_preview())

        self.var_countdown_preview = tk.StringVar(value="")
        tk.Label(wrap, textvariable=self.var_countdown_preview, bg=C["card"], fg=C["accent"],
                 font=(FONT_MONO, 24, "bold")).pack(anchor="w", pady=(10, 0))
        self.lbl_countdown_when = tk.Label(wrap, text="", bg=C["card"], fg=C["muted"],
                                           font=(FONT_UI, 9))
        self.lbl_countdown_when.pack(anchor="w", pady=(2, 0))

        ttk.Button(wrap, text="▶  开始倒计时关机", style="Accent.TButton",
                   command=self._start_countdown).pack(anchor="w", pady=(10, 0))

        tk.Label(wrap, text="提示：底部可切换关机 / 重启 / %s；勾选「同时设置系统级任务」"
                            "后，关闭本程序也能按时执行。" %
                            list(ACTIONS.keys())[-1],
                 bg=C["card"], fg=C["faint"], font=(FONT_UI, 8),
                 justify="left").pack(anchor="w", pady=(10, 0))

        self._refresh_countdown_preview()

    def _set_countdown(self, h, m, s):
        self.var_h.set(str(h))
        self.var_m.set(str(m))
        self.var_s.set(str(s))
        self._refresh_countdown_preview()

    def _countdown_seconds(self):
        try:
            h = max(0, int(self.var_h.get() or 0))
            m = max(0, min(59, int(self.var_m.get() or 0)))
            s = max(0, min(59, int(self.var_s.get() or 0)))
        except ValueError:
            return 0
        return h * 3600 + m * 60 + s

    def _refresh_countdown_preview(self):
        total = self._countdown_seconds()
        self.var_countdown_preview.set(fmt_clock(total))
        if total > 0:
            self.lbl_countdown_when.configure(
                text="将从 %s 开始，倒计时 %s 后自动%s" % (
                    dt.datetime.now().strftime("%H:%M:%S"), fmt_natural(total),
                    ACTIONS[self.var_action.get()]["label"]))
        else:
            self.lbl_countdown_when.configure(text="请设置大于 0 的时长")

    # -- 定时页 ----------------------------------------------------------
    def _build_schedule_tab(self, parent):
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        tk.Label(wrap, text="定时关机", bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 14, "bold")).pack(anchor="w")
        tk.Label(wrap, text="选择日期与时间，到点后系统将自动执行指定动作。",
                 bg=C["card"], fg=C["muted"], font=(FONT_UI, 9)).pack(anchor="w", pady=(2, 0))

        today = dt.date.today()
        self.schedule_date = today
        self.var_date_btn = tk.StringVar(value=self._date_text(today))

        row1 = tk.Frame(wrap, bg=C["card"])
        row1.pack(fill="x", pady=(12, 0))
        tk.Label(row1, text="日期", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(0, 10))
        self.btn_date = tk.Button(row1, textvariable=self.var_date_btn, relief="flat", bd=0,
                                  bg=C["card_hi"], fg=C["text"], activebackground=C["accent"],
                                  activeforeground="#04121c", font=(FONT_UI, 10),
                                  cursor="hand2", padx=16, pady=6, highlightthickness=0)
        self.btn_date.pack(side="left")
        self.btn_date.bind("<Button-1>", self._open_calendar)

        for text, offset in (("今天", 0), ("明天", 1), ("后天", 2), ("下周一", 7)):
            self._chip(row1, text, lambda o=offset: self._quick_date(o)).pack(side="left", padx=(10, 0))

        row2 = tk.Frame(wrap, bg=C["card"])
        row2.pack(fill="x", pady=(10, 0))
        tk.Label(row2, text="时间", bg=C["card"], fg=C["muted"],
                 font=(FONT_UI, 10)).pack(side="left", padx=(0, 10))
        self.var_sh = tk.StringVar(value="22")
        self.var_sm = tk.StringVar(value="30")
        tk.Label(row2, text="时", bg=C["card"], fg=C["muted"], font=(FONT_UI, 10)).pack(side="left", padx=(0, 4))
        sp_h = ttk.Spinbox(row2, from_=0, to=23, width=4, textvariable=self.var_sh,
                           font=(FONT_MONO, 13), justify="center", style="TSpinbox")
        sp_h.pack(side="left", padx=(0, 10))
        tk.Label(row2, text="分", bg=C["card"], fg=C["muted"], font=(FONT_UI, 10)).pack(side="left", padx=(0, 4))
        sp_m = ttk.Spinbox(row2, from_=0, to=59, width=4, textvariable=self.var_sm,
                           font=(FONT_MONO, 13), justify="center", style="TSpinbox")
        sp_m.pack(side="left", padx=(0, 10))
        for text, h, m in (("08:00", 8, 0), ("12:30", 12, 30), ("18:30", 18, 30), ("23:00", 23, 0)):
            self._chip(row2, text, lambda h=h, m=m: self._set_time(h, m)).pack(side="left", padx=(8, 0))

        self.var_sched_preview = tk.StringVar(value="")
        tk.Label(wrap, textvariable=self.var_sched_preview, bg=C["card"], fg=C["accent"],
                 font=(FONT_MONO, 20, "bold")).pack(anchor="w", pady=(10, 0))
        self.lbl_sched_when = tk.Label(wrap, text="", bg=C["card"], fg=C["muted"], font=(FONT_UI, 9))
        self.lbl_sched_when.pack(anchor="w", pady=(4, 0))

        ttk.Button(wrap, text="▶  开始定时关机", style="Accent.TButton",
                   command=self._start_schedule).pack(anchor="w", pady=(10, 0))

        tk.Label(wrap, text="提示：若到点前需要撤销，点击「取消该任务」或底部「取消全部关机任务」即可。",
                 bg=C["card"], fg=C["faint"], font=(FONT_UI, 8),
                 justify="left").pack(anchor="w", pady=(10, 0))

        sp_h.bind("<KeyRelease>", lambda e: self._refresh_sched_preview())
        sp_m.bind("<KeyRelease>", lambda e: self._refresh_sched_preview())
        self._refresh_sched_preview()

    def _date_text(self, date: dt.date) -> str:
        return "%d年%02d月%02d日  周%s" % (date.year, date.month, date.day,
                                          WEEK_CN[date.weekday()])

    def _quick_date(self, offset_days):
        today = dt.date.today()
        if offset_days == 7:                      # 下周一
            delta = (7 - today.weekday()) % 7 or 7
        else:
            delta = offset_days
        base = today + dt.timedelta(days=delta)
        self.schedule_date = base
        self.var_date_btn.set(self._date_text(base))
        self._refresh_sched_preview()

    def _set_time(self, h, m):
        self.var_sh.set("%02d" % h)
        self.var_sm.set("%02d" % m)
        self._refresh_sched_preview()

    def _open_calendar(self, event=None):
        CalendarPopup(self, self.schedule_date, self._on_date_picked, anchor=self.btn_date)

    def _on_date_picked(self, date: dt.date):
        self.schedule_date = date
        self.var_date_btn.set(self._date_text(date))
        self._refresh_sched_preview()

    def _schedule_target(self):
        try:
            h = int(self.var_sh.get() or 0)
            m = int(self.var_sm.get() or 0)
        except ValueError:
            return None
        if not (0 <= h <= 23 and 0 <= m <= 59):
            return None
        return dt.datetime.combine(self.schedule_date, dt.time(h, m)).timestamp()

    def _refresh_sched_preview(self):
        ts = self._schedule_target()
        if ts is None:
            self.var_sched_preview.set("--:--")
            self.lbl_sched_when.configure(text="请输入合法的时间")
            return
        remain = ts - time.time()
        self.var_sched_preview.set(dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M"))
        if remain <= 0:
            self.lbl_sched_when.configure(
                text="该时间已过，请选择将来的时间点（当前剩余 %s）" % fmt_natural(-remain))
        else:
            self.lbl_sched_when.configure(
                text="距离%s还有 %s，之后将自动%s" % (
                    dt.datetime.fromtimestamp(ts).strftime("%m月%d日 %H:%M"),
                    fmt_natural(remain), ACTIONS[self.var_action.get()]["label"]))

    # -- 日志页 ----------------------------------------------------------
    def _build_log_tab(self, parent):
        wrap = tk.Frame(parent, bg=C["card"])
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        row = tk.Frame(wrap, bg=C["card"])
        row.pack(fill="x")
        tk.Label(row, text="运行日志", bg=C["card"], fg=C["text"],
                 font=(FONT_UI, 13, "bold")).pack(side="left")
        ttk.Button(row, text="打开日志目录", style="TButton",
                   command=lambda: open_path(DATA_DIR)).pack(side="right")
        ttk.Button(row, text="清空", style="TButton", command=self._clear_log).pack(
            side="right", padx=8)

        holder = tk.Frame(wrap, bg=C["card"])
        holder.pack(fill="both", expand=True, pady=(12, 0))
        self.txt_log = tk.Text(holder, bg=C["panel"], fg=C["text"], bd=0,
                                highlightthickness=1, highlightbackground=C["border"],
                                insertbackground=C["accent"], wrap="none",
                                font=(FONT_MONO, 9), padx=10, pady=8)
        sb = ttk.Scrollbar(holder, orient="vertical", command=self.txt_log.yview,
                            style="Vertical.TScrollbar")
        self.txt_log.configure(yscrollcommand=sb.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.txt_log.configure(state="disabled")

        self._read_log_file()

    # ------------------------------------------------------------------
    # 计划控制
    # ------------------------------------------------------------------
    def _start_plan(self, mode, target_ts, desc):
        action = self.var_action.get()
        self.plan = {
            "mode": mode,
            "target": target_ts,
            "created": time.time(),
            "total": max(1.0, target_ts - time.time()),
            "action": action,
            "desc": desc,
        }
        self._save_state()
        self._log("已启动任务：%s（%s）" % (desc, fmt_datetime(target_ts)))
        if self.var_native.get() and ACTIONS[action]["delay"]:
            self._schedule_native(target_ts - time.time(), action)
        self._refresh_banner()
        self._update_status()

    def _schedule_native(self, seconds, action):
        ok, msg = self.power.schedule(seconds, action,
                                      "由定时关机助手安排的自动任务")
        if ok:
            self._log(msg)
        else:
            self._log(msg, "error")

    def _run_cmd(self, cmd):
        return self.power._run(cmd)

    def _start_countdown(self):
        total = self._countdown_seconds()
        if total <= 0:
            messagebox.showwarning(APP_TITLE, "请先设置大于 0 的倒计时时长。")
            return
        target = time.time() + total
        self._start_plan("countdown", target,
                         "倒计时 %s 后%s" % (fmt_natural(total),
                                          ACTIONS[self.var_action.get()]["label"]))

    def _start_schedule(self):
        ts = self._schedule_target()
        if ts is None:
            messagebox.showwarning(APP_TITLE, "请输入合法的时间。")
            return
        if ts <= time.time():
            messagebox.showwarning(APP_TITLE, "目标时间必须晚于当前时间。")
            return
        self._start_plan("scheduled", ts,
                         "%s 定时%s" % (dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M"),
                                        ACTIONS[self.var_action.get()]["label"]))

    def _cancel_plan(self):
        if not self.plan:
            return
        action = self.plan.get("action", "关机")
        ok, msg = self.power.abort()
        if ok:
            self._log(msg)
        self.plan = None
        self._save_state()
        self._log("任务已取消")
        self._refresh_banner()
        self._update_status()

    def _abort_all(self):
        ok, msg = self.power.abort()
        self._log(msg if ok else msg)
        if self.plan:
            self.plan = None
            self._save_state()
            self._log("本程序中的计划也已清除")
        self._refresh_banner()
        self._update_status()

    def _fire(self):
        if self._firing:
            return
        self._firing = True
        plan = self.plan or {}
        action = plan.get("action", "关机")
        self._log("倒计时结束，正在执行：%s" % ACTIONS[action]["label"], "warn")
        if IS_MAC and self.power.pid:              # 取消后台任务，避免重复触发
            self.power.abort()
        ok, msg = self.power.run_now(action)
        self._log(msg if ok else msg, "info" if ok else "error")
        self.plan = None
        self._save_state()
        self._refresh_banner()
        self._update_status()
        try:
            messagebox.showinfo(APP_TITLE, "定时任务已触发，系统即将%s。" %
                                ACTIONS[action]["label"])
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 刷新 / 定时器
    # ------------------------------------------------------------------
    def _tick(self):
        now = time.time()
        self._update_clock(now)

        if self.plan:
            remain = float(self.plan["target"]) - now
            if remain <= 0:
                self._fire()
            else:
                self._update_banner(remain)

        self._tick_health(now)
        self._after_id = self.after(250, self._tick)

    def _tick_health(self, now):
        self._roll_day()
        self._water_tick(now)
        self._kegel_tick(now)
        self._walk_tick(now)
        self._update_status()

    def _update_clock(self, now=None):
        now = now or time.time()
        now_dt = dt.datetime.fromtimestamp(now)
        self.var_clock.set(now_dt.strftime("%H:%M:%S"))
        self.var_date.set(now_dt.strftime("%Y-%m-%d") + "  星期" + WEEK_CN[now_dt.weekday()])

    def _refresh_banner(self):
        if not self.plan:
            self.banner.pack_forget()
            return
        self.banner.pack(fill="x", pady=(0, 12), before=self.nb)
        self._update_banner(float(self.plan["target"]) - time.time())

    def _update_banner(self, remain):
        if not self.plan:
            return
        self.var_plan_desc.set(self.plan.get("desc", ""))
        self.var_remain.set(fmt_clock(remain))
        total = float(self.plan.get("total", 1)) or 1
        ratio = max(0.0, min(1.0, remain / total))
        try:
            self.progress.configure(value=int(ratio * 1000))
        except Exception:
            pass

    def _update_status(self):
        if self.plan:
            remain = float(self.plan["target"]) - time.time()
            text = "状态：运行中（剩余 %s）· %s" % (fmt_natural(remain),
                                                   self.plan.get("desc", ""))
        else:
            text = "状态：空闲"

        extra = []
        if getattr(self, "water_running", False):
            extra.append("喝水 %s" % fmt_clock(max(0.0, self.water_next - time.time())))
        if getattr(self, "walk_active_until", 0):
            extra.append("走动 %s" % fmt_clock(max(0.0, self.walk_active_until - time.time())))
        elif getattr(self, "walk_running", False):
            extra.append("走动 %s" % fmt_clock(max(0.0, self.walk_next - time.time())))
        if getattr(self, "kegel_phase", "idle") in ("squeeze", "relax", "paused"):
            extra.append("提肛训练中")
        if extra:
            text += "  ·  健康：" + " / ".join(extra)
        self.var_status.set(text)

    # ------------------------------------------------------------------
    # 健康提醒：喝水 / 提肛 / 走动
    # ------------------------------------------------------------------
    def _beep(self, pattern):
        """后台线程播放提示音，避免阻塞界面"""
        if IS_MAC:
            def worker_mac():
                try:
                    subprocess.run(["osascript", "-e", "beep"],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                   timeout=10)
                except Exception:
                    pass
            threading.Thread(target=worker_mac, daemon=True).start()
            return

        try:
            import winsound
        except Exception:
            return

        def worker():
            try:
                for freq, dur in pattern:
                    winsound.Beep(freq, dur)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _show_toast(self, kind, title, message, color, on_ok=None, on_snooze=None):
        try:
            old = self._toasts.get(kind)
            if old is not None and old.winfo_exists():
                old.destroy()
        except Exception:
            pass
        toast = ReminderToast(self, kind, title, message, color, on_ok, on_snooze)
        self._toasts[kind] = toast

    # -- 通用：每日计数重置 ----------------------------------------------
    def _roll_day(self):
        today = dt.date.today().isoformat()
        if getattr(self, "water_day", None) != today:
            self.water_day = today
            self.water_count = 0
        if getattr(self, "walk_day", None) != today:
            self.walk_day = today
            self.walk_count = 0
            self.walk_total = 0

    # -- 喝水 ------------------------------------------------------------
    def _water_start(self):
        self.water_running = True
        interval = max(1, self.var_water_interval.get()) * 60
        if self.water_next <= time.time():
            self.water_next = time.time() + interval
        self._log("喝水提醒已启动，每 %d 分钟提醒一次" % (interval // 60))
        self._beep([(880, 120)])
        self._refresh_water()
        self._save_state()

    def _water_pause(self):
        self.water_running = False
        self._log("喝水提醒已暂停")
        self._refresh_water()
        self._save_state()

    def _water_reset(self):
        self.water_count = 0
        self.water_running = True
        self.water_next = time.time() + max(1, self.var_water_interval.get()) * 60
        self._log("喝水提醒已重置，今日计数清零")
        self._refresh_water()
        self._save_state()

    def _water_fire_now(self):
        self._water_fire()

    def _water_fire(self):
        self._beep([(880, 140), (1180, 140)])
        self._log("喝水提醒：到点了，该补充水分啦", "warn")
        self._show_toast(
            "water", "该喝水啦",
            "已经连续工作了一段时间，起来接杯水吧。\n"
            "喝掉一杯（约 200~300ml），顺便活动一下肩颈和眼睛。",
            C["accent"], on_ok=self._water_done,
            on_snooze=lambda: self._water_snooze(300))

    def _water_done(self):
        self.water_count += 1
        self._log("已记录一次喝水，今日 %d 杯" % self.water_count)
        self._refresh_water()
        self._save_state()

    def _water_snooze(self, seconds):
        self.water_next = time.time() + seconds
        self._log("喝水提醒推迟 %d 秒" % seconds)
        self._refresh_water()
        self._save_state()

    def _water_tick(self, now):
        if self.water_running and self.water_next and now >= self.water_next:
            self.water_next = now + max(1, self.var_water_interval.get()) * 60
            self._water_fire()
        self._refresh_water()

    def _refresh_water(self):
        interval = max(1, self.var_water_interval.get()) * 60
        if self.water_running:
            remain = max(0.0, self.water_next - time.time())
            self.var_water_rem.set(fmt_clock(remain))
            self.lbl_water_state.configure(
                text="下一次提醒：%s（间隔 %d 分钟）" %
                     (dt.datetime.fromtimestamp(self.water_next).strftime("%H:%M:%S"),
                      interval // 60))
            ratio = max(0.0, min(1.0, 1 - remain / interval))
            self.water_progress.configure(value=int(ratio * 1000))
        else:
            self.var_water_rem.set("--:--:--")
            self.lbl_water_state.configure(text="提醒已暂停")
            self.water_progress.configure(value=0)
        target = max(1, self.var_water_target.get())
        self.var_water_stat.set("今日 %d / %d 杯" % (self.water_count, target))
        self.water_stat_bar.configure(value=int(min(1.0, self.water_count / target) * 1000))

    # -- 提肛助手 --------------------------------------------------------
    def _kegel_beep(self, kind):
        if not self.var_kegel_sound.get():
            return
        patterns = {
            "squeeze": [(720, 120)],
            "relax": [(520, 160)],
            "done": [(880, 140), (1100, 140), (1320, 260)],
        }
        self._beep(patterns.get(kind, [(660, 120)]))

    def _kegel_start(self):
        if self.kegel_phase in ("squeeze", "relax"):
            return
        self.kegel_round = 1
        self.kegel_phase = "squeeze"
        self.kegel_phase_total = max(1, self.var_kegel_squeeze.get())
        self.kegel_phase_end = time.time() + self.kegel_phase_total
        self._kegel_beep("squeeze")
        self._log("提肛训练开始：收缩 %d 秒 / 放松 %d 秒 × %d 组" %
                  (self.var_kegel_squeeze.get(), self.var_kegel_relax.get(),
                   self.var_kegel_rounds.get()))
        self._refresh_kegel()

    def _kegel_pause(self):
        if self.kegel_phase in ("squeeze", "relax"):
            self.kegel_paused_phase = self.kegel_phase
            self.kegel_remaining = max(0.0, self.kegel_phase_end - time.time())
            self.kegel_phase = "paused"
            self._log("提肛训练已暂停")
        elif self.kegel_phase == "paused":
            self.kegel_phase = self.kegel_paused_phase
            self.kegel_phase_total = max(
                1, self.var_kegel_squeeze.get() if self.kegel_phase == "squeeze"
                else self.var_kegel_relax.get())
            self.kegel_phase_end = time.time() + self.kegel_remaining
            self._log("提肛训练继续")
        self._refresh_kegel()

    def _kegel_stop(self):
        self.kegel_phase = "idle"
        self.kegel_round = 0
        self._log("提肛训练已结束")
        self._refresh_kegel()
        self._save_state()

    def _kegel_reset(self):
        self.kegel_phase = "idle"
        self.kegel_round = 0
        self.kegel_done = 0
        self._log("提肛训练已重置")
        self._refresh_kegel()
        self._save_state()

    def _kegel_finish(self):
        self.kegel_phase = "done"
        self.kegel_done += 1
        self.kegel_round = 0
        self._kegel_beep("done")
        self._log("提肛训练完成，今日累计 %d 组" % self.kegel_done, "warn")
        self._show_toast(
            "kegel", "训练完成",
            "本组盆底肌训练已完成，记得放松休息。\n"
            "每天 2~3 组、循序渐进效果更佳。",
            C["success"], on_ok=self._kegel_stop)
        self._refresh_kegel()
        self._save_state()

    def _kegel_tick(self, now):
        if self.kegel_phase in ("squeeze", "relax") and now >= self.kegel_phase_end:
            if self.kegel_phase == "squeeze":
                self.kegel_phase = "relax"
                self.kegel_phase_total = max(1, self.var_kegel_relax.get())
                self.kegel_phase_end = now + self.kegel_phase_total
                self._kegel_beep("relax")
                self._log("进入放松阶段")
            else:
                rounds = max(1, self.var_kegel_rounds.get())
                if self.kegel_round >= rounds:
                    self._kegel_finish()
                else:
                    self.kegel_round += 1
                    self.kegel_phase = "squeeze"
                    self.kegel_phase_total = max(1, self.var_kegel_squeeze.get())
                    self.kegel_phase_end = now + self.kegel_phase_total
                    self._kegel_beep("squeeze")
                    self._log("开始第 %d / %d 组" % (self.kegel_round, rounds))
        self._refresh_kegel()

    def _refresh_kegel(self):
        phase = self.kegel_phase
        if phase == "squeeze":
            remain = max(0.0, self.kegel_phase_end - time.time())
            self.var_kegel_phase.set("收缩 · 收紧盆底肌")
            self.var_kegel_phase_label.configure(fg=C["warn"])
            self.lbl_kegel_timer.configure(fg=C["warn"])
            self.var_kegel_timer.set("%d" % int(math.ceil(remain)))
            ratio = max(0.0, min(1.0, 1 - remain / max(1, self.kegel_phase_total)))
            self.kegel_progress.configure(value=int(ratio * 1000))
        elif phase == "relax":
            remain = max(0.0, self.kegel_phase_end - time.time())
            self.var_kegel_phase.set("放松 · 完全放松")
            self.var_kegel_phase_label.configure(fg=C["success"])
            self.lbl_kegel_timer.configure(fg=C["success"])
            self.var_kegel_timer.set("%d" % int(math.ceil(remain)))
            ratio = max(0.0, min(1.0, 1 - remain / max(1, self.kegel_phase_total)))
            self.kegel_progress.configure(value=int(ratio * 1000))
        elif phase == "paused":
            self.var_kegel_phase.set("已暂停")
            self.var_kegel_phase_label.configure(fg=C["muted"])
            self.lbl_kegel_timer.configure(fg=C["muted"])
            self.var_kegel_timer.set("--")
            self.kegel_progress.configure(value=0)
        elif phase == "done":
            self.var_kegel_phase.set("本组训练已完成")
            self.var_kegel_phase_label.configure(fg=C["success"])
            self.lbl_kegel_timer.configure(fg=C["success"])
            self.var_kegel_timer.set("完成")
            self.kegel_progress.configure(value=1000)
        else:
            self.var_kegel_phase.set("准备开始")
            self.var_kegel_phase_label.configure(fg=C["accent"])
            self.lbl_kegel_timer.configure(fg=C["text"])
            self.var_kegel_timer.set("--")
            self.kegel_progress.configure(value=0)
        rounds = max(1, self.var_kegel_rounds.get())
        if self.kegel_phase in ("squeeze", "relax", "paused"):
            self.var_kegel_stat.set("第 %d / %d 组 · 今日已完成 %d 组" %
                                    (self.kegel_round, rounds, self.kegel_done))
        else:
            self.var_kegel_stat.set("共 %d 组 · 今日已完成 %d 组" % (rounds, self.kegel_done))

    # -- 走动提醒 --------------------------------------------------------
    def _walk_start(self):
        self.walk_running = True
        interval = max(10, self.var_walk_interval.get()) * 60
        if self.walk_next <= time.time():
            self.walk_next = time.time() + interval
        self._log("走动提醒已启动，每 %d 分钟提醒一次" % (interval // 60))
        self._beep([(660, 120)])
        self._refresh_walk()
        self._save_state()

    def _walk_pause(self):
        self.walk_running = False
        self._log("走动提醒已暂停")
        self._refresh_walk()
        self._save_state()

    def _walk_reset(self):
        self.walk_count = 0
        self.walk_total = 0
        self.walk_active_until = 0.0
        self.walk_running = True
        self.walk_next = time.time() + max(10, self.var_walk_interval.get()) * 60
        self._log("走动提醒已重置，今日统计清零")
        self._refresh_walk()
        self._save_state()

    def _walk_fire_now(self):
        self._walk_fire()

    def _walk_fire(self):
        duration = max(1, self.var_walk_duration.get()) * 60
        self.walk_active_until = time.time() + duration
        # 活动结束前不重复提醒：下一次提醒顺延一个间隔
        self.walk_next = time.time() + max(10, self.var_walk_interval.get()) * 60
        self._beep([(660, 180), (880, 180)])
        self._log("走动提醒：该起身活动啦（%d 分钟）" % (duration // 60), "warn")
        self._show_toast(
            "walk", "起来走一走",
            "已经坐了比较久，起来活动 %d 分钟吧。\n"
            "接杯水、走动两分钟、伸展肩颈和腰背。" % (duration // 60),
            C["warn"], on_ok=self._walk_finish,
            on_snooze=lambda: self._walk_snooze(120))
        self._refresh_walk()

    def _walk_finish(self):
        if self.walk_active_until:
            minutes = max(1, self.var_walk_duration.get())
            self.walk_count += 1
            self.walk_total += minutes
            self.walk_active_until = 0.0
            self._log("完成一次活动，今日 %d 次 / %d 分钟" % (self.walk_count, self.walk_total))
        self._refresh_walk()
        self._save_state()

    def _walk_snooze(self, seconds):
        self.walk_active_until = 0.0
        self.walk_next = time.time() + seconds
        self._log("走动提醒推迟 %d 秒" % seconds)
        self._refresh_walk()
        self._save_state()

    def _walk_tick(self, now):
        if self.walk_active_until:
            if now >= self.walk_active_until:
                self._walk_finish()
            self._refresh_walk()
            return
        if self.walk_running and self.walk_next and now >= self.walk_next:
            self.walk_next = now + max(10, self.var_walk_interval.get()) * 60
            self._walk_fire()
        self._refresh_walk()

    def _refresh_walk(self):
        if self.walk_active_until:
            remain = max(0.0, self.walk_active_until - time.time())
            self.var_walk_rem.set(fmt_clock(remain))
            self.lbl_walk_rem.configure(fg=C["success"])
            self.lbl_walk_state.configure(text="活动进行中：起身走动 %d 秒" %
                                          int(math.ceil(remain)))
            duration = max(1, self.var_walk_duration.get()) * 60
            ratio = max(0.0, min(1.0, 1 - remain / duration))
            self.walk_progress.configure(value=int(ratio * 1000))
        elif self.walk_running:
            remain = max(0.0, self.walk_next - time.time())
            self.var_walk_rem.set(fmt_clock(remain))
            self.lbl_walk_rem.configure(fg=C["accent"])
            self.lbl_walk_state.configure(
                text="下一次提醒：%s（间隔 %d 分钟）" %
                     (dt.datetime.fromtimestamp(self.walk_next).strftime("%H:%M:%S"),
                      max(10, self.var_walk_interval.get())))
            interval = max(10, self.var_walk_interval.get()) * 60
            ratio = max(0.0, min(1.0, 1 - remain / interval))
            self.walk_progress.configure(value=int(ratio * 1000))
        else:
            self.var_walk_rem.set("--:--:--")
            self.lbl_walk_rem.configure(fg=C["accent"])
            self.lbl_walk_state.configure(text="提醒已暂停")
            self.walk_progress.configure(value=0)
        self.var_walk_stat.set("今日活动 %d 次 · 累计 %d 分钟" %
                               (self.walk_count, self.walk_total))

    # ------------------------------------------------------------------
    # 日志
    # ------------------------------------------------------------------
    def _log(self, message, level="info"):
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = "[%s] %s" % (stamp, message)
        self.log_lines.append(line)
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass
        if getattr(self, "txt_log", None) is not None:
            tag = {"info": "info", "warn": "warn", "error": "error"}.get(level, "info")
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", line + "\n", tag)
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")

    def _read_log_file(self):
        try:
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
        except Exception:
            content = ""
        if content:
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", content + "\n")
            self.txt_log.configure(state="disabled")
            self.txt_log.see("end")

    def _clear_log(self):
        self.log_lines = []
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")
        try:
            if os.path.exists(LOG_FILE):
                os.remove(LOG_FILE)
        except Exception:
            pass
        self._log("日志已清空")

    # ------------------------------------------------------------------
    # 杂项
    # ------------------------------------------------------------------
    def _center(self, w, h):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def _on_close(self):
        self._save_state()
        try:
            self.after_cancel(self._after_id)
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


def main():
    app = ShutdownApp()
    app.mainloop()


if __name__ == "__main__":
    main()
