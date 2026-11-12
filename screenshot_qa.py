# -*- coding: utf-8 -*-
"""WorkBuddy 截图提问 —— 伴生悬浮按钮 / 全局热键 / 框选截图 / 自动进输入框附件。

交互：
    点击吸附在输入框旁的「截图提问」按钮，或按 Ctrl+Alt+Q
        -> 隐藏 WorkBuddy 窗口 -> 全屏框选 -> 截图
        -> 图片写入剪贴板 -> 恢复 WorkBuddy 并自动粘贴成附件
        -> 你直接打字提问、回车发送

    Ctrl+Alt+X 退出；按钮上点右键可打开菜单（自动粘贴、开机自启、复位、说明、退出）。

命令行：
    pythonw screenshot_qa.py            常驻（配合 启动.vbs）
    python  screenshot_qa.py --once     只执行一次截图流程后退出
    python  screenshot_qa.py --selftest 只做环境自检，不启动界面

只依赖：Python 3.11 + Pillow + tkinter（系统自带），不改动 WorkBuddy 安装目录。
"""
import ctypes
import ctypes.wintypes as wt
import json
import os
import queue
import struct
import subprocess
import sys
import threading
import time
import traceback

from collections import Counter

from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageGrab, ImageTk

# ---------------------------------------------------------------- 基础路径与配置
APP_DIR = os.path.dirname(os.path.abspath(__file__))
SHOT_DIR = os.path.join(APP_DIR, "截图")
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
ERROR_LOG = os.path.join(APP_DIR, "错误.log")
DIAG_LOG = os.path.join(APP_DIR, "点击诊断.log")        # 每次自动粘贴一行，排障用（只留最近 60 行）
HIDE_FLAG = os.path.join(APP_DIR, "窗口隐藏中.flag")   # 崩溃兜底：存在即说明 WorkBuddy 可能被我们藏了
DEBUG_MOVE = False                     # 临时：把每轮调度写进 _move.log，定位跟随延迟


def log_error(tag):
    """把当前异常写进 错误.log —— pythonw 没有控制台，不落文件就永远查不到。"""
    try:
        import traceback
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write("\n[%s] %s\n%s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"),
                                         tag, traceback.format_exc()))
    except Exception:
        pass


def log_diag(msg):
    """自动粘贴的落点记录。只在粘贴时写一行，平时零开销。"""
    try:
        lines = []
        if os.path.exists(DIAG_LOG):
            with open(DIAG_LOG, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()[-60:]
        lines.append("[%s] %s" % (time.strftime("%m-%d %H:%M:%S"), msg))
        with open(DIAG_LOG, "w", encoding="utf-8") as f:
            f.write("\n".join(lines[-60:]) + "\n")
    except Exception:
        pass

DEFAULT_CONFIG = {
    # 定位只有一个公式，没有任何缓存偏移、也不支持拖动：
    #   中间内容区 = [侧栏右缘, 右栏左缘]（无右栏时右栏左缘 = 窗口右缘）
    #   x = 该区间正中 + offset_x − 按钮宽/2
    # 侧边栏宽度实测 240px（可调但一般不变），量一次写死。
    # 右栏左缘由 detect_right_panel() 实测（当前 826），只在窗口矩形变化时更新。
    # 窗口四边取自 DWM 可见边框，最大化比 GetWindowRect 少 8px 阴影，两态才对得上。
    "sidebar_w": 240,              # 左侧边栏宽度（实测，一般不变）
    "offset_x": -60,               # 兜底横向偏移：测不到连接器图标时才用（负数往左）
    "connector_gap": 12,           # 按钮左缘距「连接器」蓝色图标右缘的间距
    "compact_below": 620,          # 内容区窄于此值时只显示图标、隐藏文字
    "recheck_sec": 1.6,            # 静止时每隔多少秒复核一次连接器位置（技能区变长靠它跟）
    "theme_recheck_sec": 2.4,      # 皮肤复核间隔（换主题是低频操作，不必跟得太勤）
    # ---- 跟随调度：运动和静止用两套节奏 ----
    # 运动中只查窗口矩形（约 0.1ms/轮），所以可以拉到 6ms 一轮，跟手到和窗口同级；
    # 静止时 120ms 一轮足够，几乎不占 CPU。
    "moving_ms": 6,                # 运动中（拖动/缩放）的轮询间隔
    "idle_ms": 16,                 # 静止时的轮询间隔（也是发现「开始动」的最大延迟）
    "moving_hold_ms": 320,         # 矩形变了之后，多久内仍算「运动中」
    "offset_y": -79,               # 兜底纵向位置：测不到图标时才用
    "auto_paste": True,
    # 粘贴前先点一下输入框正文：用户可能压根没点进输入框就截图（直接点悬浮按钮
    # 或按热键），此时窗口虽在前台，焦点却不在输入框上，Ctrl+V 会静默落空。
    "click_before_paste": True,
    "autostart": False,
    "hide_when_inactive": True,    # WorkBuddy 不在前台时按钮一起收起
    # ---- 配色跟随皮肤（开源后别人可能用深色主题，颜色一律不写死）----
    # theme=auto 按窗口实测明暗；填 light / dark 可强制。
    # bg_color / ink / hover_bg 填 "#RRGGBB" 可逐项覆盖自动值。
    "theme": "auto",
    "bg_color": "",
    "ink": "",
    "hover_bg": "",
}

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "WorkBuddy截图提问"
def pythonw_path():
    """无窗口解释器路径 —— 不写死安装位置，换台机器、换个 Python 版本都能用。

    优先级：当前解释器本身就是 pythonw → 同目录 pythonw.exe → 同目录 python.exe
             → 交给 PATH 里的 pythonw.exe。
    程序是被哪个 Python 拉起来的，就用它自己的解释器，这样「注册表自启项 /
    计划任务里记下的命令」和「手动双击启动的」永远是同一个解释器，不会打架。
    """
    if os.path.basename(sys.executable).lower().startswith("pythonw"):
        return sys.executable
    d = os.path.dirname(sys.executable)
    for name in ("pythonw.exe", "python.exe"):
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    return "pythonw.exe"


PYTHONW = pythonw_path()
MAIN_SCRIPT = os.path.join(APP_DIR, "screenshot_qa.py")
SINGLETON_NAME = "Local\\WorkBuddyScreenshotQA_Singleton"
ERROR_ALREADY_EXISTS = 183

# 按钮伪装成输入框工具栏里的原生文字按钮（对齐「默认权限」）：
# 无填充（底色当透明键）、墨色文字/图标，hover 才浮出一层同色系的淡色圆角底。
BTN_W, BTN_H = 92, 32        # 高度对齐原生胶囊：实测「默认权限」胶囊高 32px
COMPACT_W = 32              # 输入框变窄后只留图标时的宽度（方形，与高度一致）
ICON_W, ICON_H = 15, 14     # 取景框图标尺寸：原生工具栏图标约 15px 见方，再大就显突兀
INK = (25, 57, 65)          # 兜底墨色：取自「默认权限」文字主色 #193941
HOVER_BG = (241, 241, 241)  # 兜底胶囊底：实测原生 hover ≈ #F1F1F1（中性，不加色）

# 下面三个默认值只是「还没量到皮肤前」的兜底（亮色主题），
# 真正用的是运行时 probe_theme() 量出来的 PAL，深色主题也能对。
PAL = {"bg": (255, 255, 255), "ink": INK, "dark": False, "hover": HOVER_BG}

# ---------------------------------------------------------------- Win32 初始化
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)   # 坐标 == 物理像素
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
gdi32 = ctypes.windll.gdi32
winmm = ctypes.windll.winmm      # timeBeginPeriod：运动期间提高定时器精度

WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000   # 悬浮层不抢焦点，否则键盘输入会打不进输入框
GWL_EXSTYLE = -20
SW_HIDE, SW_SHOW, SW_RESTORE = 0, 5, 9
WM_HOTKEY, WM_QUIT = 0x0312, 0x0012
MOD_ALT, MOD_CONTROL = 0x0001, 0x0002
# SetWindowPos：位置直写用的标志（不改尺寸 / 不改层级 / 不抢焦点）
SWP_NOSIZE, SWP_NOZORDER, SWP_NOACTIVATE = 0x0001, 0x0004, 0x0010
SWP_NOMOVE, SWP_SHOWWINDOW = 0x0002, 0x0040
HWND_TOP = 0
GA_ROOT = 2
VK_CONTROL, VK_V, VK_MENU = 0x11, 0x56, 0x12
KEYEVENTF_KEYUP = 0x0002
MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP = 0x0002, 0x0004
CF_DIB = 8
GMEM_MOVEABLE = 0x0002
SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN = 76, 77
SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN = 78, 79
DWMWA_CLOAKED = 14          # 关到托盘 / 虚拟桌面时窗口会被 DWM cloak
DWMWA_EXTENDED_FRAME_BOUNDS = 9   # 可见边框（去掉 Win10/11 那圈不可见 resize 阴影）

try:
    dwmapi = ctypes.windll.dwmapi
    dwmapi.DwmGetWindowAttribute.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                             ctypes.c_void_p, ctypes.c_int]
    dwmapi.DwmGetWindowAttribute.restype = ctypes.c_long
except Exception:
    dwmapi = None


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class MSG(ctypes.Structure):
    _fields_ = [("hwnd", ctypes.c_void_p), ("message", ctypes.c_uint),
                ("wParam", ctypes.c_size_t), ("lParam", ctypes.c_ssize_t),
                ("time", ctypes.c_ulong), ("pt_x", ctypes.c_long), ("pt_y", ctypes.c_long)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint), ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long), ("biPlanes", ctypes.c_ushort),
                ("biBitCount", ctypes.c_ushort), ("biCompression", ctypes.c_uint),
                ("biSizeImage", ctypes.c_uint), ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", ctypes.c_uint),
                ("biClrImportant", ctypes.c_uint)]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint * 3)]


SRCCOPY = 0x00CC0020
# 64 位下句柄必须声明成指针大小，否则会被默认的 c_int 截断导致崩溃
user32.GetDC.argtypes = [ctypes.c_void_p]
user32.GetDC.restype = ctypes.c_void_p
user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = ctypes.c_void_p
gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.SelectObject.restype = ctypes.c_void_p
gdi32.BitBlt.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                         ctypes.c_int, ctypes.c_int, ctypes.c_void_p,
                         ctypes.c_int, ctypes.c_int, ctypes.c_uint]
gdi32.GetDIBits.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint,
                            ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p,
                            ctypes.c_uint]
gdi32.DeleteObject.argtypes = [ctypes.c_void_p]
gdi32.DeleteDC.argtypes = [ctypes.c_void_p]
# 补焦点用：SetCursorPos 直接把指针挪过去点一下（不改缩放 / 不改坐标原点）
user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
user32.SetCursorPos.restype = ctypes.c_int
user32.mouse_event.argtypes = [ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
                               ctypes.c_uint, ctypes.c_size_t]
# 落点归属判定用：句柄是指针大小，按默认 c_int 会被截断（超过 2^31 的句柄就认错了）
user32.WindowFromPoint.argtypes = [wt.POINT]
user32.WindowFromPoint.restype = ctypes.c_void_p
user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
user32.GetAncestor.restype = ctypes.c_void_p


# 抓屏用的 GDI 对象常驻复用。实测抓 52x5 和抓 560x2 都约 19ms ——
# 耗时几乎全在 CreateCompatibleDC / CreateCompatibleBitmap 这些固定开销上，
# 跟区域大小基本无关。跟随要每几毫秒跑一轮，这里把对象留着重用、
# 只按需扩容位图，固定开销就摊掉了（单次降到 3~5ms）。
# 只在 Tk 主线程调用，不需要加锁。
_G = {"hdc": None, "mem": None, "bmp": None, "old": None, "w": 0, "h": 0, "buf": None}


def _gdi_pair():
    if _G["hdc"] is None:
        _G["hdc"] = user32.GetDC(None)
        if _G["hdc"]:
            _G["mem"] = gdi32.CreateCompatibleDC(_G["hdc"])
    return _G["hdc"], _G["mem"]


def grab_region(x0, y0, x1, y1):
    """只拷贝屏幕指定矩形（物理像素），返回 PIL Image 或 None。

    PIL 的 ImageGrab.grab(bbox=...) 内部总是先抓整个屏幕再裁剪，
    880x100 的窄条要 40~70ms；而定位复核是周期性跑的，这个卡顿会拖慢跟随。
    这里直接 BitBlt 只拷目标区域，且复用 DC / 位图。
    """
    w, h = int(x1 - x0), int(y1 - y0)
    if w <= 0 or h <= 0:
        return None
    hdc, mem = _gdi_pair()
    if not hdc or not mem:
        return None
    if w > _G["w"] or h > _G["h"]:
        nw, nh = max(w, _G["w"]), max(h, _G["h"])
        bmp = gdi32.CreateCompatibleBitmap(hdc, nw, nh)
        if not bmp:
            return None
        if _G["bmp"]:
            gdi32.SelectObject(mem, _G["old"])
            gdi32.DeleteObject(_G["bmp"])
        _G["old"] = gdi32.SelectObject(mem, bmp)
        _G["bmp"], _G["w"], _G["h"] = bmp, nw, nh
        _G["buf"] = None
    if not gdi32.BitBlt(mem, 0, 0, w, h, hdc, int(x0), int(y0), SRCCOPY):
        return None
    bi = BITMAPINFO()
    bi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bi.bmiHeader.biWidth = w
    bi.bmiHeader.biHeight = -h            # 负数 = 从上到下，省一次翻转
    bi.bmiHeader.biPlanes = 1
    bi.bmiHeader.biBitCount = 32
    bi.bmiHeader.biCompression = 0        # BI_RGB
    need = w * h * 4
    if _G["buf"] is None or len(_G["buf"]) < need:
        _G["buf"] = ctypes.create_string_buffer(_G["w"] * _G["h"] * 4)
    if not gdi32.GetDIBits(mem, _G["bmp"], 0, h, _G["buf"], ctypes.byref(bi), 0):
        return None
    # 用 string_at 只取本次需要的那一段：_G["buf"] 是按最大区域分配的，
    # 直接 .raw 会把整块（可能 1MB 以上）都复制一遍。
    return Image.frombytes("RGB", (w, h), ctypes.string_at(_G["buf"], need),
                           "raw", "BGRX", 0, 1)


WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

user32.EnumWindows.argtypes = [WNDENUMPROC, ctypes.c_void_p]
user32.IsWindowVisible.argtypes = [ctypes.c_void_p]
user32.GetWindowTextLengthW.argtypes = [ctypes.c_void_p]
user32.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
user32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.GetWindowLongW.restype = ctypes.c_long
user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(RECT)]
user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(wt.DWORD)]
user32.GetForegroundWindow.restype = ctypes.c_void_p
user32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
user32.BringWindowToTop.argtypes = [ctypes.c_void_p]
user32.SetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]
user32.GetParent.argtypes = [ctypes.c_void_p]
user32.GetParent.restype = ctypes.c_void_p
user32.IsIconic.argtypes = [ctypes.c_void_p]
user32.OpenClipboard.argtypes = [ctypes.c_void_p]
user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
user32.SetClipboardData.restype = ctypes.c_void_p
user32.RegisterHotKey.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_uint, ctypes.c_uint]
user32.PostThreadMessageW.argtypes = [ctypes.c_ulong, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
user32.GetMessageW.argtypes = [ctypes.POINTER(MSG), ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint]
kernel32.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
kernel32.OpenProcess.restype = ctypes.c_void_p
kernel32.QueryFullProcessImageNameW.argtypes = [ctypes.c_void_p, ctypes.c_ulong,
                                                ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_ulong)]
kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]


def exe_name(pid):
    h = kernel32.OpenProcess(0x1000, False, pid)
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(1024)
        size = ctypes.c_ulong(1024)
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value).lower()
        return ""
    finally:
        kernel32.CloseHandle(h)


def window_pid(hwnd):
    pid = wt.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def window_is_active(hwnd):
    """窗口是否真的摆在桌面上：可见、未最小化、未离屏、未被 DWM 隐藏。

    IsWindowVisible 对最小化 / 关到托盘的窗口仍可能返回 True，
    所以按钮的显隐不能只看它。
    """
    if not user32.IsWindowVisible(hwnd) or user32.IsIconic(hwnd):
        return False
    rc = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rc))
    if rc.left <= -10000 or rc.top <= -10000:     # 最小化后的离屏矩形
        return False
    if dwmapi:
        try:
            v = ctypes.c_int(0)
            if dwmapi.DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED,
                                            ctypes.byref(v), ctypes.sizeof(v)) == 0 \
                    and v.value:
                return False
        except Exception:
            pass
    return True


def foreground_is_foreign(wb_hwnd, own_pid):
    """前台是不是别的程序：是的话按钮不该 topmost 浮在人家窗口上面。"""
    fg = user32.GetForegroundWindow()
    if not fg or fg == wb_hwnd:
        return False
    fg_pid = window_pid(fg)
    if fg_pid == 0 or fg_pid == own_pid:          # 自己（悬浮按钮 / 框选层）抢的焦点不算
        return False
    wb_pid = window_pid(wb_hwnd)
    return not (wb_pid and fg_pid == wb_pid)      # WorkBuddy 自己的子窗口也不算


def apply_no_activate(widget):
    """悬浮窗口设成不激活：仍能接收鼠标点击，但不会抢走键盘焦点。

    否则按钮 deiconify 后自己变成前台窗口，用户打字会落不到 WorkBuddy
    输入框里，而且前台判断会永远认为「还是自己」而拒绝隐藏。
    """
    try:
        hwnd = user32.GetParent(widget.winfo_id())
        if hwnd:
            ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                                  ex | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE)
    except Exception:
        pass


def restore_foreground(hwnd):
    """把焦点还回去；Alt trick 用来绕过 Windows 的前台锁定。"""
    try:
        user32.keybd_event(VK_MENU, 0, 0, 0)
        user32.SetForegroundWindow(hwnd)
        user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
    except Exception:
        pass


def list_windows(include_hidden=False):
    out = []

    @WNDENUMPROC
    def cb(hwnd, _):
        # 救援被我们 SW_HIDE 的窗口时不做这些过滤
        if not include_hidden and not window_is_active(hwnd):
            return True
        n = user32.GetWindowTextLengthW(hwnd)
        if n == 0:
            return True
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.GetWindowTextW(hwnd, buf, n + 1)
        title = buf.value.strip()
        if not title:
            return True
        if user32.GetWindowLongW(hwnd, GWL_EXSTYLE) & WS_EX_TOOLWINDOW:
            return True
        pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        rc = RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rc))
        w, h = rc.right - rc.left, rc.bottom - rc.top
        if w < 200 or h < 120:
            return True
        out.append({"hwnd": hwnd, "title": title, "exe": exe_name(pid.value),
                    "rect": (rc.left, rc.top, rc.right, rc.bottom), "w": w, "h": h})
        return True

    user32.EnumWindows(cb, 0)
    return out


# 上次找到的 WorkBuddy 窗口。跟随每几毫秒跑一轮，EnumWindows + 逐窗口取
# 进程名要 0.5ms，攒起来比真正的窗口查询还贵；这里缓存 hwnd，
# 后续只用 IsWindow / 可见性 / 尺寸做校验（约 20us），失效了才重新枚举。
_WB = {"wb": None}


def find_workbuddy(include_hidden=False):
    if not include_hidden:
        c = _WB["wb"]
        if c is not None:
            hwnd = c["hwnd"]
            if user32.IsWindow(hwnd) and window_is_active(hwnd):
                rc = RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rc))
                w, h = rc.right - rc.left, rc.bottom - rc.top
                if w >= 500 and h >= 400:
                    c["rect"] = (rc.left, rc.top, rc.right, rc.bottom)
                    c["w"], c["h"] = w, h
                    return c
            _WB["wb"] = None
    ws = [w for w in list_windows(include_hidden)
          if "workbuddy" in w["exe"] and w["w"] >= 500 and w["h"] >= 400]
    wb = max(ws, key=lambda w: w["w"] * w["h"]) if ws else None
    if not include_hidden:
        _WB["wb"] = wb
    return wb


def visible_rect(hwnd):
    """窗口的可见矩形（物理像素）。

    GetWindowRect 在最大化时会把 Win10/11 那圈不可见的 resize 阴影算进去：
    实测最大化返回 (-8,-8,1928,1168)，而真实可见的是 (0,0,1920,1160)——
    底边差 8px，按钮就跟着往下掉 8px，这正是「最大化上下位置变了」的原因。
    DWM 的 EXTENDED_FRAME_BOUNDS 给的是可见范围，两态一致；取不到就退回旧值。
    """
    fallback = None
    rc = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rc))
    fallback = (rc.left, rc.top, rc.right, rc.bottom)
    if not dwmapi:
        return fallback
    try:
        v = RECT()
        if dwmapi.DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
                                        ctypes.byref(v), ctypes.sizeof(v)) == 0:
            r = (v.left, v.top, v.right, v.bottom)
            # 窗口最小化 / 刚创建时可能返回空或异常值，这种一律不用
            if r[2] - r[0] >= 200 and r[3] - r[1] >= 120:
                return r
    except Exception:
        pass
    return fallback


def workbuddy_rect(wb):
    """定位唯一认可的窗口矩形：可见边框优先。"""
    return visible_rect(wb["hwnd"]) if wb else None


def _blue_mask(im):
    """把截图里的蓝色像素压成黑白掩膜。

    只用通道运算（C 实现），比逐像素 Python 循环快一个数量级 ——
    因为定位复核要周期性跑，这里必须便宜。
    """
    r, g, b = im.split()
    m = ImageChops.multiply(
        b.point(lambda v: 255 if v > 170 else 0),
        ImageChops.subtract(b, r).point(lambda v: 255 if v > 90 else 0))
    return ImageChops.multiply(
        m, ImageChops.subtract(b, g).point(lambda v: 255 if v > 25 else 0))


def detect_connector_icon(rect, sidebar_w):
    """找输入框工具栏里那个蓝色「连接器」图标。

    返回 (左缘, 右缘, 纵向中心) 绝对坐标；找不到返回 None。

    实测：图标是蓝色双箭头，窗口内宽 28px、高 16px，主色 b>170 且 b-r>90，
    左侧「+」「默认权限」是墨青灰，色相差别很大，不会混。
    返回纵向中心 cy 是为了让按钮垂直对齐图标本身 —— 技能区变高/换行时
    工具栏整体上移，y 会自己跟上，不再依赖「距窗口底边多少像素」这个死数。
    """
    left, top, right, bottom = rect
    y0, y1 = bottom - 110, bottom - 10
    x0, x1 = left + sidebar_w + 20, min(right - 20, left + sidebar_w + 980)
    vx, vy, vw, vh = virtual_screen()
    x0, x1 = max(x0, vx), min(x1, vx + vw)
    y0, y1 = max(y0, vy), min(y1, vy + vh)
    if x1 - x0 < 80 or y1 - y0 < 20:
        return None
    try:
        im = grab_region(x0, y0, x1, y1)
    except Exception:
        im = None
    if im is None:
        return None

    m = _blue_mask(im)
    W, H = im.size
    # 列投影：缩成 (W,1) 取每列平均，>2 说明这一列有蓝色像素
    cols = list(m.resize((W, 1), Image.BOX).tobytes())
    xs = [i for i, v in enumerate(cols) if v > 2]
    if not xs:
        return None

    clusters, st, pr = [], xs[0], xs[0]
    for x in xs[1:]:
        if x - pr > 14:
            clusters.append((st, pr))
            st = x
        pr = x
    clusters.append((st, pr))

    cands = []
    for a, b in clusters:
        if not 18 <= b - a + 1 <= 46:      # 图标实测 28px 宽
            continue
        sub = m.crop((a, 0, b + 1, H))
        count = sub.histogram()[255]       # 该簇的蓝色像素数
        if count < 60:
            continue
        rows = list(sub.resize((1, H), Image.BOX).tobytes())
        ys = [i for i, v in enumerate(rows) if v > 2]
        if not ys:
            continue
        cands.append((max(ys), count, a, b, (min(ys) + max(ys)) / 2.0))
    if not cands:
        return None
    # 先按「最靠下」筛行：工具栏永远在输入框最下一行，这样能排除上方的
    # 蓝色技能 chip / 链接。同一行里再取最左边那个 —— 连接器是这一行第一个
    # 蓝色图标；右侧的模型名 logo（DeepSeek 也是蓝的）不能抢走锚点。
    cands.sort(key=lambda c: (-c[0], -c[1]))
    bottom_cy = cands[0][4]
    row = [c for c in cands if abs(c[4] - bottom_cy) <= 4]
    row.sort(key=lambda c: c[2])
    _, _, a, b, cy = row[0]
    return (x0 + a, x0 + b, cy + y0)


def detect_right_panel(rect):
    """返回右侧内容区的左边缘（相对窗口左边缘的 x）；未展开返回 0。

    只在窗口矩形变化时调用，不产生每轮开销。
    实测：展开时右栏是纯白 (255,255,255) 大色块，左边缘在窗口内 x=826；
    中间内容区是 (248,253,255)。无右栏时整个右半区都是内容区背景色。
    做法：中部取 4 条水平采样线，比较「最右列背景色」与「中间区背景色」，
    不同且从右边界往左连续同色在 200px ~ 85% 采样宽之间，记一票并记下边缘；
    >=3 票判为展开，取这几条线边缘的中位数，避免个别行被气泡干扰。
    """
    left, top, right, bottom = rect
    w, h = right - left, bottom - top
    x0 = left + int(w * 0.40)
    y0 = top + int(h * 0.24)
    y1 = top + int(h * 0.68)
    vx, vy, vw, vh = virtual_screen()
    x0, x1 = max(x0, vx), min(right, vx + vw)
    y0, y1 = max(y0, vy), min(y1, vy + vh)
    if x1 - x0 < 120 or y1 - y0 < 60:
        return 0
    # 用 grab_region 而不是 ImageGrab：后者内部先抓整屏再裁剪，实测 52ms；
    # 前者只拷这块，复用 GDI 后约 20ms。这里每 0.4s 可能调一次，差别明显。
    im = grab_region(x0, y0, x1, y1)
    if im is None:
        return 0
    px = im.load()
    W, H = im.size

    def diff(a, b):
        return abs(a[0] - b[0]) + abs(a[1] - b[1]) + abs(a[2] - b[2])

    votes, edges = 0, []
    for i in range(4):
        y = int(H * (0.12 + 0.24 * i))
        bg = px[W - 1, y]                  # 右栏（或内容区）背景
        mid = px[2, y]                     # 中间内容区背景
        if diff(bg, mid) <= 4:             # 同色 = 没有右栏
            continue
        x = W - 1
        while x >= 0 and diff(px[x, y], bg) <= 8:
            x -= 1
        run = W - 1 - x
        if 200 <= run <= W * 0.85:         # 排除局部色块 / 整片同色
            votes += 1
            edges.append(W - 1 - run)
    if votes < 3 or not edges:
        return 0
    edges.sort()
    return (x0 - left) + edges[len(edges) // 2]


def mark_hidden(on):
    """隐藏 WorkBuddy 前留下标记；万一进程被强杀，下次启动能自动把窗口找回来。"""
    try:
        if on:
            with open(HIDE_FLAG, "w", encoding="utf-8") as f:
                f.write(time.strftime("%Y-%m-%d %H:%M:%S"))
        elif os.path.exists(HIDE_FLAG):
            os.remove(HIDE_FLAG)
    except Exception:
        pass


def restore_orphan():
    if not os.path.exists(HIDE_FLAG):
        return False
    wb = find_workbuddy(include_hidden=True)
    if wb:
        try:
            user32.ShowWindow(wb["hwnd"], SW_SHOW)
            user32.SetForegroundWindow(wb["hwnd"])
        except Exception:
            pass
    mark_hidden(False)
    return True


def force_restore():
    """手动救援：不管有没有标记，把 WorkBuddy 主窗口显示回来。"""
    wb = find_workbuddy(include_hidden=True)
    if wb:
        try:
            user32.ShowWindow(wb["hwnd"], SW_SHOW)
            user32.SetForegroundWindow(wb["hwnd"])
        except Exception:
            pass
    mark_hidden(False)
    return wb


def virtual_screen():
    vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
    vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
    vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
    vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    return vx, vy, vw, vh


# ---------------------------------------------------------------- 剪贴板图片
def build_dib(img):
    """把 PIL 图转成 CF_DIB 需要的 32bpp BGRA / bottom-up 数据。"""
    img = img.convert("RGBA")
    w, h = img.size
    r, g, b, a = img.split()
    bgra = Image.merge("RGBA", (b, g, r, a)).tobytes()
    stride = w * 4
    header = struct.pack("<IiiHHIIiiII", 40, w, h, 1, 32, 0,
                         stride * h, 0, 0, 0, 0)
    rows = [bgra[y * stride:(y + 1) * stride] for y in range(h - 1, -1, -1)]
    return header + b"".join(rows)


def copy_image_to_clipboard(img):
    data = build_dib(img)
    for _ in range(6):
        if user32.OpenClipboard(None):
            break
        time.sleep(0.06)
    else:
        return False, "剪贴板被其它程序占用"
    try:
        user32.EmptyClipboard()
        hmem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not hmem:
            return False, "内存分配失败"
        ptr = kernel32.GlobalLock(hmem)
        if not ptr:
            return False, "内存锁定失败"
        ctypes.memmove(ptr, data, len(data))
        kernel32.GlobalUnlock(hmem)
        if not user32.SetClipboardData(CF_DIB, hmem):
            return False, "写入剪贴板失败"
        return True, "ok"
    finally:
        user32.CloseClipboard()


# ---------------------------------------------------------------- 配置与自启
def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg.update({k: v for k, v in json.load(f).items() if k in DEFAULT_CONFIG})
    except Exception:
        pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def autostart_command():
    """注册表里的启动命令：直接拉 pythonw，不绕 wscript/vbs。"""
    pyw = PYTHONW if os.path.exists(PYTHONW) else "pythonw.exe"
    return '"%s" "%s"' % (pyw, MAIN_SCRIPT)


def get_autostart():
    """菜单勾选取注册表的真实状态，不信任本地缓存。"""
    import winreg
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_READ)
        try:
            winreg.QueryValueEx(key, RUN_NAME)
            return True
        finally:
            winreg.CloseKey(key)
    except Exception:
        return False


def set_autostart(enabled):
    import winreg
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE)
        if enabled:
            winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ, autostart_command())
        else:
            try:
                winreg.DeleteValue(key, RUN_NAME)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def launch_now():
    """安装完立刻拉起一次守护，不必等下次登录（单例锁会挡住重复启动）。"""
    pyw = PYTHONW if os.path.exists(PYTHONW) else "pythonw.exe"
    DETACHED_PROCESS = 0x00000008
    CREATE_NO_WINDOW = 0x08000000
    try:
        subprocess.Popen([pyw, MAIN_SCRIPT], cwd=APP_DIR,
                         creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW,
                         close_fds=True)
        return True
    except Exception:
        return False


# 任务名用纯 ASCII，避免 bat / 注册表里的中文编码问题
TASK_NAME = "WorkBuddyScreenshotQA_Guardian"


def _ps_quote(s):
    """PowerShell 单引号字符串：反斜杠不用转义，只把单引号翻倍。"""
    return "'" + str(s).replace("'", "''") + "'"


def _run_powershell(cmd, timeout=40):
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", cmd],
            capture_output=True, text=True, timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return r.returncode == 0, (r.stdout + r.stderr).strip()
    except Exception as e:
        return False, str(e)


def install_guardian_task():
    """注册任务计划程序的自愈任务：登录启动 + 每分钟兜底，挂掉自动补回。

    由 Task Scheduler 服务(svchost)托管，完全脱离任何会话 —— 这是
    「WorkBuddy 关掉再开，按钮仍在」的关键。IgnoreNew + 单例锁保证不会重复。
    """
    pyw = PYTHONW if os.path.exists(PYTHONW) else "pythonw.exe"
    cmd = (
        "$ErrorActionPreference='Stop';"
        "$py=%s; $s=%s; $d=%s; $u=\"$env:USERDOMAIN\\$env:USERNAME\";"
        "$a=New-ScheduledTaskAction -Execute $py -Argument ('\"' + $s + '\"') -WorkingDirectory $d;"
        "$t1=New-ScheduledTaskTrigger -AtLogOn -User $u;"
        "$t2=New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)"
        " -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650);"
        "$p=New-ScheduledTaskPrincipal -UserId $u -LogonType Interactive -RunLevel Limited;"
        "$st=New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries"
        " -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero);"
        "Register-ScheduledTask -TaskName %s -Action $a -Trigger @($t1,$t2)"
        " -Principal $p -Settings $st -Force | Out-Null;"
        "Start-ScheduledTask -TaskName %s"
    ) % (_ps_quote(pyw), _ps_quote(MAIN_SCRIPT), _ps_quote(APP_DIR),
         _ps_quote(TASK_NAME), _ps_quote(TASK_NAME))
    return _run_powershell(cmd)


def remove_guardian_task():
    """注销自愈任务。不动正在运行的守护，取消自启不等于立刻退出。"""
    cmd = (
        "$ErrorActionPreference='SilentlyContinue';"
        "Unregister-ScheduledTask -TaskName %s -Confirm:$false"
    ) % _ps_quote(TASK_NAME)
    return _run_powershell(cmd)


_MUTEX_HANDLE = None


def acquire_singleton():
    """常驻守护只允许一个实例；已有实例在跑时静默退出（避免双按钮 / 抢热键）。"""
    global _MUTEX_HANDLE
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p]
    k32.CreateMutexW.restype = ctypes.c_void_p
    k32.CloseHandle.argtypes = [ctypes.c_void_p]
    h = k32.CreateMutexW(None, 0, SINGLETON_NAME)
    if not h:
        return True                      # 创建失败就放行，不影响使用
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        k32.CloseHandle(h)
        return False
    _MUTEX_HANDLE = h                    # 进程存活期间一直持有
    return True


def _paste_round_frame(img, x0, y0, x1, y1, ink, r=2.2, arm=4.5, ss=4,
                       bg=(255, 255, 255)):
    """在 img 上画「四角断开的取景框」：四个独立的圆角 L，中间留缺口。

    不是一整圈边框 —— 每条边只从角上伸出一小段（arm），四边中央是空的，
    这才是截图/框选图标的画法。曲线用 4 倍超采样再缩回，边缘自然抗锯齿；
    线宽取 1.5px 等效（1.0 缩回来太淡发灰，1.8 就粗成两像素）。
    超采样画布的底必须是当前皮肤底色（bg），否则深色主题下图标会顶着白方块。
    """
    w, h = x1 - x0 + 1, y1 - y0 + 1
    big = Image.new("RGB", (w * ss, h * ss), bg)
    bd = ImageDraw.Draw(big)
    off = ss // 2                      # 对齐像素中心
    X0, Y0 = off, off
    X1, Y1 = (w - 1) * ss + off, (h - 1) * ss + off
    R, A = r * ss, arm * ss
    lw = max(1, int(round(ss * 1.5)))
    # 左上
    bd.arc([X0, Y0, X0 + 2 * R, Y0 + 2 * R], 180, 270, fill=ink, width=lw)
    bd.line([(X0 + R, Y0), (X0 + A, Y0)], fill=ink, width=lw)
    bd.line([(X0, Y0 + R), (X0, Y0 + A)], fill=ink, width=lw)
    # 右上
    bd.arc([X1 - 2 * R, Y0, X1, Y0 + 2 * R], 270, 360, fill=ink, width=lw)
    bd.line([(X1 - R, Y0), (X1 - A, Y0)], fill=ink, width=lw)
    bd.line([(X1, Y0 + R), (X1, Y0 + A)], fill=ink, width=lw)
    # 右下
    bd.arc([X1 - 2 * R, Y1 - 2 * R, X1, Y1], 0, 90, fill=ink, width=lw)
    bd.line([(X1 - R, Y1), (X1 - A, Y1)], fill=ink, width=lw)
    bd.line([(X1, Y1 - R), (X1, Y1 - A)], fill=ink, width=lw)
    # 左下
    bd.arc([X0, Y1 - 2 * R, X0 + 2 * R, Y1], 90, 180, fill=ink, width=lw)
    bd.line([(X0 + R, Y1), (X0 + A, Y1)], fill=ink, width=lw)
    bd.line([(X0, Y1 - R), (X0, Y1 - A)], fill=ink, width=lw)
    img.paste(big.resize((w, h), Image.BOX), (x0, y0))


def _rgb2hex(c):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(v)))) for v in c[:3])


def _hex2rgb(s):
    s = str(s or "").strip().lstrip("#")
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) != 6:
        return None
    try:
        return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return None


def _luma(c):
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]


def _near(a, b, tol=5):
    return all(abs(a[i] - b[i]) <= tol for i in range(3))


def _mode_color(img):
    """区域内出现最多的颜色 —— 空白底色区里它就是底色。"""
    if img is None:
        return None
    return Counter(img.getdata()).most_common(1)[0][0]


def _hover_from(bg):
    """兜底：量不到原生胶囊底时，按底色明暗压/提一档（约 5.5%）。

    实测亮色主题原生 hover 是 (240~242) 中性浅灰，跟底色 5% 的差，不是加色；
    所以这里不做任何"染色"，只动明度，深浅主题都不会跑偏。
    """
    if _luma(bg) < 128:
        return tuple(int(round(bg[i] + 0.14 * (255 - bg[i]))) for i in range(3))
    return tuple(int(round(bg[i] * 0.945)) for i in range(3))


def probe_hover_tint(rect, sidebar_w):
    """量「发现应用」那种常显胶囊的底色 —— 它就是原生 hover 底的同款颜色。

    为什么不用输入框底色推：输入框永远是白的（就算主题带颜色），推不出主题色；
    而侧栏顶部这个胶囊跟着外观主题走、且常显，实测亮色主题 = #F2F2F2，
    与工具栏图标 hover 实测 (240,240,240) / 文字按钮 hover (242,242,242) 完全一致。
    取法：侧栏底色取众数，胶囊底取「出现≥40 次、与侧栏底色差 3~40 灰阶」的下一色，
    这样自然跳过文字（差 175 灰阶）和小图标。
    """
    left, top = rect[0], rect[1]
    sb = int(sidebar_w)
    x0 = left + int(sb * 0.60)
    x1 = left + int(sb * 0.95)
    img = grab_region(x0, top + 40, x1, top + 86)
    if img is None:
        return None
    cnt = Counter(img.getdata())
    side = cnt.most_common(1)[0][0]
    for col, n in cnt.most_common(8):
        if n < 40:
            continue
        d = abs(_luma(col) - _luma(side))
        if 3 <= d <= 40:            # 比侧栏底色明显一点，但又不是文字/图标
            return col
    return None


def probe_bg(bx, by):
    """只量一小条底色：按钮正上方、输入卡片里、工具栏行上方的空白（约 1ms）。

    这块地方永远被我们的按钮「盖着但不占」—— 按钮在那里是透明的，
    所以取到的就是真实底色，用户换了深浅主题立刻反映出来。
    """
    return _mode_color(grab_region(bx + 18, by - 10, bx + 70, by - 5))


def probe_theme(bx, by, rect=None, sidebar_w=0, icon_rel=None):
    """量当前皮肤的底色 / 墨色 / hover 胶囊色。返回 None = 没量到。"""
    bg = probe_bg(bx, by)
    if bg is None:
        return None
    dark = _luma(bg) < 128
    # 墨色必须从「工具栏内部」取，而且不能跟着按钮走（按钮会随技能区变长往右跑，
    # 跑到模型名边上就会把模型名的字当成墨色，导致按钮比旁人都黑一档）。
    # 最稳的锚是连接器图标本身：工具栏在它左边必然排着「+ / 默认权限」，
    # 从这里往左取必然命中文字行。原来固定取 left+246..+576，遇到宽窗口
    # 工具栏整体右移、或深色主题下输入卡片只占中间一段时会整片落在卡片外，
    # 取到的只有背景色（深色下就是 (20,20,20)，文字看着「没变白」）。
    if rect and icon_rel:
        left = rect[0]
        x1 = left + int(icon_rel) - 8
        x0 = max(left + int(sidebar_w) + 4, x1 - 260)
    elif rect and sidebar_w:
        x0 = rect[0] + int(sidebar_w) + 6
        x1 = x0 + 330
    else:
        x0, x1 = bx - 240, bx - 24
    band = grab_region(x0, by + 4, x1, by + 26)
    ink = None
    if band is not None:
        lb = _luma(bg)
        cand = []
        for c, n in Counter(band.getdata()).items():
            if n < 10:
                continue
            if max(c) - min(c) > 100:       # 排除彩色：蓝色连接器、彩色小图标
                continue
            # 只认对比方向和底色相反的：深色主题要「比底亮」，亮色主题要「比底暗」。
            # 光用「差值够大」会把卡片外的更暗背景（深色下 20 vs 31）也算进来。
            d = _luma(c) - lb
            if d < 12 if dark else d > -12:
                continue
            cand.append((c, n))
        if cand:
            # 先按出现次数排（文字核心笔画像素最多，抗锯齿过渡色一般更少），
            # 再在前几名里取最极端的那个，避免被单个高光像素带偏。
            cand.sort(key=lambda t: -t[1])
            top = cand[:5]
            ink = (max if dark else min)(top, key=lambda t: _luma(t[0]))[0]
    if ink is None:
        ink = (235, 238, 240) if dark else (25, 57, 65)
    # hover 胶囊优先用「发现应用」那种常显胶囊的实测色（跟主题走、与原生 hover 同款）；
    # 量不到才按底色明暗压一档兜底。太贴近底色也走兜底，否则按下去看不见。
    hover = probe_hover_tint(rect, sidebar_w) if rect and sidebar_w else None
    if hover is None or abs(_luma(hover) - _luma(bg)) < 5:
        hover = _hover_from(bg)
    return {"bg": bg, "ink": ink, "dark": dark, "hover": hover}


def render_button_image(hover=False, compact=False, pal=None):
    """画出按钮的 PIL 图（不依赖 tkinter，方便离屏预览 / 深色主题模拟）。"""
    pal = pal or PAL
    bg, ink, hover_bg = pal["bg"], pal["ink"], pal["hover"]
    w = COMPACT_W if compact else BTN_W
    img = Image.new("RGB", (w, BTN_H), bg)
    d = ImageDraw.Draw(img)
    if hover:
        # 与原生胶囊同款：铺满按钮 32px 高，两端半圆
        d.rounded_rectangle((0, 0, w - 1, BTN_H - 1), radius=BTN_H // 2,
                            fill=hover_bg)

    # 四角取景框图标：四个角各自独立、中间断开（不是一整圈边框）。
    # 高度 14px、y 取 9..22，墨色中心 15.5，与文字墨色中心严格对齐，
    # 也对齐原生工具栏（原生「默认权限」文字墨色中心实测 abs 944.5）。
    # 尺寸 15x14：原生图标约这么大，17x16 会明显比旁边图标胖一圈。
    if compact:
        x0 = (w - ICON_W) // 2              # 图标水平居中
    else:
        x0 = 8
    # 图标的超采样底要用「它实际压着的那层」：hover 时是胶囊底，否则是面板底
    plate = hover_bg if hover else bg
    _paste_round_frame(img, x0, 9, x0 + ICON_W - 1, 9 + ICON_H - 1, ink, bg=plate)

    if not compact:
        try:
            font = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 13)
        except Exception:
            font = ImageFont.load_default()
        # 微软雅黑 13px 的字形偏上：anchor=lm 放在 y=16 时墨色中心 15.5，
        # 与四角框图标同一条中线，也跟原生「默认权限」(abs 944.5) 对齐。
        d.text((30, BTN_H // 2), "截图提问", font=font, fill=ink, anchor="lm")
    return img


def next_element_left(icon_right, by, bg, limit=560):
    """从连接器右缘往右扫一行，找下一个原生控件（模型名 / 技能 chip）的左缘。

    扫描行取 by+BTN_H-6：原生图标和文字都落在按钮高度内的 9~22 行，
    这一行只剩控件自己的底色，所以只命中「有底色的控件」，
    不会把字形（包括我们自己的文字）当成控件。返回绝对 x，没有则 None。
    """
    y = by + BTN_H - 6
    img = grab_region(icon_right + 2, y, icon_right + 2 + limit, y + 2)
    if img is None:
        return None
    px = img.load()
    run = 0
    for x in range(img.size[0]):
        c = px[x, 0]
        if max(abs(c[i] - bg[i]) for i in range(3)) > 4:
            run += 1
            if run >= 12:          # 连续 12px 有底色才算控件，滤掉单个像素噪点
                return icon_right + 2 + x - 11
        else:
            run = 0
    return None


def make_button_image(hover=False, compact=False, pal=None):
    """复刻输入框工具栏里的文字按钮：与面板同底（透明键）、墨色线性图标 + 文字，
    hover 才浮出一层淡色胶囊底。所有颜色都来自 pal，换主题自动跟着变。

    compact=True 时输入框已经很窄，只画四角框图标、不写字，省出 62px 空间。
    """
    return ImageTk.PhotoImage(render_button_image(hover, compact, pal))


# ---------------------------------------------------------------- 主程序
class ScreenshotQA:
    def __init__(self, once=False):
        import tkinter as tk
        self.tk = tk
        self.once = once
        self.cfg = load_config()
        # 主题配色：先按兜底值建窗口，place_button 一量到皮肤立刻换掉。
        # 透明键必须等于「底色」，否则深色主题下文字边缘会留一圈白边。
        self.pal = self.pal_with_cfg(dict(PAL))
        self.pal_ok = False          # 是否已按实测皮肤上色
        self.pal_vote = None         # 待确认的新底色（连续两次一致才换，防误判）
        self.pal_at = 0.0
        self.hover_on = False        # hover 态（轮询兜底，不依赖 Tk 的 Enter/Leave）
        self.q = queue.Queue()
        self.capturing = False
        self.btn_visible = True
        self.wb_hwnd = None
        self.hotkey_tid = 0
        self.full = None
        self.mask = None
        self.preview_item = None
        self.capture_started = 0.0
        self.last_preview = 0.0
        self.own_pid = os.getpid()
        self.startup_fg = user32.GetForegroundWindow()   # 建窗口前的焦点，稍后还回去
        self.foreign_since = 0.0            # 前台是别的程序持续多久，用于防抖
        # 右栏两态判定：矩形变了才截图检测，连续两轮一致才切换状态
        self.last_rect = None                 # 上次已复核过的窗口矩形
        self.pending_rect = None              # 正在等待重排稳定的新矩形
        self.rect_changed_at = 0.0
        self.last_panel_check = 0.0
        self.panel_left = 0                   # 右栏左缘（相对窗口；0 = 未展开）
        self.panel_vote = None
        self.panel_votes = 0
        self.last_y = None                    # 已应用的 y，用于上下去抖
        self.last_geo = None                  # 已应用的 (x, y)，相同就不再重设
        # 锚在「连接器」蓝色图标之后：icon_rel = 图标右缘相对窗口左边缘的距离
        self.icon_rel = None                  # 缓存值；None = 没测到，退回居中公式
        self.icon_cy = None                   # 图标纵向中心（绝对 y），按钮跟着它走
        # 图标中心到窗口「可见底边」的距离。工具栏是贴底布局，这个值恒定
        # （实测 1432x976 / 1150x976 / 1432x716 下都是 63.5），所以缩放高度时
        # 用 bottom - icon_gap 就能立刻算出 y，不必等精测。
        self.icon_gap = None
        self.area_w = None                    # 上次精测时的内容区宽度（缩放时按半宽增量推锚点）
        self.icon_size = None                 # 上次精测时的窗口尺寸
        self.icon_size_at = 0.0               # 该尺寸出现时刻，稳定 0.35s 后再精测
        self.icon_ready = False               # 当前尺寸是否已结束精测
        self.icon_since = 0.0                 # 上次探测时刻，用于周期性复核
        self.icon_vote = None                 # 待确认的新锚点（连续两次一致才采信）
        self.w_cur = BTN_W                    # 按钮当前宽度（窄空间下会切成图标模式）
        # ---- 跟随调度：运动中要跟得紧，静止时尽量省 ----
        # 运动中（拖动/缩放）走「零截屏」快路径，只查窗口矩形，成本约 0.1ms/轮，
        # 所以可以拉到 6ms 一轮；而 Tk 的 after 受系统定时器 15.6ms 精度限制，
        # 不提高精度的话 after(6) 实际要 16ms 才触发（这就是「慢一拍」的根因），
        # 因此仅在运动期间开 timeBeginPeriod(1)，停了立刻关掉。
        self.moving = False
        self.hp_timer = False
        self.moving_at = 0.0
        self.nxt_key = None                   # 右侧余量缓存：位置没变就不用重扫
        self.nxt_val = None
        self.recovered = restore_orphan()   # 上次异常退出遗留的隐藏窗口，先救回来

        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title("WorkBuddy 截图提问")

        self.btn = tk.Toplevel(self.root)
        self.btn.overrideredirect(True)
        self.btn.attributes("-topmost", True)
        try:
            self.btn.attributes("-toolwindow", True)
        except Exception:
            pass
        hexbg = _rgb2hex(self.pal["bg"])
        try:
            self.btn.attributes("-transparentcolor", hexbg)
        except Exception:
            pass
        self.btn.geometry("%dx%d+300+300" % (BTN_W, BTN_H))
        self.canvas = tk.Canvas(self.btn, width=BTN_W, height=BTN_H,
                                highlightthickness=0, bd=0, bg=hexbg)
        self.canvas.pack()
        self.img_normal = make_button_image(False, False, self.pal)
        self.img_hover = make_button_image(True, False, self.pal)
        self.icon_item = self.canvas.create_image(0, 0, anchor="nw", image=self.img_normal)
        self.canvas.bind("<Enter>", lambda e: self.canvas.itemconfig(self.icon_item, image=self.img_hover))
        self.canvas.bind("<Leave>", lambda e: self.canvas.itemconfig(self.icon_item, image=self.img_normal))
        self.canvas.bind("<ButtonPress-1>", self.on_btn_press)
        self.canvas.bind("<Button-3>", self.on_btn_menu)
        self.btn.after(50, self._make_toolwindow)
        self.btn_hwnd = 0
        self.btn.after(60, self._bind_btn_hwnd)

        self.menu = tk.Menu(self.root, tearoff=0)
        self.auto_var = tk.BooleanVar(value=self.cfg["auto_paste"])
        self.cfg["autostart"] = get_autostart()      # 以注册表为准
        self.start_var = tk.BooleanVar(value=self.cfg["autostart"])
        self.menu.add_command(label="截图提问    Ctrl+Alt+Q", command=self.start_capture)
        self.menu.add_checkbutton(label="截完自动粘贴到输入框", variable=self.auto_var,
                                  command=self.toggle_auto_paste)
        self.menu.add_checkbutton(label="开机自动启动", variable=self.start_var,
                                  command=self.toggle_autostart)
        self.menu.add_command(label="按钮回到默认位置", command=self.reset_position)
        self.menu.add_command(label="使用说明", command=self.show_help)
        self.menu.add_separator()
        self.menu.add_command(label="退出    Ctrl+Alt+X", command=self.quit)

        self.place_button()
        if not find_workbuddy():          # WorkBuddy 还没开：按钮先别闪一下
            self.btn.withdraw()
            self.btn_visible = False
        if not self.once:
            self.start_hotkeys()          # --once 只是单次截图，不抢全局热键
        self.root.after(100, self.poll_queue)
        self.root.after(600, self.follow_window)
        if self.recovered:
            self.root.after(700, lambda: self.toast("检测到上次异常退出，已自动恢复 WorkBuddy 窗口", 3600))
        if self.once:
            self.root.after(250, self.start_capture)

    def _make_toolwindow(self):
        apply_no_activate(self.btn)
        # Toplevel 首次映射会自己抢一次前台，把焦点还给启动前的窗口
        fg = user32.GetForegroundWindow()
        if fg and fg != self.startup_fg and window_pid(fg) == self.own_pid \
                and self.startup_fg and window_pid(self.startup_fg) != self.own_pid:
            restore_foreground(self.startup_fg)

    def show_button(self):
        """显示按钮，并把它抢走的键盘焦点还给原来的窗口。"""
        before = user32.GetForegroundWindow()
        self.btn.deiconify()
        self.btn_visible = True
        fg = user32.GetForegroundWindow()
        if fg and fg != before and window_pid(fg) == self.own_pid \
                and before and window_pid(before) != self.own_pid:
            restore_foreground(before)

    # ------------------------------------------------------------ 主题配色
    def pal_with_cfg(self, pal):
        """配置里手填的色值 / 强制明暗，优先于实测值（方便别人改主题）。"""
        p = dict(pal)
        for key, item in (("bg_color", "bg"), ("ink", "ink"), ("hover_bg", "hover")):
            v = _hex2rgb(self.cfg.get(key))
            if v:
                p[item] = v
        mode = str(self.cfg.get("theme") or "auto").lower()
        if mode in ("light", "dark"):
            p["dark"] = mode == "dark"
        return p

    def apply_palette(self, pal):
        """换皮肤：透明键、画布底色、图标与 hover 底一次性跟上。"""
        if self.pal_ok and pal == self.pal:
            return
        self.pal = dict(pal)
        self.pal_ok = True
        hexbg = _rgb2hex(self.pal["bg"])
        try:
            self.btn.attributes("-transparentcolor", hexbg)
        except Exception:
            pass
        try:
            self.canvas.config(bg=hexbg)
        except Exception:
            pass
        compact = self.w_cur < BTN_W
        self.img_normal = make_button_image(False, compact, self.pal)
        self.img_hover = make_button_image(True, compact, self.pal)
        self.canvas.itemconfig(self.icon_item, image=self.img_normal)

    def follow_theme(self, bx, by, rect=None, sidebar_w=0, force=False):
        """皮肤跟随：定期量两小块区域（输入框底色 + 侧栏常显胶囊色）。

        只看底色是不够的：换主题时输入框往往还是白的，真正变的是
        「发现应用」那种胶囊色和文字色。所以两样都盯，任一变了就整套重量。
        要连续两轮一致才采信，避免撞到输入框里的文字 / 侧栏选中的瞬时状态。
        """
        now = time.time()
        # 皮肤复核用的是独立间隔：换主题是低频操作，2.4s 内跟上完全够用，
        # 换来的是常态每秒少抓两次屏（每次 19ms，是稳态最大的一笔开销）。
        if not force and now - self.pal_at < float(self.cfg.get("theme_recheck_sec", 2.4)):
            return
        self.pal_at = now
        bg = probe_bg(bx, by)
        if bg is None:
            return
        tint = probe_hover_tint(rect, sidebar_w) if rect and sidebar_w else None
        sig = (bg, tint)
        if self.pal_ok:
            same = _near(bg, self.pal["bg"], 5) and \
                (tint is None or _near(tint, self.pal["hover"], 5))
            if same:
                self.pal_vote = None
                return
            # 见到差异：先记下来，下一轮还是这个值才真正换（防止换主题的过渡帧）
            if self.pal_vote is not None and not self._sig_same(self.pal_vote, sig):
                self.pal_vote = sig
                return
        pal = probe_theme(bx, by, rect, sidebar_w, self.icon_rel)
        self.pal_vote = sig
        if pal:
            self.apply_palette(self.pal_with_cfg(pal))

    def cursor_on_button(self):
        """鼠标是否压在按钮上（hover 时胶囊底会盖住扫描行，余量探测要跳过）。"""
        g = self.last_geo
        if not g:
            return False
        pt = wt.POINT()
        if not user32.GetCursorPos(ctypes.byref(pt)):
            return False
        x, y, w = g
        return x <= pt.x < x + w and y <= pt.y < y + BTN_H

    @staticmethod
    def _sig_same(a, b):
        if not _near(a[0], b[0], 5):
            return False
        if a[1] is None or b[1] is None:
            return True
        return _near(a[1], b[1], 5)

    # ------------------------------------------------------------ 按钮定位
    def set_button_width(self, w):
        """切换「文字 + 图标」/「只有图标」两种形态；尺寸没变就直接返回。"""
        if w == self.w_cur:
            return
        self.w_cur = w
        compact = w < BTN_W
        self.canvas.config(width=w, height=BTN_H)
        self.img_normal = make_button_image(False, compact, self.pal)
        self.img_hover = make_button_image(True, compact, self.pal)
        self.canvas.itemconfig(self.icon_item, image=self.img_normal)
        self.last_geo = None            # 逼下一轮重设 geometry（含新宽度）

    def _bind_btn_hwnd(self):
        """拿到按钮真正的顶层 HWND（Tk 的 winfo_id 是内部子窗口，父才是它）。

        Tk 的 wm geometry 是异步的：只记一笔，要等事件循环空闲才真去移动窗口，
        跟随窗口时这一等就是十几毫秒，看着就是「慢一拍」。
        所以运动中改成 SetWindowPos 直写（立即生效），geometry 仍照记，
        免得 deiconify / 改属性时 Tk 把窗口拉回它记忆里的旧位置。
        """
        try:
            self.btn_hwnd = user32.GetParent(self.btn.winfo_id()) or self.btn.winfo_id()
        except Exception:
            self.btn_hwnd = 0

    def move_to(self, x, y):
        """把按钮挪到 (x, y)；位置上没变化就直接返回。"""
        geo = (x, y, self.w_cur)
        if geo == self.last_geo:
            return
        self.last_geo = geo
        self.btn.geometry("%dx%d+%d+%d" % (self.w_cur, BTN_H, x, y))
        # 运动中直写：geometry 要等 idle 才落地，那点延迟就是「慢一拍」。
        # SetWindowPos 约 0.3ms，只在真的位移时调用，拖动时才频繁。
        if self.moving and self.btn_hwnd:
            user32.SetWindowPos(self.btn_hwnd, 0, int(x), int(y), 0, 0,
                                SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE)

    def place_button(self):
        wb = find_workbuddy()
        if wb:
            # 定位基准一律用「可见边框」：最大化时 GetWindowRect 会多算 8px 阴影，
            # 用它算 y 按钮就会比工具栏低 8px（这就是最大化上下会变的原因）。
            rect = workbuddy_rect(wb)
            left, top, right, bottom = rect
            now = time.time()
            # 「运动中」= 窗口矩形在最近 N 毫秒内变过（拖动 / 缩放 / 重排）。
            # 这一阶段只做纯几何推算：单次截屏就是 20ms 量级的同步阻塞，
            # 会把 Tk 的窗口移动整整推后一拍，肉眼看就是「按钮慢半拍」。
            hold = float(self.cfg.get("moving_hold_ms", 320)) / 1000.0
            self.moving = bool(self.rect_changed_at) and now - self.rect_changed_at < hold
            # 右栏判定只在窗口矩形变化时做；首帧直接采信，之后要连续两轮一致，
            # 避免重排动画 / 聊天气泡造成的瞬时误判把按钮来回甩。
            if rect != self.last_rect:
                if rect != self.pending_rect:
                    self.pending_rect = rect
                    self.rect_changed_at = now
                settled = self.last_rect is None or now - self.rect_changed_at >= 0.4
                if settled and now - self.last_panel_check >= 0.4:
                    edge = detect_right_panel(rect)
                    self.last_panel_check = now
                    if self.last_rect is None or edge == self.panel_vote:
                        if edge != self.panel_left:
                            # 右栏开合会压缩中间区（蓝色图标跟着挪），强制重测锚点
                            self.panel_left = edge
                            self.icon_ready = False
                            self.icon_size_at = 0.0
                        self.last_rect = rect
                        self.pending_rect = None
                    self.panel_vote = edge
            else:
                self.pending_rect = None

            sb = int(self.cfg["sidebar_w"])
            w, h = right - left, bottom - top
            size = (w, h)
            size_changed = size != self.icon_size
            prev_size = self.icon_size
            # 右栏贴着窗口右边：宽度变了 Δ，它的左缘也跟着右移 Δ。
            # 不补这一笔，下面算出的内容区宽度就不会变，锚点也不会跟着平移。
            if size_changed and self.panel_left and prev_size:
                self.panel_left += w - prev_size[0]

            # 中间内容区 = 侧栏右缘 → 右栏左缘（无右栏就是窗口右缘）
            area_l = left + sb
            area_r = (left + self.panel_left) if self.panel_left else right
            if area_r - area_l < BTN_W + 40:
                area_r = right                 # 右栏判定异常时退回整窗
            span = area_r - area_l

            # 空间不够就只留图标：内容区窄于 compact_below 时隐藏「截图提问」四个字。
            # 实测输入框宽 ≈ 内容区宽 − 50px，所以 620 对应输入框约 570px。
            # 阈值带 40px 迟滞（收 620 / 展 660）：缩放正好压在阈值线上时，
            # 不会反复在两种形态之间重绘闪一下。
            thr = int(self.cfg.get("compact_below", 620))
            narrow = span < (thr + 40 if self.w_cur == COMPACT_W else thr)
            self.set_button_width(COMPACT_W if narrow else BTN_W)
            bw = self.w_cur

            # ---- 锚点：跟在工具栏「连接器」蓝色图标之后 ----
            if size_changed:
                # 尺寸刚变：先平移旧锚点，等布局稳定再精测。首帧（还没锚点）立刻测。
                # 必须用「居中增量」而不是等比缩放：工具栏是在中间内容区居中的，
                # 内容区每宽 1px 它只右移 0.5px。实测 1432→900 时等比法偏 61px，
                # 居中增量法 0px（1432→1150 也是精确解）。
                if self.icon_rel is not None and self.area_w is not None:
                    self.icon_rel = int(round(self.icon_rel + (span - self.area_w) / 2.0))
                self.area_w = span
                self.icon_size = size
                self.icon_size_at = now if prev_size is not None else 0.0
                self.icon_ready = False
                self.icon_since = 0.0
                self.icon_vote = None
            # 复核节奏：窗口尺寸不变也要周期性复测。
            # 技能区变长 / 加了技能 chip 会把连接器图标往右推，窗口矩形却没变，
            # 只认「尺寸变化」会永远追不上，按钮就压在图标上了。
            interval = float(self.cfg.get("recheck_sec", 0.8))
            min_gap = interval if self.icon_ready else 0.2
            # 按钮不可见时（WorkBuddy 在后台 / 最小化）不复核：
            # 此时截到的可能是别的窗口，会污染锚点，而且没人看见位置。
            # 运动中也不复核：抓到的必然是被拖动的中间帧，纯浪费还会卡住跟随。
            if self.btn_visible and not self.moving and now - self.icon_since >= min_gap \
                    and now - self.icon_size_at >= 0.2:
                icon = detect_connector_icon(rect, sb)
                self.icon_since = now
                if icon is not None and icon[1] - left < w * 0.62:
                    rel, cy = icon[1] - left, icon[2]
                    if self.icon_rel is None or abs(rel - self.icon_rel) <= 2:
                        self.icon_rel, self.icon_cy = rel, cy
                        self.icon_gap = bottom - cy
                        self.icon_ready, self.icon_vote = True, None
                    elif self.icon_vote is not None and \
                            abs(rel - self.icon_vote) <= 2:
                        # 变化比较大：连续两次落到同一位置才采信，
                        # 滤掉重排动画中间帧（旧版就是被这种帧坑死的）
                        self.icon_rel, self.icon_cy = rel, cy
                        self.icon_gap = bottom - cy
                        self.icon_ready, self.icon_vote = True, None
                    else:
                        self.icon_vote = rel
                elif not self.icon_ready and now - self.icon_size_at >= 2.5:
                    self.icon_rel = None       # 一直测不到（被遮挡 / 改了主题）：退回公式
                    self.icon_ready = True

            shift = int(self.cfg.get("offset_x", 0))
            if self.icon_rel is not None:
                x = left + self.icon_rel + int(self.cfg.get("connector_gap", 12))
            else:
                x = int(round((area_l + area_r) / 2.0)) + shift - bw // 2
            if self.icon_gap is not None:
                # 工具栏贴底布局：图标中心到窗口可见底边的距离恒定（实测 63.5px，
                # 1432x976 / 1150x976 / 1432x716 完全一致）。所以用 bottom - gap
                # 反推 y，窗口高度一变就立刻跟上；只认 icon_cy 的话，拉下边缘
                # 改高度时 y 会卡在旧值不动，等精测才「跳」过去，很扎眼。
                y = int(round(bottom - self.icon_gap)) - BTN_H // 2
            elif self.icon_rel is not None and self.icon_cy is not None:
                # 还没测到 gap 时的退路：直接用图标的绝对中心
                y = int(self.icon_cy + 0.5) - BTN_H // 2
            else:
                y = bottom + int(self.cfg["offset_y"])

            # ---- 右侧余量：跟在连接器后面，但不能压到下一个原生控件 ----
            # 开了右侧栏 / 技能 chip 变多时，「连接器」后面剩下的空位可能不到
            # 92px，按固定 12px 间距摆就会把「模型名」盖住。这里量一下下一个
            # 控件的左缘，余量不够就先切成图标形态，再不够就贴着它左缘摆。
            if self.icon_rel is not None and not self.cursor_on_button():
                # 余量扫描要抓 560px 宽的条带（约 4ms）。锚点、纵向位置、
                # 按钮宽度、当前底色都没变就没必要重扫，直接复用上次结果。
                # 底色必须进 key：扫描是按「和底色差多少」判断有没有控件的，
                # 换了主题（白 → 深灰）还用旧结果的话，会把整行背景都当成控件。
                key = (self.icon_rel, y, self.w_cur, self.pal["bg"])
                if key != self.nxt_key:
                    self.nxt_key = key
                    self.nxt_val = None if self.moving else \
                        next_element_left(left + self.icon_rel, y, self.pal["bg"])
                nxt = self.nxt_val
                if nxt is None:
                    # 右边一路都是空的（测不到下一个控件）＝余量无限。
                    # 之前若为了腾地方收成了图标，现在该展开回文字。
                    if self.w_cur == COMPACT_W and not narrow:
                        self.set_button_width(BTN_W)
                        bw = self.w_cur
                else:
                    room = nxt - 6 - x
                    if room < self.w_cur:            # 太挤：先切成图标
                        self.set_button_width(COMPACT_W)
                        bw = self.w_cur
                        if room < COMPACT_W:         # 连图标都放不下：贴着左缘
                            x = max(area_l, nxt - 6 - COMPACT_W)
                    elif self.w_cur == COMPACT_W and not narrow and room >= BTN_W:
                        self.set_button_width(BTN_W)  # 空间又够了：恢复文字
                        bw = self.w_cur
            x = max(area_l, min(x, right - bw - 16))
            # 兜底：无论窗口几何多怪，按钮都不许掉出屏幕可见区
            vx, vy, vw, vh = virtual_screen()
            y = max(vy, min(y, vy + vh - BTN_H - 2))
            x = max(vx, min(x, vx + vw - bw - 2))
            # 1px 以内视为没动（滤掉静态时的逐帧抖动）。
            # 但运动中要原样跟随：死区会把连续的小位移吃掉，拖动时成了阶梯状。
            if self.moving or self.last_y is None or abs(y - self.last_y) >= 2:
                self.last_y = y
            y = self.last_y
            # 皮肤跟随：位置定了才有地方取底色（按钮透明，量到的是真实底色）。
            # 运动中不取色：截屏会阻塞，而且此刻底色也被窗口动画搅着。
            if self.btn_visible and not self.moving:
                self.follow_theme(x, y, rect, sb)
        else:
            vx, vy, vw, vh = virtual_screen()
            x, y = vx + vw - self.w_cur - 30, vy + vh - BTN_H - 120
        # 位置/尺寸没变就不重设，避免每轮 after 都触发窗口重绘（看着像"跳一下"）
        self.move_to(x, y)
        # hover 由轮询兜底：改过 transparentcolor / 重设 geometry 后，
        # Tk 的 <Enter> 有时收不到（窗口样式被重设），光靠绑定会「悬停不亮」。
        # 运动中不重算，避免拖动时按钮底色闪来闪去。
        if self.moving:
            on = self.hover_on
        else:
            on = bool(self.btn_visible and wb and self.cursor_on_button())
        if on != self.hover_on:
            self.hover_on = on
            self.canvas.itemconfig(self.icon_item,
                                   image=self.img_hover if on else self.img_normal)

    def follow_window(self):
        # 框选期间卡死 / 崩溃前的兜底：90 秒还没结束就自动恢复窗口
        if self.capturing and self.capture_started and \
                time.time() - self.capture_started > 90:
            self.abort("截图超时，已恢复 WorkBuddy 窗口")
        wb = None
        show = False
        if not self.capturing:
            wb = find_workbuddy()      # 已过滤最小化 / 关托盘 / 离屏 / cloaked
            if wb:
                # 切到别的程序时别让 topmost 按钮浮在人家窗口上；1.2s 防抖，
                # 点按钮或弹菜单造成的瞬时焦点变化不会闪一下。
                if self.cfg.get("hide_when_inactive", True) and \
                        foreground_is_foreign(wb["hwnd"], self.own_pid):
                    if not self.foreign_since:
                        self.foreign_since = time.time()
                    show = time.time() - self.foreign_since < 1.2
                else:
                    self.foreign_since = 0.0
                    show = True
                if show:
                    # 单轮出错不能让整个跟随循环停摆：pythonw 没有控制台，
                    # 异常只能落进 错误.log，否则表现成「按钮突然不动了」。
                    try:
                        self.place_button()
                    except Exception:
                        log_error("place_button")
                else:
                    self.moving = False      # 按钮收起时别让高精度定时器空转
            else:
                self.foreign_since = 0.0
            if not self.once:
                if show and not self.btn_visible:
                    self.show_button()
                elif not show and self.btn_visible:
                    self.btn.withdraw()
                    self.btn_visible = False
        # 调度分三档：
        #   运动中 6ms —— 只查窗口矩形（约 0.1ms），跟手程度和窗口本身同级；
        #   静止   120ms —— 足够发现切前台 / 按钮显隐，几乎不耗 CPU；
        #   无窗口 500ms。
        # 注意 Tk 的 after 受系统定时器精度限制（默认 15.6ms，after(6) 实际
        # 16ms 才触发），所以运动期间用 timeBeginPeriod(1) 把精度提到 1ms，
        # 一停就关掉，不长期影响系统功耗。
        if self.capturing:
            interval = 200
        elif self.moving:
            interval = int(self.cfg.get("moving_ms", 6))
        elif wb:
            interval = int(self.cfg.get("idle_ms", 16))
        else:
            interval = 500
        self.set_timer_precision(self.moving)
        if DEBUG_MOVE:
            try:
                with open(os.path.join(APP_DIR, "_move.log"), "a", encoding="utf-8") as f:
                    f.write("%.3f\t%d\t%d\t%s\n" % (time.time(), interval,
                                                    int(self.moving), self.last_geo))
            except Exception:
                pass
        try:
            self.root.after(interval, self.follow_window)
        except Exception:
            log_error("after 注册失败")

    def set_timer_precision(self, on):
        """运动中把系统定时器精度提到 1ms，否则 after(6) 实际只能跑到 16ms。

        只在拖动/缩放这几秒里开着，停稳立即恢复，避免长期拉高功耗。
        """
        on = bool(on)
        if on == self.hp_timer:
            return
        self.hp_timer = on
        try:
            (winmm.timeBeginPeriod if on else winmm.timeEndPeriod)(1)
        except Exception:
            pass

    def on_btn_press(self, e):
        # 位置完全由窗口矩形算出，不支持拖动；按下即进入框选
        self.start_capture()

    def on_btn_menu(self, e):
        try:
            self.menu.tk_popup(e.x_root, e.y_root)
        finally:
            self.menu.grab_release()

    def reset_position(self):
        for k in ("sidebar_w", "offset_x", "connector_gap", "compact_below", "offset_y"):
            self.cfg[k] = DEFAULT_CONFIG[k]
        self.panel_left = 0
        self.panel_vote = None
        self.last_rect = None
        self.pending_rect = None
        self.last_y = None
        self.last_geo = None
        self.icon_rel = None
        self.icon_cy = None
        self.icon_gap = None
        self.area_w = None
        self.icon_size = None
        self.icon_ready = False
        self.icon_size_at = 0.0
        self.icon_since = 0.0
        self.icon_vote = None
        save_config(self.cfg)
        self.place_button()

    def toggle_auto_paste(self):
        self.cfg["auto_paste"] = bool(self.auto_var.get())
        save_config(self.cfg)

    def toggle_autostart(self):
        want = bool(self.start_var.get())
        ok = set_autostart(want)
        if not ok:
            self.start_var.set(not want)
            self.toast("设置开机自启失败（权限不足）")
            return
        self.cfg["autostart"] = want
        save_config(self.cfg)
        # 注册表项与计划任务一起同步；放后台线程，别卡住右键菜单
        threading.Thread(target=install_guardian_task if want else remove_guardian_task,
                         daemon=True, name="guardian-task").start()

    # ------------------------------------------------------------ 热键
    def start_hotkeys(self):
        def loop():
            tid = kernel32.GetCurrentThreadId()
            self.hotkey_tid = tid
            ok1 = user32.RegisterHotKey(None, 1, MOD_CONTROL | MOD_ALT, ord("Q"))
            ok2 = user32.RegisterHotKey(None, 2, MOD_CONTROL | MOD_ALT, ord("X"))
            if not (ok1 and ok2):
                self.q.put(("warn", "全局热键注册失败，可直接点击悬浮按钮"))
            msg = MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == WM_HOTKEY:
                    if msg.wParam == 1:
                        self.q.put(("capture", None))
                    elif msg.wParam == 2:
                        self.q.put(("quit", None))
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))

        threading.Thread(target=loop, daemon=True, name="hotkey").start()

    def poll_queue(self):
        try:
            while True:
                cmd, arg = self.q.get_nowait()
                if cmd == "capture":
                    self.start_capture()
                elif cmd == "quit":
                    self.quit()
                elif cmd == "warn":
                    self.toast(arg, 3200)
        except queue.Empty:
            pass
        self.root.after(120, self.poll_queue)

    # ------------------------------------------------------------ 提示气泡
    def toast(self, text, ms=2400):
        tk = self.tk
        try:
            t = tk.Toplevel(self.root)
            t.overrideredirect(True)
            t.attributes("-topmost", True)
            lbl = tk.Label(t, text=text, bg="#1b2130", fg="#f4f7ff",
                           font=("Microsoft YaHei UI", 10), padx=16, pady=9)
            lbl.pack()
            t.update_idletasks()
            apply_no_activate(t)
            vx, vy, vw, vh = virtual_screen()
            w, h = t.winfo_width(), t.winfo_height()
            t.geometry("+%d+%d" % (vx + (vw - w) // 2, vy + vh - h - 90))
            t.after(ms, t.destroy)
        except Exception:
            pass

    def show_help(self):
        from tkinter import messagebox
        messagebox.showinfo(
            "截图提问 · 使用说明",
            "1. 点击输入框旁的「截图提问」按钮，或按 Ctrl+Alt+Q\n"
            "2. WorkBuddy 会自动隐藏，鼠标拖出要截的区域（Esc 取消）\n"
            "3. 松手后窗口恢复，截图自动变成输入框里的附件\n"
            "4. 直接输入你的问题，回车发送\n\n"
            "· 按钮可以直接拖动，位置会被记住\n"
            "· 装了开机自启后，本程序常驻后台；WorkBuddy 一开按钮就出现，\n"
            "  最小化、关掉或切到别的程序时按钮跟着隐藏，不用再手动启动\n"
            "· 右键按钮：自动粘贴开关 / 开机自启 / 复位 / 退出\n"
            "· 截图同时保存在本目录 截图\\ 下，latest.png 为最近一张\n"
            "· 正常截图不弹任何提示；只有出错才提示，此时图片仍在剪贴板，\n"
            "  点输入框按 Ctrl+V 即可\n"
            "· 卸载：先双击 取消自启.bat（注销开机启动项与自愈任务），\n"
            "  按 Ctrl+Alt+X 退出守护，再删除本文件夹\n"
            "· 自愈任务由系统任务计划服务托管，与 WorkBuddy 关开无关；\n"
            "  守护万一挂掉，1 分钟内会自动补回")

    # ------------------------------------------------------------ 截图流程
    def start_capture(self):
        if self.capturing:
            return
        self.capturing = True
        self.capture_started = time.time()
        wb = find_workbuddy()
        self.wb_hwnd = wb["hwnd"] if wb else None
        self.btn.withdraw()
        if self.wb_hwnd:
            mark_hidden(True)
            user32.ShowWindow(self.wb_hwnd, SW_HIDE)
        time.sleep(0.20)          # 等窗口真正消失再抓屏
        try:
            self.full = ImageGrab.grab(bbox=None, all_screens=True).convert("RGB")
        except Exception as e:
            self.abort("抓屏失败：%s" % e)
            return
        self.open_mask()

    def abort(self, reason=""):
        if self.mask is not None:
            try:
                self.mask.grab_release()
                self.mask.destroy()
            except Exception:
                pass
            self.mask = None
        self.restore_window()
        self.full = None
        self.capturing = False
        if not self.once:
            self.show_button()
        if reason:
            self.toast(reason, 3200)
        if self.once:
            self.root.after(400, self.quit)

    def open_mask(self):
        tk = self.tk
        vx, vy, vw, vh = virtual_screen()
        full = self.full
        dim = Image.blend(full, Image.new("RGB", full.size, (10, 14, 22)), 0.42)

        m = tk.Toplevel(self.root)
        m.overrideredirect(True)
        m.attributes("-topmost", True)
        m.geometry("%dx%d+%d+%d" % (vw, vh, vx, vy))
        c = tk.Canvas(m, width=vw, height=vh, highlightthickness=0,
                      bd=0, cursor="crosshair")
        c.pack()
        self.dim_photo = ImageTk.PhotoImage(dim)
        c.create_image(0, 0, anchor="nw", image=self.dim_photo)
        self.preview_photo = None
        self.preview_item = None
        self.rect_item = c.create_rectangle(0, 0, 0, 0, outline="#ffffff", width=2)
        self.size_item = c.create_text(0, 0, anchor="sw", text="", fill="#ffffff",
                                       font=("Microsoft YaHei UI", 10))
        self.mask, self.mask_canvas = m, c
        self.vx, self.vy = vx, vy
        self.sel_start = None

        c.bind("<ButtonPress-1>", self.on_sel_press)
        c.bind("<B1-Motion>", self.on_sel_move)
        c.bind("<ButtonRelease-1>", self.on_sel_release)
        m.bind("<Escape>", lambda e: self.abort())
        c.bind("<Button-3>", lambda e: self.abort())
        m.after(120, self._focus_mask)

    def _focus_mask(self):
        if self.mask is None:
            return
        try:
            self.mask.focus_force()
            self.mask.grab_set()
        except Exception:
            pass

    def on_sel_press(self, e):
        self.sel_start = (e.x, e.y)

    def on_sel_move(self, e):
        if not self.sel_start:
            return
        x0, y0 = self.sel_start
        x1, y1 = max(0, min(e.x, self.full.width)), max(0, min(e.y, self.full.height))
        a, b = min(x0, x1), min(y0, y1)
        c, d = max(x0, x1), max(y0, y1)
        cv = self.mask_canvas
        now = time.time()
        if c - a > 2 and d - b > 2 and now - self.last_preview > 0.045:
            crop = self.full.crop((a, b, c, d))
            self.preview_photo = ImageTk.PhotoImage(crop)
            if self.preview_item is None:
                self.preview_item = cv.create_image(a, b, anchor="nw", image=self.preview_photo)
            else:
                cv.itemconfig(self.preview_item, image=self.preview_photo)
                cv.coords(self.preview_item, a, b)
            self.last_preview = now
        cv.coords(self.rect_item, a, b, c, d)
        cv.itemconfig(self.size_item, text="%d × %d" % (c - a, d - b))
        cv.coords(self.size_item, a + 6, max(14, b - 6))
        cv.tag_raise(self.rect_item)
        cv.tag_raise(self.size_item)

    def on_sel_release(self, e):
        if not self.sel_start:
            self.abort()
            return
        x0, y0 = self.sel_start
        a, b = min(x0, e.x), min(y0, e.y)
        c, d = max(x0, e.x), max(y0, e.y)
        a, b = max(0, a), max(0, b)
        c, d = min(self.full.width, c), min(self.full.height, d)
        if c - a < 8 or d - b < 8:
            self.abort()
            return
        try:
            self.mask.grab_release()
            self.mask.destroy()
        except Exception:
            pass
        self.mask = None
        shot = self.full.crop((a, b, c, d))
        self.full = None
        self.finish(shot)

    def finish(self, shot):
        # 1) 存盘
        try:
            os.makedirs(SHOT_DIR, exist_ok=True)
            name = "shot_%s.png" % time.strftime("%Y%m%d_%H%M%S")
            shot.save(os.path.join(SHOT_DIR, name))
            shot.save(os.path.join(SHOT_DIR, "latest.png"))
        except Exception:
            pass

        # 2) 写入剪贴板
        ok, msg = copy_image_to_clipboard(shot)
        if not ok:
            self.abort("剪贴板写入失败：%s" % msg)
            return

        # 3) 恢复 WorkBuddy
        self.restore_window()
        self.capturing = False
        if not self.once:
            self.show_button()

        # 4) 自动粘贴成附件（延时等窗口拿回焦点）。
        #    成功路径完全静默：只有剪贴板写入失败、窗口/焦点没恢复才弹提示。
        if self.cfg["auto_paste"]:
            self.root.after(380, self._do_paste)
        if self.once:
            self.root.after(1200, self.quit)

    def restore_window(self):
        mark_hidden(False)
        if self.wb_hwnd:
            try:
                h = self.wb_hwnd
                # 被最小化过就先用 SW_RESTORE：对最小化窗口 SW_SHOW 只会让它继续
                # 缩在任务栏，而 IsWindowVisible 仍返回 True —— 后面那一下点击就会
                # 落到桌面上（这就是「点开了桌面文件」的其中一条路）。
                if user32.IsIconic(h):
                    user32.ShowWindow(h, SW_RESTORE)
                else:
                    user32.ShowWindow(h, SW_SHOW)
                user32.keybd_event(VK_MENU, 0, 0, 0)      # Alt trick，绕过前台锁定
                user32.SetForegroundWindow(h)
                user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
                user32.BringWindowToTop(h)
                # 再顶一次到最上层：只靠 SetForegroundWindow 时，别的程序的
                # topmost 窗口仍可能压在它前面。
                user32.SetWindowPos(h, HWND_TOP, 0, 0, 0, 0,
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
            except Exception:
                pass

    def _input_click_point(self, left, top, right, bottom):
        """算一个「输入框正文里的空白点」，用来把键盘焦点点回输入框。

        为什么需要：用户可能没点进输入框就直接截图（点悬浮按钮 / 按热键），
        窗口虽然回到前台，焦点却落在页面空白处，Ctrl+V 没接收者 —— 表现就是
        「图截好了、剪贴板也有，可输入框里什么也没出现」。

        为什么不写死像素：输入框正文的高度随内容变（空框最矮，贴图 / 多行会撑高）。
        这里从工具栏往上扫一条窄带，找「离工具栏最近的一段纯底色」，点它的正中：
          · 空框：落在占位文字与工具栏之间的空白；
          · 有附件缩略图：落在缩略图下方，不会点到缩略图上（点它会弹预览）；
          · 有草稿：落在最后一行下方，光标顺势停在末尾，最接近手动操作。
        扫不到就退回「工具栏上方 58px」的经验值。
        """
        if self.icon_gap is None and self.icon_cy is None:
            return None                # 锚点还没量到：宁可不点，也不瞎点
        if self.icon_gap is not None:
            cy = bottom - self.icon_gap
        else:
            cy = self.icon_cy
        if self.icon_rel is not None:
            x = left + self.icon_rel - 40      # 连接器图标左侧的正文区（不会压在按钮上）
        else:
            x = int((left + int(self.cfg["sidebar_w"]) + right) / 2.0)
        x = int(max(left + 8, min(x, right - 8)))
        bg = self.pal.get("bg") or (255, 255, 255)
        # 扫描窗只取工具栏上方 100px 内：底面积不能往上贪，否则会扫到输入框外的
        # 页面底色（浅色主题下两者只差 7 个灰阶），空白带一拉长中心点就跑出框外了。
        y0, y1 = int(cy) - 100, int(cy) - 26
        y = None
        im = grab_region(x - 1, y0, x + 2, y1)
        if im is not None:
            px = im.load()
            h = im.size[1]
            j = h - 1
            while j >= 0:
                k = j
                while k >= 0:
                    c = px[1, k][:3]
                    if not _near(c, bg, 4) or (max(c) - min(c)) > 6:
                        break                    # 4 个灰阶：页面底和输入框底要区分开
                    k -= 1
                if j - k >= 12:                     # 至少 12px 高的纯底色才算一片空白
                    y = y0 + (k + 1 + j) // 2
                    break
                j = k - 1 if k < j else j - 1
        # 扫不到一片确定的空白就返回 None（不瞎点）：宁可退回「直接 Ctrl+V」，
        # 也不拿一个靠猜的坐标去点 —— 猜错就是点到聊天里的文件卡片 / 别的控件。
        return (x, int(y)) if y is not None else None

    def _window_at(self, x, y):
        """(x, y) 下面那层窗口的顶层句柄 + 进程号 —— 点之前用它确认落点归属。"""
        pt = wt.POINT(int(x), int(y))
        hwnd = user32.WindowFromPoint(pt)
        if not hwnd:
            return None, 0
        root_ = user32.GetAncestor(hwnd, GA_ROOT) or hwnd
        return root_, window_pid(root_)

    def _wait_target(self, x, y, ms=800):
        """等这一点真的被 WorkBuddy 的窗口盖住再下手。

        为什么必须等：恢复窗口那一瞬间 foreground 已经是 WorkBuddy 了，可 DWM 还在
        做显示动画，屏幕上此刻压着这一点的仍是别的程序（桌面 / 上一个应用）。
        这时候点下去就是点别人 —— 实测把桌面图标点开过。
        """
        if not self.wb_hwnd:
            return False
        pid = window_pid(self.wb_hwnd)
        t0 = time.time()
        while True:
            hwnd, p = self._window_at(x, y)
            if hwnd == self.wb_hwnd or (p and p == pid):
                return True
            if (time.time() - t0) * 1000.0 >= ms:
                return False
            time.sleep(0.05)

    def _focus_input_box(self):
        """鼠标点一下输入框正文，把键盘焦点交进去（点完把指针原样放回去）。"""
        wb = find_workbuddy()
        if not wb:
            return False
        left, top, right, bottom = workbuddy_rect(wb)
        pt = self._input_click_point(left, top, right, bottom)
        if pt is None:
            log_diag("跳过：锚点未就绪")
            return False
        x, y = pt
        if not (left + 2 <= x <= right - 2 and top + 2 <= y <= bottom - 2):
            log_diag("跳过：落点出界 (%d,%d) rect=%s" % (x, y, (left, top, right, bottom)))
            return False
        # 落点必须是 WorkBuddy 自己的窗口。等不到就不点：宁可这次没粘上，
        # 也不能把这一下点击送给桌面或别的程序。
        if not self._wait_target(x, y):
            h, p = self._window_at(x, y)
            log_diag("跳过：落点被占用 (%d,%d) 上面是 %s" % (x, y, exe_name(p) if p else "?"))
            return False
        cur = wt.POINT()
        if not user32.GetCursorPos(ctypes.byref(cur)):
            return False
        ox, oy = cur.x, cur.y
        user32.SetCursorPos(x, y)
        # 指针没挪过去（被别的程序锁住 / ClipCursor）就绝不能点：
        # 那样这一下会落在用户当前鼠标所在的位置上。
        got = wt.POINT()
        user32.GetCursorPos(ctypes.byref(got))
        if abs(got.x - x) > 1 or abs(got.y - y) > 1:
            log_diag("跳过：指针挪不动，要到 (%d,%d) 实际 (%d,%d)" % (x, y, got.x, got.y))
            return False
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        log_diag("点击 (%d,%d) rect=%s" % (x, y, (left, top, right, bottom)))

        def _give_back():
            # 指针还停在我们放的位置才还原：这期间用户自己动过鼠标就别抢回来
            p = wt.POINT()
            if user32.GetCursorPos(ctypes.byref(p)) and \
                    abs(p.x - x) <= 2 and abs(p.y - y) <= 2:
                user32.SetCursorPos(ox, oy)
        self.root.after(70, _give_back)
        return True

    def _do_paste(self):
        if not self.wb_hwnd or not user32.IsWindowVisible(self.wb_hwnd) \
                or user32.IsIconic(self.wb_hwnd):
            log_diag("放弃：窗口状态异常 visible=%s iconic=%s"
                     % (bool(self.wb_hwnd and user32.IsWindowVisible(self.wb_hwnd)),
                        bool(self.wb_hwnd and user32.IsIconic(self.wb_hwnd))))
            self.toast("WorkBuddy 窗口未恢复，图片在剪贴板，请手动 Ctrl+V", 3600)
            return
        # 焦点没回到 WorkBuddy 就不盲发按键，避免把图粘到别的程序里
        fg = user32.GetForegroundWindow()
        if fg and fg != self.wb_hwnd and user32.GetParent(fg) != self.wb_hwnd:
            log_diag("放弃：前台不是 WorkBuddy fg=%s" % fg)
            self.toast("窗口焦点未恢复：图片在剪贴板，点输入框按 Ctrl+V", 4200)
            return
        # 先把焦点点进输入框，再发 Ctrl+V（点不到就照旧直接发）
        if self.cfg.get("click_before_paste", True):
            try:
                if self._focus_input_box():
                    self.root.after(110, self._send_paste)
                    return
            except Exception:
                log_error("focus_input_box")
        self._send_paste()

    def _send_paste(self):
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        user32.keybd_event(VK_V, 0, 0, 0)
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        log_diag("已发 Ctrl+V")

    # ------------------------------------------------------------ 退出
    def quit(self):
        try:
            if self.hotkey_tid:
                user32.PostThreadMessageW(self.hotkey_tid, WM_QUIT, 0, 0)
        except Exception:
            pass
        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass
        # 热键线程阻塞在 GetMessageW，主线程直接退出即可
        os._exit(0)

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------- 命令行入口
def selftest():
    print("python :", sys.executable)
    print("pillow :", Image.__version__)
    print("tkinter:", __import__("tkinter").TkVersion)
    vx, vy, vw, vh = virtual_screen()
    print("virtual screen: %d,%d  %dx%d" % (vx, vy, vw, vh))
    wb = find_workbuddy()
    print("workbuddy window:", (wb["title"], wb["exe"], wb["rect"]) if wb else "NOT FOUND")
    print("window active:", window_is_active(wb["hwnd"]) if wb else "n/a",
          "| foreground foreign:",
          foreground_is_foreign(wb["hwnd"], os.getpid()) if wb else "n/a")
    img = Image.new("RGB", (4, 3), (200, 30, 30))
    dib = build_dib(img)
    print("CF_DIB bytes:", len(dib), "(expect %d)" % (40 + 4 * 3 * 4))
    one = ImageGrab.grab(bbox=(vx, vy, vx + 8, vy + 8))
    print("grab test:", one.size, one.mode)
    print("config:", CONFIG_PATH)
    print("hidden flag:", os.path.exists(HIDE_FLAG))
    print("autostart:", get_autostart(), "|", autostart_command())
    cfg = load_config()
    if wb:
        vrect = workbuddy_rect(wb)
        left, _, right, bottom = vrect
        sb = int(cfg["sidebar_w"])
        edge = detect_right_panel(vrect)
        area_l = left + sb
        area_r = (left + edge) if edge else right
        shift = int(cfg.get("offset_x", 0))
        bx = int(round((area_l + area_r) / 2.0)) + shift - BTN_W // 2
        by = bottom + int(cfg["offset_y"])
        icon = detect_connector_icon(vrect, sb)
        if icon:
            bx = icon[1] + int(cfg.get("connector_gap", 12))
            by = int(icon[2] + 0.5) - BTN_H // 2
        print("locate: GetWindowRect=%s visible=%s" % (wb["rect"], vrect))
        print("locate: panel_left=%s area=%d..%d connector=%s gap=%s "
              "button=(%d,%d)" %
              (edge, area_l, area_r, icon, cfg.get("connector_gap"), bx, by))
        pal = probe_theme(bx, by, vrect, sb,
                          (icon[1] - left) if icon else None)
        if pal:
            print("theme: bg=%s ink=%s hover=%s %s" %
                  (_rgb2hex(pal["bg"]), _rgb2hex(pal["ink"]),
                   _rgb2hex(pal["hover"]), "dark" if pal["dark"] else "light"))
        else:
            print("theme: probe failed (按钮被遮挡 / 窗口不可见)")
    print("SELFTEST OK")


def main():
    if "--selftest" in sys.argv:
        selftest()
        return
    if "--install-autostart" in sys.argv or "--remove-autostart" in sys.argv:
        enable = "--install-autostart" in sys.argv
        ok = set_autostart(enable)
        if ok:
            cfg = load_config()
            cfg["autostart"] = enable
            save_config(cfg)
        tok, tmsg = install_guardian_task() if enable else remove_guardian_task()
        if not ok:
            print("设置失败（权限不足或注册表不可写）", flush=True)
        elif enable:
            print("已开启开机自启：%s" % autostart_command(), flush=True)
            print("守护自愈任务：%s" % ("已注册并启动" if tok else "注册失败 " + tmsg), flush=True)
            if not tok:      # 任务不可用时才退回 detached 启动，避免两种启动方式抢单例锁
                print("兜底启动：%s" % ("成功" if launch_now() else "未能启动，双击 启动.vbs 即可"),
                      flush=True)
        else:
            print("已取消开机自启；自愈任务注销：%s" % ("成功" if tok else tmsg), flush=True)
            print("当前仍在运行的守护可按 Ctrl+Alt+X 退出，或重启后不再自动启动", flush=True)
        return
    if "--lock-test" in sys.argv:
        got = acquire_singleton()
        print("lock acquired:", got, flush=True)
        if got:
            time.sleep(3)
        return
    if "--restore" in sys.argv:
        wb = force_restore()
        print("restored:", wb["title"] if wb else "no WorkBuddy window found")
        return
    if "--clip-test" in sys.argv:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        old = None
        try:
            old = root.clipboard_get()      # 尽量保住原来的文本内容
        except Exception:
            pass
        ok, msg = copy_image_to_clipboard(Image.new("RGB", (8, 8), (220, 40, 40)))
        back = ImageGrab.grabclipboard()
        print("write:", ok, msg)
        print("readback:", isinstance(back, Image.Image),
              back.convert("RGB").getpixel((4, 4)) if isinstance(back, Image.Image) else None)
        if old is not None:
            root.clipboard_clear()
            root.clipboard_append(old)
            root.update()
        root.destroy()
        return
    if any(a.startswith("--smoke") for a in sys.argv):
        secs = 1.2
        for a in sys.argv:
            if a.startswith("--smoke="):
                try:
                    secs = float(a.split("=", 1)[1])
                except ValueError:
                    pass
        app = ScreenshotQA(once=False)
        app.root.update_idletasks()
        app.btn.update_idletasks()
        print("UI OK | button at %d,%d | hotkey thread=%s | hold %.1fs" %
              (app.btn.winfo_x(), app.btn.winfo_y(), bool(app.hotkey_tid), secs), flush=True)
        app.root.after(int(secs * 1000), app.quit)
        app.run()
        return
    once = "--once" in sys.argv
    if not once and not acquire_singleton():
        return                       # 已有守护实例在跑，静默退出
    app = ScreenshotQA(once=once)
    app.run()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        os.makedirs(APP_DIR, exist_ok=True)
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write("\n==== %s ====\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
            f.write(traceback.format_exc())
        raise
