# -*- coding: utf-8 -*-
import ctypes
import time
import threading
import sys

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)

# ============================================================
# **Win32 types (self-defined, no reliance on ctypes.wintypes' missing names)**
# ============================================================
VOID_P = ctypes.c_void_p
HANDLE = VOID_P

HWND = HANDLE
HICON = HANDLE
HCURSOR = HANDLE
HBRUSH = HANDLE
HMENU = HANDLE
HINSTANCE = HANDLE
HHOOK = HANDLE
LPVOID = VOID_P

BOOL = ctypes.c_int
UINT = ctypes.c_uint
DWORD = ctypes.c_uint32
WORD = ctypes.c_uint16
LONG = ctypes.c_int32

WCHAR = ctypes.c_wchar
LPCWSTR = ctypes.c_wchar_p
LPWSTR = ctypes.c_wchar_p

kernel32.GetCurrentProcessId.argtypes = []
kernel32.GetCurrentProcessId.restype = DWORD
user32.UnregisterClassW.argtypes = [LPCWSTR, HINSTANCE]
user32.UnregisterClassW.restype = BOOL

_tray_class_counter = 0

# pointer-size dependent integer types
if ctypes.sizeof(VOID_P) == 8:
    WPARAM = ctypes.c_uint64
    LPARAM = ctypes.c_int64
    LRESULT = ctypes.c_int64
    UINT_PTR = ctypes.c_uint64
else:
    WPARAM = ctypes.c_uint32
    LPARAM = ctypes.c_int32
    LRESULT = ctypes.c_int32
    UINT_PTR = ctypes.c_uint32

ATOM = WORD

WNDPROC = ctypes.WINFUNCTYPE(LRESULT, HWND, UINT, WPARAM, LPARAM)
EnumWindowsProc = ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)

# ============================================================
# **Structures**
# ============================================================
class POINT(ctypes.Structure):
    _fields_ = [("x", LONG), ("y", LONG)]

class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", HWND),
        ("message", UINT),
        ("wParam", WPARAM),
        ("lParam", LPARAM),
        ("time", DWORD),
        ("pt", POINT),
    ]

class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", HINSTANCE),
        ("hIcon", HICON),
        ("hCursor", HCURSOR),
        ("hbrBackground", HBRUSH),
        ("lpszMenuName", LPCWSTR),
        ("lpszClassName", LPCWSTR),
    ]

class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", DWORD),
        ("hWnd", HWND),
        ("uID", UINT),
        ("uFlags", UINT),
        ("uCallbackMessage", UINT),
        ("hIcon", HICON),
        ("szTip", WCHAR * 128),
    ]

# ============================================================
# **Win32 API prototypes**
# ============================================================
# window list / title
user32.EnumWindows.argtypes = [EnumWindowsProc, LPARAM]
user32.EnumWindows.restype = BOOL

user32.GetWindowTextLengthW.argtypes = [HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int

user32.GetWindowTextW.argtypes = [HWND, LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.IsWindowVisible.argtypes = [HWND]
user32.IsWindowVisible.restype = BOOL

user32.GetShellWindow.argtypes = []
user32.GetShellWindow.restype = HWND

user32.SetWindowTextW.argtypes = [HWND, LPCWSTR]
user32.SetWindowTextW.restype = BOOL

user32.IsWindow.argtypes = [HWND]
user32.IsWindow.restype = BOOL

# console window control
kernel32.GetConsoleWindow.argtypes = []
kernel32.GetConsoleWindow.restype = HWND

user32.ShowWindow.argtypes = [HWND, ctypes.c_int]
user32.ShowWindow.restype = BOOL

kernel32.SetConsoleTitleW.argtypes = [LPCWSTR]
kernel32.SetConsoleTitleW.restype = BOOL

# tray + hidden message window
user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
user32.RegisterClassW.restype = ATOM

user32.CreateWindowExW.argtypes = [
    DWORD, LPCWSTR, LPCWSTR, DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    HWND, HMENU, HINSTANCE, LPVOID
]
user32.CreateWindowExW.restype = HWND

user32.DefWindowProcW.argtypes = [HWND, UINT, WPARAM, LPARAM]
user32.DefWindowProcW.restype = LRESULT

user32.DestroyWindow.argtypes = [HWND]
user32.DestroyWindow.restype = BOOL

user32.PostQuitMessage.argtypes = [ctypes.c_int]
user32.PostQuitMessage.restype = None

user32.GetMessageW.argtypes = [ctypes.POINTER(MSG), HWND, UINT, UINT]
user32.GetMessageW.restype = BOOL

user32.TranslateMessage.argtypes = [ctypes.POINTER(MSG)]
user32.TranslateMessage.restype = BOOL

user32.DispatchMessageW.argtypes = [ctypes.POINTER(MSG)]
user32.DispatchMessageW.restype = LRESULT

user32.CreatePopupMenu.argtypes = []
user32.CreatePopupMenu.restype = HMENU

user32.AppendMenuW.argtypes = [HMENU, UINT, UINT_PTR, LPCWSTR]
user32.AppendMenuW.restype = BOOL

user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
user32.GetCursorPos.restype = BOOL

user32.SetForegroundWindow.argtypes = [HWND]
user32.SetForegroundWindow.restype = BOOL

user32.TrackPopupMenu.argtypes = [HMENU, UINT, ctypes.c_int, ctypes.c_int, ctypes.c_int, HWND, LPVOID]
user32.TrackPopupMenu.restype = UINT

shell32.Shell_NotifyIconW.argtypes = [DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
shell32.Shell_NotifyIconW.restype = BOOL

user32.LoadIconW.argtypes = [HINSTANCE, LPCWSTR]
user32.LoadIconW.restype = HICON

kernel32.GetModuleHandleW.argtypes = [LPCWSTR]
kernel32.GetModuleHandleW.restype = HINSTANCE

# ============================================================
# **Constants**
# ============================================================
SW_HIDE = 0
SW_SHOW = 5

WM_APP = 0x8000
WM_TRAYICON = WM_APP + 1
WM_COMMAND = 0x0111
WM_CLOSE = 0x0010
WM_DESTROY = 0x0002

NIM_ADD = 0x00000000
NIM_DELETE = 0x00000002

NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004

IDI_APPLICATION = 32512

TPM_RIGHTBUTTON = 0x0002
TPM_NONOTIFY = 0x0080
TPM_RETURNCMD = 0x0100

ID_TRAY_SHOW = 1001
ID_TRAY_EXIT = 1002

def _get_last_error():
    return ctypes.get_last_error()

def make_int_resource(i: int) -> LPCWSTR:
    # MAKEINTRESOURCEW
    return ctypes.cast(ctypes.c_void_p(i), LPCWSTR)

# ============================================================
# **Core: enumerate windows + set title**
# ============================================================
def get_window_title(hwnd: int) -> str:
    ln = user32.GetWindowTextLengthW(HWND(hwnd))
    if ln <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(ln + 1)
    user32.GetWindowTextW(HWND(hwnd), buf, len(buf))
    return buf.value.strip()

def list_open_windows():
    shell = user32.GetShellWindow()
    windows = []

    @EnumWindowsProc
    def callback(hwnd, lparam):
        if hwnd == shell:
            return True
        if not user32.IsWindowVisible(hwnd):
            return True
        title = get_window_title(int(hwnd))
        if not title:
            return True
        windows.append((int(hwnd), title))
        return True

    ok = user32.EnumWindows(callback, 0)
    if not ok:
        raise OSError(f"EnumWindows failed, GetLastError={_get_last_error()}")
    windows.sort(key=lambda x: (x[1].lower(), x[0]))
    return windows

def set_window_title(hwnd: int, new_title: str) -> bool:
    return bool(user32.SetWindowTextW(HWND(hwnd), new_title))

# ============================================================
# **Persistent renamer (keep titles)**
# ============================================================
class PersistentRenamer:
    def __init__(self):
        self._lock = threading.Lock()
        self._rules = {}  # hwnd(int) -> title(str)
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._worker, daemon=True)

    def start(self):
        self._thread.start()

    def stop(self):
        self._stop.set()
        self._thread.join(timeout=2)

    def add_or_update(self, hwnd: int, title: str):
        with self._lock:
            self._rules[hwnd] = title

    def remove(self, hwnd: int):
        with self._lock:
            self._rules.pop(hwnd, None)

    def list_rules(self):
        with self._lock:
            return dict(self._rules)

    def _worker(self):
        while not self._stop.is_set():
            time.sleep(1)
            with self._lock:
                items = list(self._rules.items())
            if not items:
                continue
            dead = []
            for hwnd, title in items:
                if not user32.IsWindow(HWND(hwnd)):
                    dead.append(hwnd)
                    continue
                user32.SetWindowTextW(HWND(hwnd), title)
            if dead:
                with self._lock:
                    for hwnd in dead:
                        self._rules.pop(hwnd, None)

# ============================================================
# **Tray controller**
# ============================================================
class TrayController:
    """
    - 隐藏 console
    - 创建隐藏消息窗口 + 托盘图标
    - 右键/左键弹出菜单：Show / Exit
    """
    def __init__(self, tooltip="Window Title Renamer"):
        self.tooltip = tooltip
        self.hwnd_msg = None
        self._should_exit = False
        self._should_show = False
        self._wndproc = WNDPROC(self._wndproc_impl)

    def _wndproc_impl(self, hwnd, msg, wparam, lparam):
        if msg == WM_TRAYICON:
            # WM_RBUTTONUP=0x0205, WM_LBUTTONUP=0x0202
            if int(lparam) in (0x0205, 0x0202):
                self._show_context_menu(hwnd)
            return 0

        if msg == WM_COMMAND:
            cmd_id = int(wparam) & 0xFFFF
            if cmd_id == ID_TRAY_SHOW:
                self._should_show = True
                user32.PostQuitMessage(0)
                return 0
            if cmd_id == ID_TRAY_EXIT:
                self._should_exit = True
                user32.PostQuitMessage(0)
                return 0

        if msg in (WM_CLOSE, WM_DESTROY):
            user32.PostQuitMessage(0)
            return 0

        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _show_context_menu(self, hwnd):
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, 0x0000, ID_TRAY_SHOW, "Show (回到前台)")
        user32.AppendMenuW(menu, 0x0000, ID_TRAY_EXIT, "Exit (退出)")

        pt = POINT()
        user32.GetCursorPos(ctypes.byref(pt))

        # 关键：弹菜单前把自己设为前台窗口，否则菜单/命令经常不稳定
        user32.SetForegroundWindow(hwnd)

        # 关键：TPM_RETURNCMD 会让 TrackPopupMenu 返回命令ID，而不是发 WM_COMMAND
        cmd = user32.TrackPopupMenu(
            menu,
            TPM_RIGHTBUTTON | TPM_RETURNCMD,  # 去掉 TPM_NONOTIFY（不需要）
            pt.x, pt.y,
            0,
            hwnd,
            None
        )

        # cmd == 0 表示用户点空白处取消
        if cmd == ID_TRAY_SHOW:
            self._should_show = True
            user32.PostQuitMessage(0)
        elif cmd == ID_TRAY_EXIT:
            self._should_exit = True
            user32.PostQuitMessage(0)

        # 常见托盘菜单“小坑”：需要发一个空消息让菜单正确收尾
        user32.DefWindowProcW(hwnd, 0, 0, 0)

    def _add_icon(self):
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd_msg
        nid.uID = 1
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        nid.uCallbackMessage = WM_TRAYICON
        nid.hIcon = user32.LoadIconW(None, make_int_resource(IDI_APPLICATION))
        nid.szTip = self.tooltip[:127]
        return bool(shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)))

    def _remove_icon(self):
        if not self.hwnd_msg:
            return
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd_msg
        nid.uID = 1
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))

    def run(self):
        hinst = kernel32.GetModuleHandleW(None)
        global _tray_class_counter
        _tray_class_counter += 1
        class_name = f"WTR_TRAY_MSG_WINDOW_{kernel32.GetCurrentProcessId()}_{_tray_class_counter}"


        wc = WNDCLASSW()
        wc.style = 0
        wc.lpfnWndProc = self._wndproc
        wc.cbClsExtra = 0
        wc.cbWndExtra = 0
        wc.hInstance = hinst
        wc.hIcon = None
        wc.hCursor = None
        wc.hbrBackground = None
        wc.lpszMenuName = None
        wc.lpszClassName = class_name

        user32.RegisterClassW(ctypes.byref(wc))

        self.hwnd_msg = user32.CreateWindowExW(
            0,
            class_name,
            class_name,
            0,
            0, 0, 0, 0,
            None, None, hinst, None
        )

        self._add_icon()

        msg = MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        self._remove_icon()
        if self.hwnd_msg:
            user32.DestroyWindow(self.hwnd_msg)
            self.hwnd_msg = None
            user32.UnregisterClassW(class_name, hinst)

        return self._should_show, self._should_exit

def hide_to_tray():
    console_hwnd = kernel32.GetConsoleWindow()
    if console_hwnd:
        user32.ShowWindow(console_hwnd, SW_HIDE)

    tray = TrayController(tooltip="Window Title Renamer")
    should_show, should_exit = tray.run()

    if console_hwnd:
        user32.ShowWindow(console_hwnd, SW_SHOW)

    return should_show, should_exit

# ============================================================
# **I/O helpers**
# ============================================================
def read_int_allow_0(prompt: str, lo: int, hi: int) -> int:
    while True:
        s = input(prompt).strip()
        try:
            v = int(s)
        except ValueError:
            print(f"请输入 0 或 {lo} 到 {hi} 的整数。")
            continue
        if v == 0:
            return 0
        if lo <= v <= hi:
            return v
        print(f"请输入 0 或 {lo} 到 {hi} 的整数。")

def read_nonempty(prompt: str) -> str:
    while True:
        s = input(prompt)
        if s.strip():
            return s
        print("不能为空。")

def read_yesno(prompt: str) -> bool:
    while True:
        s = input(prompt).strip().lower()
        if s in ("y", "yes", "是", "true", "1"):
            return True
        if s in ("n", "no", "否", "false", "0"):
            return False
        print("请输入 y/n（也可以输入 是/否）。")

# ============================================================
# **Main loop**
# ============================================================
def main():
    # **(3) set our own window title**
    kernel32.SetConsoleTitleW("Window Title Renamer")

    keeper = PersistentRenamer()
    keeper.start()

    print("Window Title Renamer")
    print("输入编号时输入 0：隐藏到托盘后台运行（Show 返回 / Exit 退出）")
    print()

    try:
        while True:
            wins = list_open_windows()
            rules = keeper.list_rules()

            if not wins:
                print("没有找到可重命名的窗口（可见且标题非空）。1 秒后重试…")
                time.sleep(1)
                continue

            print(f"当前长久保持规则数：{len(rules)}（后台每秒重设一次）")
            print("-" * 90)
            for i, (hwnd, title) in enumerate(wins, start=1):
                mark = " *保持中*" if hwnd in rules else ""
                print(f"[{i:3d}] {title}{mark}   (HWND=0x{hwnd:016X})")
            print("-" * 90)

            choice = read_int_allow_0(
                "请输入要重命名的窗口编号（或 0 隐藏到托盘）: ",
                1, len(wins)
            )

            # **(2) 0 -> tray**
            if choice == 0:
                _, should_exit = hide_to_tray()
                if should_exit:
                    return 0
                print()
                continue

            hwnd, old_title = wins[choice - 1]
            if not user32.IsWindow(HWND(hwnd)):
                print("目标窗口已不存在（可能已关闭）。回到列表。\n")
                continue

            new_title = read_nonempty("请输入新窗口标题: ")
            persist = read_yesno("是否长久保持（每秒重复设置一次）？(y/n): ")

            ok = set_window_title(hwnd, new_title)
            if ok:
                print(f"已重命名：{old_title} -> {new_title}")
            else:
                print(f"重命名失败（GetLastError={_get_last_error()}）。可能原因：权限不足/窗口不接受 SetWindowText。")

            if persist:
                keeper.add_or_update(hwnd, new_title)
                print("已加入长久保持。\n")
            else:
                keeper.remove(hwnd)
                print()

            # **(1) loop continues: re-list and repeat**

    except KeyboardInterrupt:
        print("\n收到 Ctrl+C，退出。")
        return 0
    finally:
        keeper.stop()

if __name__ == "__main__":
    sys.exit(main())