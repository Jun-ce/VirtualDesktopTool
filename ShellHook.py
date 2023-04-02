import win32api
import ctypes
from ctypes import wintypes

# 定义常量
WM_SHELLHOOKMESSAGE = win32api.RegisterWindowMessage("SHELLHOOK")
HSHELL_VIRTUAL_DESKTOP_CHANGED = 32772

# 定义 MSG 结构
class MSG(ctypes.Structure):
    _fields_ = [("hwnd", wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam", wintypes.WPARAM),
                ("lParam", wintypes.LPARAM),
                ("time", wintypes.DWORD),
                ("pt", wintypes.POINT)]

# 注册ShellHook
def RegisterShellHook(hwnd):
    hwnd = ctypes.c_void_p(int(hwnd))
    result = ctypes.windll.user32.RegisterShellHookWindow(hwnd)
    if result == 0:
        raise ctypes.WinError()
    return result