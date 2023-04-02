# === VDA dll 接口

# 载入第三方 Windows 虚拟桌面接口
import os
import ctypes
from ctypes import wintypes

# 载入第三方 Windows 虚拟桌面接口变量的类封装
class VirtualDesktopAccessor:
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _dll_path = os.path.join(_script_dir, "dll\\VirtualDesktopAccessor.dll")
    _vda_dll = ctypes.CDLL(_dll_path)

    # 原始移动窗口到虚拟桌面的函数
    _move_window_to_desktop = _vda_dll.MoveWindowToDesktopNumber
    _move_window_to_desktop.argtypes = [wintypes.HWND, wintypes.INT]  # move_window_to_desktop(hwnd, 0)

    # 原始获取当前虚拟桌面序号的函数
    _get_current_destop_number = _vda_dll.GetCurrentDesktopNumber

    # 原始获取虚拟桌面名称的函数
    _get_desktop_name = _vda_dll.GetDesktopName
    _get_desktop_name.argtypes = [wintypes.INT, ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t]
    _get_desktop_name.restype = wintypes.INT

    # 原始获取窗口所在虚拟桌面序号的函数
    _get_window_desktop_number = _vda_dll.GetWindowDesktopNumber
    _get_window_desktop_number.argtypes = [wintypes.HWND]
    _get_window_desktop_number.restype = wintypes.INT

# 获取虚拟桌面名称的函数
def get_desktop_name(desktop_number: int) -> str:
    # 转换参数类型
    buffer_size = 256
    buffer = ctypes.create_string_buffer(buffer_size)
    buffer_pointer = ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte))

    # 调用函数，并将结果存储在 Python 变量中
    get_result = VirtualDesktopAccessor._get_desktop_name(desktop_number, buffer_pointer, buffer_size)

    # 输出结果
    result = ''
    if get_result == 1:
        result = buffer.value.decode('utf-8')
    else:
        result = "Error: " + str(desktop_number)
    return result

# 获取当前虚拟桌面序号的函数
def get_current_desktop_number() -> int:
    return VirtualDesktopAccessor._get_current_destop_number()

def get_window_desktop_number(hwnd: wintypes.HWND) -> int:
    return VirtualDesktopAccessor._get_window_desktop_number(hwnd)

# 移动窗口到虚拟桌面的函数
def move_window_to_desktop(hwnd: wintypes.HWND, desktop_number: int):
    VirtualDesktopAccessor._move_window_to_desktop(hwnd, desktop_number)