# === VDA dll 接口

# 载入第三方 Windows 虚拟桌面接口
import os
import ctypes
from ctypes import wintypes
from LP_Wrapper import lp_wrapper

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


    # Pin 的窗口获取不到虚拟桌面序号，但是也不需要监视事件了
    
    # fn IsPinnedWindow(hwnd: HWND) -> i32
    # fn PinWindow(hwnd: HWND) -> i32
    # fn UnPinWindow(hwnd: HWND) -> i32
    # fn IsPinnedApp(hwnd: HWND) -> i32
    # fn PinApp(hwnd: HWND) -> i32
    # fn UnPinApp(hwnd: HWND) -> i32 

    _get_window_is_pinned = _vda_dll.IsPinnedWindow
    _get_window_is_pinned.argtypes = [wintypes.HWND]
    _get_window_is_pinned.restype = wintypes.INT

    _get_window_is_pinned_app = _vda_dll.IsPinnedApp
    _get_window_is_pinned_app.argtypes = [wintypes.HWND]
    _get_window_is_pinned_app.restype = wintypes.INT

    _set_window_pin = _vda_dll.PinWindow
    _set_window_pin.argtypes = [wintypes.HWND]
    _set_window_pin.restype = wintypes.INT

    _set_window_pin_app = _vda_dll.PinApp
    _set_window_pin_app.argtypes = [wintypes.HWND]
    _set_window_pin_app.restype = wintypes.INT

    _set_window_unpin = _vda_dll.UnPinWindow
    _set_window_unpin.argtypes = [wintypes.HWND]
    _set_window_unpin.restype = wintypes.INT

    _set_window_unpin_app = _vda_dll.UnPinApp
    _set_window_unpin_app.argtypes = [wintypes.HWND]
    _set_window_unpin_app.restype = wintypes.INT

    

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

# 获取窗口所在虚拟桌面序号的函数
# 很慢
# @lp_wrapper
def get_window_desktop_number(hwnd: wintypes.HWND) -> int:
    return VirtualDesktopAccessor._get_window_desktop_number(hwnd)

# # 移动窗口到虚拟桌面的函数
def move_window_to_desktop(hwnd: wintypes.HWND, desktop_number: int):
    VirtualDesktopAccessor._move_window_to_desktop(hwnd, desktop_number)

# Pin 相关函数

def get_window_is_pinned(hwnd: wintypes.HWND) -> bool:
    return VirtualDesktopAccessor._get_window_is_pinned(hwnd) == 1

def get_window_is_pinned_app(hwnd: wintypes.HWND) -> bool:
    return VirtualDesktopAccessor._get_window_is_pinned_app(hwnd) == 1

def set_window_pin(hwnd: wintypes.HWND):
    VirtualDesktopAccessor._set_window_pin(hwnd)

def set_window_pin_app(hwnd: wintypes.HWND):
    VirtualDesktopAccessor._set_window_pin_app(hwnd)

def set_window_unpin(hwnd: wintypes.HWND):
    VirtualDesktopAccessor._set_window_unpin(hwnd)

def set_window_unpin_app(hwnd: wintypes.HWND):
    VirtualDesktopAccessor._set_window_unpin_app(hwnd)




# from pyvda import AppView, VirtualDesktop

# # def get_desktop_name(desktop_number: int) -> str:
# #     # VirtualDesktop 的 number 属性从 1 开始，而不是从 0 开始
# #     return VirtualDesktop(number=desktop_number+1).name

# # def get_current_desktop_number() -> int:
# #     return VirtualDesktop.current().number - 1

# # @lp_wrapper
# # def get_window_desktop_number(hwnd: wintypes.HWND) -> int:
# #     num = -1
# #     try:
# #         if hwnd is None or hwnd <= 0:
# #             return -1
# #         app_view = AppView(hwnd=hwnd)
# #         if app_view is None:
# #             raise Exception("app_view is None")
# #         dsk = app_view.desktop
# #         if dsk is None:
# #             raise Exception("dsk is None")
# #         num = len(dsk.name)
# #     except Exception as e:
# #         # print(e)
# #         return -1
# #     return num

# # def move_window_to_desktop(hwnd: wintypes.HWND, desktop_number: int):
# #     # VirtualDesktop 的 number 属性从 1 开始，而不是从 0 开始
# #     AppView(hwnd=hwnd).move_to(VirtualDesktop(number=desktop_number+1))