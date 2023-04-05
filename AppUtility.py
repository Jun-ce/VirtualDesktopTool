# -*- coding: utf-8 -*-

import psutil
from pywinauto import Application
from typing import List, Dict
from PIL import Image, ImageQt
import win32gui
import win32ui
import win32con
import win32process
import win32api
from PyQt5.QtGui import QImage
from UWP_Utility import get_icon_from_UWP_hwnd
from LP_Wrapper import lp_wrapper


# 判断应用是否已经启动
def get_is_app_running(app_name: str, case_sensitive : bool = True) -> bool:
    found_processes = psutil.process_iter(['pid', 'name'])
    is_running = False

    for process in found_processes:
        process_name_to_check = process.info['name'] if case_sensitive else process.info['name'].lower()
        app_name_to_check = app_name if case_sensitive else app_name.lower()
        if process_name_to_check == app_name_to_check:
            is_running = True
            break
    return is_running

# 获取多个应用程序的进程 ID
def get_apps_pids(app_names: List[str], case_sensitive : bool = True) -> Dict[str, List[int]]:
    app_names_to_check = app_names if case_sensitive else  [name.lower() for name in app_names] 
    found_processes = psutil.process_iter(['pid', 'name'])
    results = {name: [] for name in app_names}
    for process in found_processes:
        name_to_check = process.info['name'] if case_sensitive else process.info['name'].lower()
        if name_to_check in app_names_to_check:
            idx = app_names_to_check.index(name_to_check)
            results[app_names[idx]].append(process.info['pid'])
    return results

# 获取单个应用程序的进程 ID
def get_app_pid(app_name : str, case_sensitive : bool = True) -> int:
    r = get_apps_pids([app_name], case_sensitive)
    return r[app_name][0] if len(r[app_name])>0 else None

# 判断多个应用程序是否已经启动
def get_apps_is_running(app_names : List[str], case_sensitive : bool = True) -> Dict[str, bool]:
    pids = get_apps_pids(app_names, case_sensitive)
    results = {name: len(pids[name])>0 for name in app_names}
    return results

# 获取进程 ID 对应的应用名称
def __get_app_name_from_pid(pid: int) -> str:
    name = psutil.Process(pid).name()
    if name == 'ApplicationFrameHost.exe':
        # print(f'UWP应用，pid={pid}')
        pass
    return name

def get_app_name_from_hwnd(hwnd: int) -> str:
    if hwnd is None or hwnd <= 0:
        return None
    name = None
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    name = __get_app_name_from_pid(pid)
    if name == 'ApplicationFrameHost.exe':
        try:
            core_hwnd = get_UWP_core_hwnd(hwnd)
            if not core_hwnd or core_hwnd <= 0:
                return None
            core_pid = win32process.GetWindowThreadProcessId(core_hwnd)[1]
            name = __get_app_name_from_pid(core_pid)
        except Exception as excep:
            print(f'{hwnd} - {core_hwnd} - 获取 UWP 应用名称失败，{excep}')
            pass
    return name

# 获取窗口句柄对应的窗口标题
def get_window_title_from_hwnd(hwnd: int) -> str:
    app = Application().connect(handle=hwnd)
    return app.top_window().window_text()
                
# 获取句柄的 exe 文件路径
# @lp_wrapper
def get_exe_path_from_hwnd(hwnd: int) -> str:
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    if pid is None or pid <= 0:
        return None
    try:
        process = psutil.Process(pid)
        exe_path = process.exe()
    except Exception as excep:
        print(f'{hwnd} - 获取 exe 文件路径失败，{excep}')
        return None
    return exe_path
    # for proc in psutil.process_iter(['pid', 'name', 'exe']):
    #     if proc.info['pid'] == pid:
    #         return proc.info['exe']
    # return None

# 从 exe 文件中提取最大 32x32 的图标，返回 QImage
# @lp_wrapper
def get_icon_from_exe(exe_path: str, icon_resize: int = 32) -> QImage:
    large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0, 1)
    icons = large_icons if len(large_icons) > 0 else small_icons
    img = None
    # icons = small_icons
    if icons:
        icon = icons[0]
        try:
            # 获取图标信息
            icon_info = win32gui.GetIconInfo(icon)  # （fIcon, xHotspot, yHotspot, hbmMask, hbmColor)
            # 获取图标的尺寸
            Width, Height, BitDepth = get_icon_size(icon_info)

            # 创建位图
            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, Width, Height)
            hdc = hdc.CreateCompatibleDC() # 可能失效？
            hdc.SelectObject(hbmp)
            win32gui.DrawIconEx(hdc.GetHandleOutput(), 0, 0, icon, 0, 0, 0, 0, win32con.DI_NORMAL)

            # 将位图转换为字节
            bmp_str = hbmp.GetBitmapBits(True)
            img = Image.frombuffer(
                'RGBA',
                (Width, Height),
                bmp_str,
                'raw',
                'BGRA',
                0,
                1
            )

            img = img.resize((icon_resize, icon_resize), Image.LANCZOS)
        except Exception as e:
            print(f"{exe_path} 从 exe 中提取图标失败: {e}")
        finally:
            # 销毁图标句柄
            win32gui.DestroyIcon(icon)
        if img:
            # 将 PIL.Image 转换为 QImage
            qimg = ImageQt.ImageQt(img)
            return qimg
        else:
            print(f"{exe_path} 从 exe 中提取图标失败")
    return None

# 从图标信息中获取图标尺寸和位深
def get_icon_size(icon_info) -> tuple[int, int, int]:
    if icon_info:
            if icon_info[4]: # Icon has colour plane
                bmp = win32gui.GetObject(icon_info[4])
                
                if bmp:
                    Width = bmp.bmWidth
                    Height = bmp.bmHeight
                    BitDepth = bmp.bmBitsPixel
                    return (Width, Height, BitDepth)

            else: # Icon has no colour plane, image data stored in mask
                bmp = win32gui.GetObject(icon_info[4])
                
                if bmp:
                    Width = bmp.bmWidth
                    Height = bmp.bmHeight // 2
                    BitDepth = 1
                    return (Width, Height, BitDepth)
    return (0, 0, 0)

# 从窗口句柄中提取图标，返回 QImage
# Todo 
# # if hwnd.class_name == "TaskManagerWindow": deal with Taskmgr
#   get_icon_from_exe("C:\\Windows\\System32\\Taskmgr.exe").save('Taskmgr-new.png')
# @lp_wrapper
def get_icon_from_hwnd(hwnd: int, icon_resize: int = 32) -> QImage:

    image = None

    module_name = ""
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        if handle:
            module_name = win32process.GetModuleFileNameEx(handle, 0)
            win32api.CloseHandle(handle)
    except Exception as e:
        if module_name == "": # 无法获取模块名称
            image = get_icon_special_case(hwnd, icon_resize)
        if image is None:
            # print(f"hwnd: {hwnd} 获取图标时无法打开进程: {e}")
            pass
    if image:
        return image

    # 试图从 EXE 文件中获取图标
    exe_path = get_exe_path_from_hwnd(hwnd)
    if exe_path:
        image = get_icon_from_exe(exe_path, icon_resize)

    # 试图从 UWP 包中获取图标
    if image is None: 
        core_hwnd = get_UWP_core_hwnd(hwnd)
        if core_hwnd:
            image = get_icon_from_UWP_hwnd(core_hwnd, icon_resize)

    return image

def get_icon_special_case(hwnd: int, icon_resize: int = 32) -> QImage:
    if hwnd is None or hwnd <= 0:
        return None
    # 任务管理器
    try:
        class_name = win32gui.GetClassName(hwnd)
        if class_name == "TaskManagerWindow":
            return get_icon_from_exe("C:\\Windows\\System32\\Taskmgr.exe", icon_resize)
    except Exception as e:
        pass
    return None

# 对一般exe和UWP都适用
def get_exe_path_from_pid(pid: int) -> str:
    process = psutil.Process(pid)
    return process.exe()

# 获取 UWP 的 CoreWindow 句柄
def get_UWP_core_hwnd(hwnd: int) -> int:
    if(hwnd == 0):
        return 0
    core_hwnd = 0
    try:
        if win32gui.GetClassName(hwnd) == "Windows.UI.Core.CoreWindow":  # 如果 UWP 应用被最小化或已经处于别的虚拟桌面，则顶层窗口就是 CoreWindow 的句柄
            core_hwnd = hwnd
        elif win32gui.GetClassName(hwnd) == "ApplicationFrameWindow":  # 如果 UWP 应用没有被最小化，则顶层窗口是 ApplicationFrameWindow 的句柄
            core_hwnd = win32gui.FindWindowEx(hwnd, 0, "Windows.UI.Core.CoreWindow", None)
            if not core_hwnd:
                # print(f"在沙盒中找不到 UWP 的 CoreWindow: {hwnd}，可能 UWP 应用已经被最小化或已经处于别的虚拟桌面")  # 这时候一定会存在一个直接就是 CoreWindow 的句柄，移动它就行了
                return 0
    except Exception as e:
        return 0
    return core_hwnd

# 获取 UWP 的 CoreWindow 的 PID
def get_UWP_core_pid(hwnd: int) -> int:
    core_hwnd = get_UWP_core_hwnd(hwnd)
    pid = None
    if core_hwnd:
        _, pid = win32process.GetWindowThreadProcessId(core_hwnd)
    return pid

# 获取窗口句柄的信息
def get_raw_window_info(hwnd: int) -> dict:
    info = {}

    # Get window title
    title = win32gui.GetWindowText(hwnd)
    info['title'] = title

    # Get window class name
    class_name = win32gui.GetClassName(hwnd)
    info['class_name'] = class_name

    # Get window position and size
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    info['position'] = (left, top)
    info['size'] = (right - left, bottom - top)

    # Get window style
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    info['style'] = style

    # Get window extended style
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    info['ex_style'] = ex_style

    # Get window owner
    owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
    if owner:
        owner_title = win32gui.GetWindowText(owner)
        info['owner_title'] = owner_title

    # Get window process ID
    _, process_id = win32process.GetWindowThreadProcessId(hwnd)
    info['process_id'] = process_id

    # Get window thread ID
    thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
    info['thread_id'] = thread_id

    # Get window visible state
    visible = win32gui.IsWindowVisible(hwnd)
    info['visible'] = visible

    # Get foreground window
    foreground = win32gui.GetForegroundWindow() == hwnd
    info['foreground'] = foreground

    # Get window parent
    parent = win32gui.GetParent(hwnd)
    info['parent'] = parent

    # Get window menu
    menu = win32gui.GetMenu(hwnd)
    info['menu'] = menu

    # Get window icon
    icon = win32gui.SendMessage(hwnd, win32con.WM_GETICON, win32con.ICON_BIG, 0)
    info['icon'] = icon

    # Get window placement
    window_placement = win32gui.GetWindowPlacement(hwnd)
    info['window_placement'] = window_placement

    return info
