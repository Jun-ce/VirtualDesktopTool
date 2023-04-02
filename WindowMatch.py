from enum import Enum
from AppUtility import get_app_pid, get_app_name_from_pid
from typing import List
import pywinauto
from pywinauto import Application
import pywinauto.findwindows as findwindows
import pywinauto.win32structures as win32structures
import win32gui
import win32process
import ctypes
from VirtualDesktopAccessor import get_window_desktop_number
from UWP_Utility import get_icon_from_UWP_hwnd, package_full_name_from_handle, _get_windows_thread_process_id, _open_process, PROCESS_QUERY_LIMITED_INFORMATION

GLOBAL_MATCH_CONFIG_TOP_LEVEL_ONLY = True # 不移动子窗口，好像也没啥问题？都会跟着顶级窗口移动？
GLOBAL_MATCH_CONFIG_VISIBLE_ONLY = True # 不可见窗口没必要匹配
GLOBAL_MATCH_CONFIG_ENABLED_ONLY = False # 当打开保存对话框时，enabled_only=True 会导致无法找到窗口

# === 窗口匹配基础类
class WindowMatchMode(Enum):
    TITLE = 1
    CLASS = 2
    APP = 3
    TITLE_AND_CLASS = 4
    TITLE_AND_APP = 5
    CLASS_AND_APP = 6
    ALL = 7

# 窗口匹配配置
class WindowMatchConfig:
    def __init__(self,
                 active: bool,
                 title : str, 
                 window_class: str, 
                 app_name: str, 
                 match_mode: WindowMatchMode, 
                 ):
        self.active: bool = active
        self.title: str = title
        self.window_class: str = window_class
        self.app_name: str = app_name
        self.match_mode: WindowMatchConfig = match_mode
    
    # 从当前配置获取匹配的窗口句柄
    def get_matched_hwnds(self) -> List[int]:
        if not self.active:
            return []
        kwargs = {}
        if self.match_mode in [WindowMatchMode.TITLE, WindowMatchMode.TITLE_AND_APP, WindowMatchMode.TITLE_AND_CLASS, WindowMatchMode.ALL]:
            kwargs['title_re'] = self.title
        if self.match_mode in [WindowMatchMode.CLASS, WindowMatchMode.TITLE_AND_CLASS, WindowMatchMode.CLASS_AND_APP, WindowMatchMode.ALL]:
            kwargs['class_name'] = self.window_class
        if self.match_mode in [WindowMatchMode.APP, WindowMatchMode.TITLE_AND_APP, WindowMatchMode.CLASS_AND_APP, WindowMatchMode.ALL]:
            kwargs['process'] = get_app_pid(self.app_name)
        kwargs['enabled_only'] = GLOBAL_MATCH_CONFIG_ENABLED_ONLY
        kwargs['visible_only'] = GLOBAL_MATCH_CONFIG_VISIBLE_ONLY
        kwargs['top_level_only'] = GLOBAL_MATCH_CONFIG_TOP_LEVEL_ONLY
        return pywinauto.findwindows.find_windows(**kwargs)
    
    def get_matched_window_infos(self) -> List['WindowInfo']:
        if not self.active:
            return []
        hwnds = self.get_matched_hwnds()
        return [WindowInfo(hwnd, True) for hwnd in hwnds]
    
    def get_is_window_info_matched(self, window_info: 'WindowInfo') -> bool:
        if not self.active:
            return False
        return window_info.get_is_matched_for_config(self)
    
# 所有窗口和他们的标题、类名、进程名、应用程序名、以及是否满足当前"Match Window"中的任意匹配条件
class WindowInfo:
    def __init__(self, hwnd: int, matched: bool = False) -> None:
        self.hwnd: int = hwnd
        self.title: str = None
        self.window_class: str = None
        self.process_id: str = None
        self.app_name: str = None
        self.matched: bool = matched
        self.valid: bool = False
        self.current_desktop_idx: int = None
        self.set_window_info_from_hwnd(hwnd)

    def set_window_info_from_hwnd(self, hwnd: int) -> None:
        try:
            self.title = win32gui.GetWindowText(hwnd)
            self.window_class = win32gui.GetClassName(hwnd)
            _thread_id, self.process_id = win32process.GetWindowThreadProcessId(hwnd)
            self.app_name = get_app_name_from_pid(self.process_id)
            self.current_desktop_idx = get_window_desktop_number(hwnd)
            
            # # 测试 UWP 用
            # try:
            #     pid = ctypes.wintypes.DWORD()
            #     _get_windows_thread_process_id(
            #         hwnd,
            #         ctypes.byref(pid)
            #     )

            #     hprocess = _open_process(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            #     package_name = package_full_name_from_handle(hprocess)
            #     if (ctypes.wstring_at(package_name)):  # UWP应用
            #         print(f'UWP应用, hex_hwnd={hwnd}, {ctypes.wstring_at(package_name)}')
            # except Exception as e:
            #     print(f'获取UWP应用信息失败，hex_hwnd={hwnd}, {e}')
            # # 测试结束
            
            

            if self.current_desktop_idx is not None:
                self.valid = True
            else:
                self.valid = False
                raise Exception(f'获取窗口信息无效，hex_hwnd={hex(hwnd)}')
            if self.current_desktop_idx <= 0:
                self.valid = False
                raise Exception(f'窗口信息虚拟桌面信息获取无效，hex_hwnd={hex(hwnd)}，虚拟桌面ID {self.current_desktop_idx}')
            if self.window_class in ['Windows.UI.Core.CoreWindow', 'ApplicationFrameWindow']:
                    pass
                    # print(f'UWP应用, hex_hwnd={hwnd}, {self.title}, {self.app_name}, {self.window_class}')
        except Exception as e:
            self.valid = False
            # print(f'Error: hwnd={hwnd}，{e}')

    def get_is_matched_for_config(self, config: WindowMatchConfig) -> bool:
        if not config.active:
            return False
        if config.match_mode == WindowMatchMode.TITLE:
            return self.title == config.title
        elif config.match_mode == WindowMatchMode.CLASS:
            return self.window_class == config.window_class
        elif config.match_mode == WindowMatchMode.APP:
            return self.app_name == config.app_name
        elif config.match_mode == WindowMatchMode.TITLE_AND_CLASS:
            return self.title == config.title and self.window_class == config.window_class
        elif config.match_mode == WindowMatchMode.TITLE_AND_APP:
            return self.title == config.title and self.app_name == config.app_name
        elif config.match_mode == WindowMatchMode.CLASS_AND_APP:
            return self.window_class == config.window_class and self.app_name == config.app_name
        elif config.match_mode == WindowMatchMode.ALL:
            return self.title == config.title and self.window_class == config.window_class and self.app_name == config.app_name
        else:
            return False
        
