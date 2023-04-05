# -*- coding: utf-8 -*-

from enum import Enum
from AppUtility import *
from UWP_Utility import *
from typing import List
import pywinauto
from pywinauto import Application
import pywinauto.findwindows as findwindows
import pywinauto.win32structures as win32structures
import win32gui
import win32process
import ctypes
from VirtualDesktopAccessor import *
from LP_Wrapper import lp_wrapper
import line_profiler

GLOBAL_MATCH_CONFIG_TOP_LEVEL_ONLY = True # 不移动子窗口，好像也没啥问题？都会跟着顶级窗口移动？
GLOBAL_MATCH_CONFIG_VISIBLE_ONLY = True # 不可见窗口没必要匹配
GLOBAL_MATCH_CONFIG_ENABLED_ONLY = False # 当打开保存对话框时，enabled_only=True 会导致无法找到窗口

# === 窗口匹配基础类
class WindowMatchMode(Enum):
    TITLE = 1
    CLASS = 2 # 如果是 UWP 应用，那么匹配时窗口类替换为 UWP 应用的包名
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
                 is_UWP: bool,
                 package_name: str,
                 match_mode: WindowMatchMode, 
                 ):
        self.active: bool = active
        self.title: str = title
        self.window_class: str = window_class
        self.app_name: str = app_name
        self.is_UWP: bool = is_UWP  # 是否是 UWP 应用，如果是 UWP 应用，那么匹配时窗口类视为 UWP 应用的包名
        self. package_name: str = package_name
        self.match_mode: WindowMatchConfig = match_mode
    
    # 从当前配置获取匹配的窗口句柄
    def get_matched_hwnds(self) -> List[int]:
        if not self.active:
            return []
        kwargs = {}
        if not self.is_UWP: #匹配一般窗口
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
        else: #匹配 UWP 窗口
            pass # TODO: 匹配 UWP 窗口
    
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
        self.hwnd: int = hwnd  # Top level window handle, even if the window is a UWP app.  It's not the UWP core window handle.
        self.title: str = None
        self.window_class: str = None
        self.is_UWP: bool = False  # 如果是 UWP 应用，那么匹配时窗口类替换为 UWP 应用的包名
        self.package_name: str = None
        self.process_id: str = None
        self.app_name: str = None
        self.valid: bool = False
        self.current_desktop_idx: int = None
        self.matched: bool = matched
        self.pinned: bool = False
        self.refresh_window_info_from_hwnd(hwnd)

    # @lp_wrapper
    def refresh_window_info_from_hwnd(self, hwnd: int) -> None:
        # try:
        self.window_class = win32gui.GetClassName(hwnd)
        if self.window_class is None or self.window_class == '':
            self.valid = False
            return
        self.is_UWP = self.window_class in ['Windows.UI.Core.CoreWindow', 'ApplicationFrameWindow']

        if self.window_class == 'Windows.UI.Core.CoreWindow':
            self.is_UWP = True
        elif self.window_class == 'ApplicationFrameWindow':
            self.hwnd = get_UWP_core_hwnd(hwnd)
            if self.hwnd is None or self.hwnd <= 0:  # 这种情况是空的 UWP 沙盒，UWP 的 Core Window 最小化或在其他虚拟桌面的情况，Core Window 是额外的顶层窗口
                self.valid = False  
                return
            self.is_UWP = True

        # 为了能匹配到最小化的 UWP 窗口，必须采用 Core Window 的标题，这可能和用户看到的标题不一致，例如 Core Window 的标题为 "Calander" 的应用，显示的标题是 "Month View - Calender"，这个标题只有沙盒窗口才有        
        self.title = win32gui.GetWindowText(hwnd)
        
        if self.title is None or self.title == '': # 隐藏窗口的情况
            self.valid = False
            return
            # raise Exception(f'获取窗口信息无效，标题为空，hex_hwnd={hex(hwnd)}')

        if self.is_UWP:
            self.process_id = get_UWP_core_pid(hwnd)
            self.package_name = package_full_name_from_handle(get_UWP_core_hwnd(hwnd))
        else:
            _thread_id, self.process_id = win32process.GetWindowThreadProcessId(hwnd)
        
        if self.process_id is None or self.process_id <= 0:
            self.valid = False
            return
            # raise Exception(f'获取窗口信息无效，进程信息失效，hex_hwnd={hex(hwnd)}')
        
        has_pin_info = False

        try:
            self.pinned = get_window_is_pinned(hwnd)
            has_pin_info = True
        except Exception as e:
            pass

        if not has_pin_info:
            self.valid = False
            return

        try:
            self.current_desktop_idx = get_window_desktop_number(hwnd)
        except Exception as e:
            pass
        self.valid = True
        if self.current_desktop_idx is not None:
            self.valid = True
            if self.current_desktop_idx < 0 and not self.pinned:
                self.valid = False
                return
                # raise Exception(f'窗口信息虚拟桌面信息获取无效，hex_hwnd={hex(hwnd)}，虚拟桌面ID {self.current_desktop_idx}')
        else:
            self.valid = False
            return
            raise Exception(f'获取窗口信息无效，hex_hwnd={hex(hwnd)}')


        self.app_name = get_app_name_from_hwnd(hwnd)

    def get_is_matched_for_config(self, config: WindowMatchConfig) -> bool:
        if not config.active:
            return False
        if self.is_UWP != config.is_UWP:
            return False
        
        title_matched = self.title == config.title
        class_matched = self.window_class == config.window_class if not self.is_UWP else self.package_name == config.package_name
        app_matched = self.app_name == config.app_name

        if config.match_mode == WindowMatchMode.TITLE:
            return title_matched
        elif config.match_mode == WindowMatchMode.CLASS:
            return class_matched
        elif config.match_mode == WindowMatchMode.APP:
            return app_matched
        elif config.match_mode == WindowMatchMode.TITLE_AND_CLASS:
            return title_matched and class_matched
        elif config.match_mode == WindowMatchMode.TITLE_AND_APP:
            return title_matched and app_matched
        elif config.match_mode == WindowMatchMode.CLASS_AND_APP:
            return class_matched and app_matched
        elif config.match_mode == WindowMatchMode.ALL:
            return  title_matched and class_matched and app_matched
        else:
            return False
        
