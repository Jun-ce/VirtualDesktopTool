from VirtualDesktopAccessor import get_current_desktop_number, get_desktop_name, move_window_to_desktop
from typing import List
from AppUtility import get_window_title_from_hwnd
from WindowMatch import WindowMatchConfig, WindowInfo, WindowMatchMode, GLOBAL_MATCH_CONFIG_ENABLED_ONLY, GLOBAL_MATCH_CONFIG_VISIBLE_ONLY, GLOBAL_MATCH_CONFIG_TOP_LEVEL_ONLY
import VirtualDesktopEnhancerWindow as VDE_Window
import sys
from PyQt5.QtWidgets import QApplication
import pywinauto.findwindows as findwindows
from UWP_Utility import get_windows


class VirtualDesktopEnhancerCore:
    def __init__(self):
        self.last_desktop_idx : int = get_current_desktop_number()
        self.match_configs: List[WindowMatchConfig] = []
        self.window_infos: List[WindowInfo] = []
        self.monitoring: bool = False

        # App related
        self.last_edit_match_mode: WindowMatchMode = None

        self.load_config_file() # 加载配置文件，记得加载 GUI 语言

        self.vde_window: VDE_Window.VirtualDesktopEnhancerWindow = None
        self.qapp: QApplication = None

    def load_config_file(self):
        self.match_configs = [] # ... To do 从磁盘加载配置文件

    def save_config_file(self, path: str) -> bool:
        # ... 将配置文件保存到磁盘，返回是否成功。
        # json 文件或者 xml 文件均可
        success = True
        return success

    def add_config(self, config: WindowMatchConfig):
        self.match_configs.append(config)
        self.refresh_all_windows()

    def refresh_all_windows(self):

        # 寻找所有窗口
        # kwargs = {}
        # kwargs['enabled_only'] = GLOBAL_MATCH_CONFIG_ENABLED_ONLY
        # kwargs['visible_only'] = GLOBAL_MATCH_CONFIG_VISIBLE_ONLY
        # kwargs['top_level_only'] = GLOBAL_MATCH_CONFIG_TOP_LEVEL_ONLY
        # hwnds = findwindows.find_windows(**kwargs)

        hwnds = get_windows()

        raw_window_infos = [WindowInfo(hwnd) for hwnd in hwnds]
        self.window_infos = [info for info in raw_window_infos if info.valid]

        # 设置窗口的匹配状态
        for window in self.window_infos:
            window.matched = False
            for config in self.match_configs:
                if window.get_is_matched_for_config(config):
                    window.matched = True
                    break

    def on_desktop_changed(self):
        current_desktop_idx = get_current_desktop_number()
        if current_desktop_idx != self.last_desktop_idx:
            self.last_desktop_idx = current_desktop_idx
            self.move_matched_windows_to_desktop(current_desktop_idx)

    def on_add_window(self):
        print("on_add_window")
        # ... 将选中的窗口添加到“Match Window”列表框。

    def on_remove_window(self):
        print("on_remove_window")
        # ... 从“Match Window”列表框中删除选中的窗口。

    def on_save_config(self):
        print("on_save_config")
        # ... 保存配置文件。

    def on_load_config(self):
        print("on_load_config")
        # ... 从磁盘加载配置文件。

    def on_start_auto_move(self):
        print("on_start_auto_move")
        # ... 开始监听虚拟桌面切换事件。

    def on_stop_auto_move(self):
        print("on_stop_auto_move")
        # ... 停止监听虚拟桌面切换事件。

    def on_exit(self):
        print("on_exit")
        # ... 退出程序。

    def on_about(self):
        print("on_about")
        # ... 弹出关于窗口。

    def on_language_changed(self, index):
        print("on_language_changed")
        # ... 更改界面语言。

    def get_windows_to_move(self) -> List[int]:
        print("get_windows_to_move")
        return []
        # ... 根据配置文件中的所有条目，返回一个列表，列表中包含了所有匹配的窗口。

    def on_virtual_desktop_changed(self) -> bool:
        if not self.monitoring:
            print("不在监听虚拟桌面切换事件")
            return False
        index = get_current_desktop_number()
        if self.last_desktop_idx == index: # 事实上这个条件可能在不是切换虚拟桌面的时候也会满足，所以需要进一步判断当前的虚拟桌面序号是否真的发生变动
            print(f"同一桌面{index} {get_desktop_name(index)}的重复回调")
            pass # 不做操作
        else:
            print(f"切换到虚拟桌面 {index} {get_desktop_name(index)}")
            for hwnd in self.get_windows_to_move():
                print(f"移动句柄{hwnd} {get_window_title_from_hwnd(hwnd)} 到虚拟桌面 {index} {get_desktop_name(index)}")
                move_window_to_desktop(hwnd, index)
            self.last_desktop_idx = index
        return True


    def run(self):
        if(self.qapp is None):
            self.qapp = QApplication(sys.argv)
        self.qapp.setQuitOnLastWindowClosed(False)

        if(self.vde_window is None):
            self.vde_window = VDE_Window.VirtualDesktopEnhancerWindow(self)
        self.vde_window.show()

        sys.exit(self.qapp.exec_())
