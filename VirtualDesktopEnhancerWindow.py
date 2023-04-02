import ctypes
from PyQt5.QtCore import Qt, QTimer, QSize, QLocale, QRect
from PyQt5.QtGui import QImage, QFont, QPainter, QPen, QPixmap, QIcon
from PyQt5.QtWidgets import *
from VirtualDesktopAccessor import get_current_desktop_number, get_desktop_name, move_window_to_desktop
import sys
from ShellHook import WM_SHELLHOOKMESSAGE, HSHELL_VIRTUAL_DESKTOP_CHANGED, MSG, RegisterShellHook
import VirtualDesktopEnhancerCore as VDE_Core
from AppUtility import get_icon_from_hwnd
from UWP_Utility import package_full_name_from_handle
from WindowMatch import WindowInfo, WindowMatchConfig, WindowMatchMode

class CurrentWindowItem(QListWidgetItem):
    @classmethod
    def from_WindowInfo(self, window_info: WindowInfo):
        # self.icon_img = get_icon_from_hwnd(window_info.hwnd)
        item = CurrentWindowItem(window_info.matched, window_info.hwnd, window_info.title, window_info)
        return item

    def __init__(self, matched: bool, hwnd: int, title: str, window_info: WindowInfo = None):
        super().__init__()
        self.window_info = window_info
        self.icon_img = None
        self.refresh()
    
    def refresh(self):
        self.matched = self.window_info.matched
        self.hwnd = self.window_info.hwnd
        self.title = self.window_info.title
        self.current_desktop_idx = self.window_info.current_desktop_idx
        self.pid = self.window_info.process_id
        # if self.title in ["Calculator", "Microsoft Store", "任务管理器"]:
        #     print(f"窗口： {self.title}")

        self.icon_img = get_icon_from_hwnd(self.hwnd)
        
        icon = QIcon(QPixmap.fromImage(self.icon_img)) if self.icon_img is not None else None
        if icon is not None:
            self.setIcon(icon)
        else:
            print(f"未找到图标： {self.title}")
            # app = QApplication.instance()
            # app_icon = app.style().standardPixmap(app.style().SP_DesktopIcon)
            # fallback_icon = QIcon()
            # fallback_icon.addPixmap(app_icon, QIcon.Normal, QIcon.Off)
            # self.setIcon(fallback_icon)

        self.set_current_window_text()


    def set_current_window_text(self):
        # TODO: get from config
        show_desktop_name = True
        show_matched_state = True
        show_hwnd = True
        show_pid = True
        show_hex = True

        check_mark = "[✓]" if self.matched and show_matched_state else ""
        desktop_name = f"{self.current_desktop_idx} {get_desktop_name(self.current_desktop_idx)} - " if show_desktop_name else ""
        hwnd_text = f" - hwnd: {self.hwnd if not show_hex else hex(self.hwnd)}" if show_hwnd else ""
        pid_text = f" - pid: {self.pid if not show_hex else hex(self.pid)}" if show_pid else ""
        self.setText(f"{check_mark}{desktop_name}{self.title}{pid_text}{hwnd_text}")
        self.setForeground(Qt.blue if self.matched else Qt.black)

class MatchedWindowItem(QListWidgetItem):
    def __init__(self, hwnd: int, title: str, icon: QImage):
        super().__init__()

        self.hwnd = hwnd
        self.title = title
        self.icon = icon

        self.set_matched_window_text()
        self.setIcon(QIcon(QPixmap.fromImage(self.icon)))

    def set_matched_window_text(self):
        self.setText(f"{self.title} - {self.hwnd}")
        self.setForeground(Qt.black)

class VirtualDesktopEnhancerWindow(QMainWindow):
    def __init__(self, core: VDE_Core.VirtualDesktopEnhancerCore):
        super().__init__()

        self.core: VDE_Core.VirtualDesktopEnhancerCore = core
        self.locale = QLocale()

        # 初始化UI
        self.init_ui()

        # 加载配置文件
        self.load_config()

        # 设置定时器
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_all_windows)
        self.refresh_timer.start(1000)

        # 注册ShellHook
        RegisterShellHook(self.winId())


    def init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # 主布局
        vbox_main = QVBoxLayout()
        central_widget.setLayout(vbox_main)

        # 语言选择
        lang_hbox = QHBoxLayout()
        vbox_main.addLayout(lang_hbox)

        lang_label = QLabel(self.tr("Language:"), self)
        lang_hbox.addWidget(lang_label)

        lang_combo = QComboBox(self)
        lang_combo.addItems(["English", "中文 - Chinese"])  # Todo: 从 config 里读取
        lang_hbox.addWidget(lang_combo)

        # 显示当前虚拟桌面
        vbox_current_desktop_label = QVBoxLayout()
        vbox_main.addLayout(vbox_current_desktop_label)
        
        self.current_vd_label = QLabel(self)
        self.current_vd_label.setText(f"Current Virtual Desktop: {get_current_desktop_number()} {get_desktop_name(get_current_desktop_number())}")
        vbox_current_desktop_label.addWidget(self.current_vd_label)

        # 窗口列表和分割器
        hbox_window_lists = QHBoxLayout()
        vbox_main.addLayout(hbox_window_lists)
        vbox_main.setStretchFactor(hbox_window_lists, 1)

        window_list_splitter = QSplitter()
        hbox_window_lists.addWidget(window_list_splitter)
        window_list_splitter.setStretchFactor(0, 1)
        window_list_splitter.setStretchFactor(1, 1)
        window_list_splitter.setChildrenCollapsible(False)

        # 所有窗口列表        
        all_window_frame = QFrame()
        all_window_frame.setMinimumWidth(200)
        window_list_splitter.addWidget(all_window_frame)
        
        all_window_vbox = QVBoxLayout()
        all_window_frame.setLayout(all_window_vbox)

        all_windows_list_vbox = QVBoxLayout()
        all_window_vbox.addLayout(all_windows_list_vbox)

        all_windows_label = QLabel("All Windows:", self)
        all_windows_list_vbox.addWidget(all_windows_label)

        self.all_windows_list = QListWidget(self)
        all_windows_list_vbox.addWidget(self.all_windows_list)

        add_window_btn_vbox = QVBoxLayout()
        all_window_vbox.addLayout(add_window_btn_vbox)

        self.add_window_btn = QPushButton("Add Window", self)
        self.add_window_btn.clicked.connect(self.on_add_window)
        add_window_btn_vbox.addWidget(self.add_window_btn)

        # 匹配窗口列表
        match_window_frame = QFrame()
        match_window_frame.setMinimumWidth(200)
        window_list_splitter.addWidget(match_window_frame)

        match_window_vbox = QVBoxLayout()
        match_window_frame.setLayout(match_window_vbox)

        match_windows_vbox = QVBoxLayout()
        match_window_vbox.addLayout(match_windows_vbox)

        match_windows_label = QLabel("Match Windows:", self)
        match_windows_vbox.addWidget(match_windows_label)

        self.match_window_list = QListWidget(self)
        match_windows_vbox.addWidget(self.match_window_list)

        remove_window_vbox = QVBoxLayout()
        match_window_vbox.addLayout(remove_window_vbox)

        self.remove_window_btn = QPushButton("Remove Window", self)
        self.remove_window_btn.clicked.connect(self.on_remove_window)
        remove_window_vbox.addWidget(self.remove_window_btn)

        # 控制面板
        hbox_control_panel = QHBoxLayout()
        vbox_main.addLayout(hbox_control_panel)

        save_load_vbox = QVBoxLayout()
        hbox_control_panel.addLayout(save_load_vbox)

        self.save_config_btn = QPushButton("Save Config", self)
        self.save_config_btn.clicked.connect(self.on_save_config)
        save_load_vbox.addWidget(self.save_config_btn)

        self.load_config_btn = QPushButton("Load Config", self)
        self.load_config_btn.clicked.connect(self.on_load_config)
        save_load_vbox.addWidget(self.load_config_btn)

        spacer1 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        hbox_control_panel.addItem(spacer1)

        self.move_window_to_me_btn = QPushButton("Move windows to me", self)
        self.move_window_to_me_btn.clicked.connect(self.on_move_window_to_me)
        hbox_control_panel.addWidget(self.move_window_to_me_btn)

        spacer2 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        hbox_control_panel.addItem(spacer2)

        running_control_vbox = QVBoxLayout()
        hbox_control_panel.addLayout(running_control_vbox)

        running_state_vbox = QVBoxLayout()
        running_control_vbox.addLayout(running_state_vbox)

        self.running_state_label = QLabel(self)
        self.running_state_label.setText("Auto Move: Stopped")
        running_state_vbox.addWidget(self.running_state_label)

        self.auto_move_btn = QPushButton("Start Auto Move", self)
        self.auto_move_btn.clicked.connect(self.on_start_auto_move)
        running_control_vbox.addWidget(self.auto_move_btn)

        self.test_btn = QPushButton("Test", self)
        self.test_btn.clicked.connect(self.on_test)
        running_control_vbox.addWidget(self.test_btn)

        self.test_text_box = QLineEdit(self)
        running_control_vbox.addWidget(self.test_text_box)

        # # 测试
        # qItem1 = QListWidgetItem("测试标题1")
        # # qItem2 = QListWidgetItem("测试标题2")
        # item1 = CurrentWindowItem(False, 0x0F13069A, "测试标题1", get_icon_from_hwnd(0x0F13069A))
        # item2 = CurrentWindowItem(True, 0x0F13069A, "测试标题2", get_icon_from_hwnd(0x0F13069A))

        self.init_window_list_content()
        self.setWindowTitle("Virtual Desktop Enhancer")

        self.size = QSize(1080, 960)

    def init_window_list_content(self):
        self.core.refresh_all_windows() 
        for window_info in self.core.window_infos:
            item = CurrentWindowItem.from_WindowInfo(window_info)
            self.all_windows_list.addItem(item)

        # 测试用：把所有无图标的移动到最前面
        for i in range(self.all_windows_list.count()):
            item = self.all_windows_list.item(i)
            if item.icon_img is None:
                # print(f"item {item.title} has no icon")
                self.all_windows_list.takeItem(i)
                self.all_windows_list.insertItem(0, item)
                print(f"item {item.title} with hwnd {item.hwnd} package name = {package_full_name_from_handle(item.hwnd)}")
        
    def on_test(self):
        move_window_to_desktop(int(self.test_text_box.text(), 16), 1)

    def load_config(self):
        print("load_config")
        # ... 从磁盘加载配置文件，并更新界面元素。

    def save_config(self):
        print("save_config")
        # ... 将配置文件保存到磁盘。

    def refresh_all_windows(self):
        pass
        # print("refresh_all_windows")
        # ... 每隔1秒刷新“All Window”列表框。

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

    def get_windows_to_move(self) -> list[int]:
        print("get_windows_to_move")
        return []
        # ... 根据配置文件中的所有条目，返回一个列表，列表中包含了所有匹配的窗口。

    def move_windows(self, hwnds: list[int]):
        print("move_windows")
        # ... 将匹配的窗口移动到当前虚拟桌面。

    def on_move_window_to_me(self):
        print("on_move_window_to_me")
        # ... 将所有匹配的窗口移动到当前虚拟桌面。

    def tr(self, text: str) -> str:
        return text

    def nativeEvent(self, eventType, message):
        msg = ctypes.wintypes.MSG.from_address(message.__int__())
        if msg.message == WM_SHELLHOOKMESSAGE and msg.wParam == HSHELL_VIRTUAL_DESKTOP_CHANGED:
            success = self.core.on_virtual_desktop_changed()
            if success:
                # self.label.setText(f"当前虚拟桌面: {get_current_destop_number()} {get_desktop_name(get_current_destop_number())}")
                self.update()
        return super(VirtualDesktopEnhancerWindow, self).nativeEvent(eventType, message)