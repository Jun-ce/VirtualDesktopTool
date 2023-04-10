# -*- coding: utf-8 -*-

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
from LP_Wrapper import lp_wrapper

SHOW_MATCHED_STATE = True
SHOW_DESKTOP_NAME = True
SHOW_APP_NAME = True
SHOW_PID = True
SHOW_HWND = True
SHOW_HEX = False

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
        self.pinned = self.window_info.pinned
        self.hwnd = self.window_info.hwnd
        self.title = self.window_info.title
        self.current_desktop_idx = self.window_info.current_desktop_idx
        self.app_name = self.window_info.app_name
        self.pid = self.window_info.process_id
        self.is_UWP = self.window_info.is_UWP

        self.icon_img = get_icon_from_hwnd(self.hwnd)
        
        icon = QIcon(QPixmap.fromImage(self.icon_img)) if self.icon_img is not None else None
        if icon is not None:
            self.setIcon(icon)
        else:
            print(f"hwnd: {self.hwnd} 未找到图标： {self.title} - {self.app_name}")
            app = QApplication.instance()
            app_icon = app.style().standardPixmap(app.style().SP_DesktopIcon)
            fallback_icon = QIcon()
            fallback_icon.addPixmap(app_icon, QIcon.Normal, QIcon.Off)
            self.setIcon(fallback_icon)
        self.set_current_window_text()


    def set_current_window_text(self):
        # TODO: get from config
        show_desktop_name = SHOW_DESKTOP_NAME
        show_matched_state = SHOW_MATCHED_STATE
        show_hwnd = SHOW_HWND
        show_pid = SHOW_PID
        show_hex = SHOW_HEX
        show_app_name = SHOW_APP_NAME

        if show_matched_state:
            if self.matched and self.pinned:
                check_mark = "[✓] "
            elif self.matched and not self.pinned:
                check_mark = "[▲] "
            elif not self.matched and self.pinned:
                check_mark = "[▼] "
            else:
                check_mark = "     "
        else:
            check_mark = "     "
        desktop_name = (f"{self.current_desktop_idx}<{get_desktop_name(self.current_desktop_idx)}> - " if not self.pinned else "<Pinned> - ") if show_desktop_name else ""
        app_name_text = f"[{self.app_name}] - " if show_app_name else ""
        hwnd_text = f" - HWND: {self.hwnd if not show_hex else hex(self.hwnd)}" if show_hwnd else ""
        pid_text = f" - PID: {self.pid if not show_hex else hex(self.pid)}" if show_pid else ""
        is_UWP_text = " - (UWP)" if self.is_UWP else ""
        self.setText(f"{check_mark}{desktop_name}{app_name_text}{self.title}{pid_text}{hwnd_text}{is_UWP_text}")

        color = Qt.black
        if self.matched:
            if self.pinned:
                color = Qt.darkGreen
            else:
                color = Qt.yellow
        else:
            if self.pinned:
                color = Qt.darkRed

        self.setForeground(color)

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
        # self.refresh_timer = QTimer(self)
        # self.refresh_timer.timeout.connect(self.refresh_all_windows)
        # self.refresh_timer.start(5000)

        # 注册ShellHook
        # RegisterShellHook(self.winId())

        # 初始化系统托盘
        self.init_system_tray()


    def init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # 主布局
        vbox_main = QVBoxLayout()
        central_widget.setLayout(vbox_main)

        # 语言选择
        # lang_hbox = QHBoxLayout()
        # vbox_main.addLayout(lang_hbox)

        # lang_label = QLabel(self.tr("Language:"), self)
        # lang_hbox.addWidget(lang_label)

        # lang_combo = QComboBox(self)
        # lang_combo.addItems(["English", "中文 - Chinese"])  # Todo: 从 config 里读取
        # lang_hbox.addWidget(lang_combo)

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
        all_window_frame.setMinimumWidth(720)
        all_window_frame.setMinimumHeight(960)
        window_list_splitter.addWidget(all_window_frame)
        
        all_window_vbox = QVBoxLayout()
        all_window_frame.setLayout(all_window_vbox)

        all_windows_list_vbox = QVBoxLayout()
        all_window_vbox.addLayout(all_windows_list_vbox)

        self.all_windows_label = QLabel("All Windows:", self)
        all_windows_list_vbox.addWidget(self.all_windows_label)

        self.all_windows_list = QListWidget(self)
        self.all_windows_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        all_windows_list_vbox.addWidget(self.all_windows_list)

        refresh_all_windows_btn_Hbox = QHBoxLayout()
        all_window_vbox.addLayout(refresh_all_windows_btn_Hbox)

        self.refresh_all_windows_btn = QPushButton("Refresh", self)
        self.refresh_all_windows_btn.clicked.connect(self.on_refresh_button_clicked)
        refresh_all_windows_btn_Hbox.addWidget(self.refresh_all_windows_btn)
        
        add_window_btn_Hbox = QHBoxLayout()
        all_window_vbox.addLayout(add_window_btn_Hbox)        

        # self.add_window_btn = QPushButton("Add Window", self)
        # self.add_window_btn.clicked.connect(self.on_add_window)
        # add_window_btn_Hbox.addWidget(self.add_window_btn)

        self.unpin_all_window_btn = QPushButton("Unpin All Windows", self)
        self.unpin_all_window_btn.clicked.connect(self.on_unpin_all_windows)
        add_window_btn_Hbox.addWidget(self.unpin_all_window_btn)

        self.toggle_pin_window_btn = QPushButton("Toggle Pin", self)
        self.toggle_pin_window_btn.clicked.connect(self.on_toggle_pin_window)
        add_window_btn_Hbox.addWidget(self.toggle_pin_window_btn)

        # 显示配置面板
        self.display_config_checkboxes_vbox = QVBoxLayout()
        all_window_vbox.addLayout(self.display_config_checkboxes_vbox)

        self.show_matched_state_checkbox = QCheckBox("Show Pinned State", self)
        self.show_matched_state_checkbox.setChecked(SHOW_MATCHED_STATE)
        self.show_matched_state_checkbox.stateChanged.connect(self.on_show_matched_state_checkbox_state_changed)
        self.display_config_checkboxes_vbox.addWidget(self.show_matched_state_checkbox)
        
        self.show_desktop_name_checkbox = QCheckBox("Show Desktop Name", self)
        self.show_desktop_name_checkbox.setChecked(SHOW_DESKTOP_NAME)
        self.show_desktop_name_checkbox.stateChanged.connect(self.on_show_desktop_name_checkbox_state_changed)
        self.display_config_checkboxes_vbox.addWidget(self.show_desktop_name_checkbox)

        self.show_app_name_checkbox = QCheckBox("Show App Name", self)
        self.show_app_name_checkbox.setChecked(SHOW_APP_NAME)
        self.show_app_name_checkbox.stateChanged.connect(self.on_show_app_name_checkbox_state_changed)
        self.display_config_checkboxes_vbox.addWidget(self.show_app_name_checkbox)
        
        self.show_pid_checkbox = QCheckBox("Show PID", self)
        self.show_pid_checkbox.setChecked(SHOW_PID)
        self.show_pid_checkbox.stateChanged.connect(self.on_show_pid_checkbox_state_changed)
        self.display_config_checkboxes_vbox.addWidget(self.show_pid_checkbox)

        self.show_hwnd_checkbox = QCheckBox("Show HWND", self)
        self.show_hwnd_checkbox.setChecked(SHOW_HWND)
        self.show_hwnd_checkbox.stateChanged.connect(self.on_show_hwnd_checkbox_state_changed)
        self.display_config_checkboxes_vbox.addWidget(self.show_hwnd_checkbox)

        self.show_hex_checkbox = QCheckBox("Show PID and HWND in hexadecimal", self)
        self.show_hex_checkbox.setChecked(SHOW_HEX)
        self.show_hex_checkbox.stateChanged.connect(self.on_show_hex_checkbox_state_changed)
        self.display_config_checkboxes_vbox.addWidget(self.show_hex_checkbox)


 
        # 匹配窗口列表
        # match_window_frame = QFrame()
        # match_window_frame.setMinimumWidth(200)
        # window_list_splitter.addWidget(match_window_frame)

        # match_window_vbox = QVBoxLayout()
        # match_window_frame.setLayout(match_window_vbox)

        # match_windows_vbox = QVBoxLayout()
        # match_window_vbox.addLayout(match_windows_vbox)

        # match_windows_label = QLabel("Match Windows:", self)
        # match_windows_vbox.addWidget(match_windows_label)

        # self.match_window_list = QListWidget(self)
        # match_windows_vbox.addWidget(self.match_window_list)

        # remove_window_vbox = QVBoxLayout()
        # match_window_vbox.addLayout(remove_window_vbox)

        # self.remove_window_btn = QPushButton("Remove Window", self)
        # self.remove_window_btn.clicked.connect(self.on_remove_window)
        # remove_window_vbox.addWidget(self.remove_window_btn)

        # 控制面板
        # hbox_control_panel = QHBoxLayout()
        # vbox_main.addLayout(hbox_control_panel)

        # save_load_vbox = QVBoxLayout()
        # hbox_control_panel.addLayout(save_load_vbox)

        # self.save_config_btn = QPushButton("Save Config", self)
        # self.save_config_btn.clicked.connect(self.on_save_config)
        # save_load_vbox.addWidget(self.save_config_btn)

        # self.load_config_btn = QPushButton("Load Config", self)
        # self.load_config_btn.clicked.connect(self.on_load_config)
        # save_load_vbox.addWidget(self.load_config_btn)

        # spacer1 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        # hbox_control_panel.addItem(spacer1)

        # self.move_window_to_me_btn = QPushButton("Move windows to me", self)
        # self.move_window_to_me_btn.clicked.connect(self.on_move_window_to_me)
        # hbox_control_panel.addWidget(self.move_window_to_me_btn)

        # spacer2 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        # hbox_control_panel.addItem(spacer2)

        # running_control_vbox = QVBoxLayout()
        # hbox_control_panel.addLayout(running_control_vbox)

        # running_state_vbox = QVBoxLayout()
        # running_control_vbox.addLayout(running_state_vbox)

        # self.running_state_label = QLabel(self)
        # self.running_state_label.setText("Auto Move: Stopped")
        # running_state_vbox.addWidget(self.running_state_label)

        # self.auto_move_btn = QPushButton("Start Auto Move", self)
        # self.auto_move_btn.clicked.connect(self.on_start_auto_move)
        # running_control_vbox.addWidget(self.auto_move_btn)

        # self.test_btn = QPushButton("Test", self)
        # self.test_btn.clicked.connect(self.on_test)
        # running_control_vbox.addWidget(self.test_btn)

        # self.test_text_box = QLineEdit(self)
        # running_control_vbox.addWidget(self.test_text_box)
        
        # 其他初始化
        self.core.refresh_all_windows() 
        self.refresh_window_list_content()
        self.setWindowTitle("Virtual Desktop Enhancer")

        self.setGeometry(0, 0, 800, 960)

        screen = QDesktopWidget().screenGeometry()  # 获取屏幕尺寸
        window = self.geometry()
        self.move((screen.width() - window.width()) // 2, (screen.height() - window.height()) // 4)


    def refresh_window_list_content(self):
        self.all_windows_list.clear()
        for window_info in self.core.window_infos:
            item = CurrentWindowItem.from_WindowInfo(window_info)
            self.all_windows_list.addItem(item)
        self.all_windows_label.setText(f"All Windows ({self.all_windows_list.count()}):")
            
        # 把匹配和 pinned 的窗口放到最前面
        for i in range(self.all_windows_list.count()):
            item = self.all_windows_list.item(i)
            if item.matched is True:
                self.all_windows_list.takeItem(i)
                self.all_windows_list.insertItem(0, item)
        
        for i in range(self.all_windows_list.count()):
            item = self.all_windows_list.item(i)
            if item.pinned is True:
                self.all_windows_list.takeItem(i)
                self.all_windows_list.insertItem(0, item)
        
    def on_test(self):
        move_window_to_desktop(int(self.test_text_box.text(), 16), 1)

    def load_config(self):
        print("load_config")
        # ... 从磁盘加载配置文件，并更新界面元素。

    def save_config(self):
        print("save_config")
        # ... 将配置文件保存到磁盘。

    # @lp_wrapper
    def refresh_all_windows(self):
        self.core.refresh_all_windows()
        self.refresh_window_list_content()

    def on_add_window(self):
        print("on_add_window")
        # ... 将选中的窗口添加到“Match Window”列表框。

    def on_toggle_pin_window(self):
        for item in self.all_windows_list.selectedItems():
            self.core.toggle_pin_window(item.window_info)
            item.refresh()
        self.refresh_window_list_content()
        print("on_toggle_pin_window")
        # ... 切换选中的窗口的“Pin”状态。

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

    # def nativeEvent(self, eventType, message):
    #     msg = ctypes.wintypes.MSG.from_address(message.__int__())
    #     if msg.message == WM_SHELLHOOKMESSAGE and msg.wParam == HSHELL_VIRTUAL_DESKTOP_CHANGED:
    #         success = self.core.on_virtual_desktop_changed()
    #         if success:
    #             # self.label.setText(f"当前虚拟桌面: {get_current_destop_number()} {get_desktop_name(get_current_destop_number())}")
    #             self.update()
    #     return super(VirtualDesktopEnhancerWindow, self).nativeEvent(eventType, message)

    def init_system_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        app = QApplication.instance()
        app_icon = app.style().standardPixmap(app.style().SP_DesktopIcon)
        fallback_icon = QIcon()
        fallback_icon.addPixmap(app_icon, QIcon.Normal, QIcon.Off)
        self.tray_icon.setIcon(fallback_icon)
        self.tray_icon.setToolTip("Virtual Desktop Enhancer")
        self.tray_icon.show()
        

        # 创建右键菜单
        tray_menu = QMenu()

        unpin_all_action = QAction("Unpin All Windows", self)
        unpin_all_action.triggered.connect(self.on_unpin_all_windows)
        tray_menu.addAction(unpin_all_action)

        tray_menu.addSeparator()
        
        # 创建恢复窗口的操作
        restore_action = QAction("Restore", self)
        restore_action.triggered.connect(self.restore_window)
        tray_menu.addAction(restore_action)

        # 创建退出程序的操作
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.exit_app)
        tray_menu.addAction(exit_action)

        # 设置系统托盘图标的菜单
        self.tray_icon.setContextMenu(tray_menu)

        # 当系统托盘图标被激活时（例如，双击它），显示窗口
        self.tray_icon.activated.connect(self.on_tray_icon_activated)

    def closeEvent(self, event):
        # 重写关闭事件，最小化到系统托盘而不关闭应用程序
        event.ignore()
        self.hide()
        # self.tray_icon.show()
        # print("closeEvent")

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.restore_window()

    def restore_window(self):
        self.show()
        move_window_to_desktop(int(self.winId()), get_current_desktop_number())
        self.showNormal()
        self.on_refresh_button_clicked()

    def exit_app(self):
        print("exit_app")
        self.tray_icon.hide()
        QApplication.quit()

    def on_unpin_all_windows(self):
        self.core.unpin_all_windows()
        self.refresh_window_list_content()

    def on_show_matched_state_checkbox_state_changed(self, state):
        global SHOW_MATCHED_STATE
        if state == Qt.Checked:
            SHOW_MATCHED_STATE = True
        else:
            SHOW_MATCHED_STATE = False
        self.refresh_window_list_content()

    def on_show_desktop_name_checkbox_state_changed(self, state):
        global SHOW_DESKTOP_NAME
        if state == Qt.Checked:
            SHOW_DESKTOP_NAME = True
        else:
            SHOW_DESKTOP_NAME = False
        self.refresh_window_list_content()

    def on_show_app_name_checkbox_state_changed(self, state):
        global SHOW_APP_NAME
        if state == Qt.Checked:
            SHOW_APP_NAME = True
        else:
            SHOW_APP_NAME = False
        self.refresh_window_list_content()

    def on_show_pid_checkbox_state_changed(self, state):
        global SHOW_PID
        if state == Qt.Checked:
            SHOW_PID = True
        else:
            SHOW_PID = False
        self.refresh_window_list_content()

    def on_show_hwnd_checkbox_state_changed(self, state):
        global SHOW_HWND
        if state == Qt.Checked:
            SHOW_HWND = True
        else:
            SHOW_HWND = False
        self.refresh_window_list_content()

    def on_show_hex_checkbox_state_changed(self, state):
        global SHOW_HEX
        if state == Qt.Checked:
            SHOW_HEX = True
        else:
            SHOW_HEX = False
        self.refresh_window_list_content()

    def on_refresh_button_clicked(self):
        self.core.refresh_all_windows() 
        self.refresh_window_list_content()
        self.current_vd_label.setText(f"Current Virtual Desktop: {get_current_desktop_number()} {get_desktop_name(get_current_desktop_number())}")