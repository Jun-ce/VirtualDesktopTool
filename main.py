# -*- coding: utf-8 -*-
from VirtualDesktopAccessor import get_current_desktop_number, get_desktop_name, move_window_to_desktop
from AppUtility import get_is_app_running, get_app_pid, get_apps_pids, get_apps_is_running, get_window_title_from_hwnd, get_icon_from_hwnd
from ShellHook import WM_SHELLHOOKMESSAGE, HSHELL_VIRTUAL_DESKTOP_CHANGED, MSG, RegisterShellHook
from VirtualDesktopEnhancerWindow import VirtualDesktopEnhancerWindow
from VirtualDesktopEnhancerCore import VirtualDesktopEnhancerCore
# from line_profiler import LineProfiler


def main():
    core = VirtualDesktopEnhancerCore()
    core.run()  
    # get_icon_from_hwnd(0x00081AF2).save('test.png')
    # move_window_to_desktop(0x00C7293C, 3)

if __name__ == '__main__':
    main()