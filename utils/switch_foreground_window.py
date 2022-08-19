# coding=utf-8
"""
    切换置顶窗口
    参考1：https://blog.csdn.net/qq_39369520/article/details/119520185
    参考2：https://cloud.tencent.com/developer/article/1711584
"""
import win32gui
import win32com.client
import win32con


class WindowObject:
    def __init__(self):
        window_allow_list = ['Visual Studio Code']
        window_list = self.get_all_windows()
        # 从众多中选一个出来作为安全窗口
        self.window_allow = self.get_window_allow(window_allow_list, window_list)

    def get_foreground_window(self):
        """
            获取当前前置窗口的信息
            @return 当前前置窗口的信息
        """
        hwnd_active = win32gui.GetForegroundWindow()
        title, class_name = self.get_info_by_hwnd(hwnd_active)
        return {
            'hwnd': hwnd_active,
            'text': title,
            'class_name': class_name
        }

    def get_all_windows(self):
        """
            获取所有窗口的基本信息
        """
        window_list = []
        hwnd_list = []
        win32gui.EnumWindows(lambda hwnd_, param: param.append(hwnd_), hwnd_list)
        for hwnd in hwnd_list:
            title, class_name = self.get_info_by_hwnd(hwnd)
            window_list.append({
                'hwnd': hwnd,
                'text': title,
                'class_name': class_name
            })
        return window_list

    @staticmethod
    def get_info_by_hwnd(hwnd):
        """
            根据窗口句柄获取窗口标题和类名
            @param hwnd 窗口句柄
            @return title 窗口标题
            @return class_name 窗口类名
        """
        return win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd)

    @staticmethod
    def get_window_allow(window_allow_list, window_list):
        """
            筛选出需要在”危险“状况下前置的窗口的信息
        """
        window_allow = {}
        for window in window_list:
            re_flag = False
            for allow in window_allow_list:
                if allow not in window['text']:
                    continue
                re_flag = True
                break
            if re_flag:
                window_allow = window
                break
        return window_allow

    @staticmethod
    def switch_foreground_window(the_window):
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        window_found = win32gui.FindWindow(the_window['class_name'], the_window['text'])
        win32gui.SetForegroundWindow(window_found)
        win32gui.ShowWindow(window_found, win32con.SW_SHOW)
