# coding: utf-8
import time

# Both of these modules are not separate libraries,
# and pywin32 library components
# Оба этих модуля не отдельные библиотеки,
# а компоненты библиотеки pywin32
import win32gui
import win32con
import win32process


class AppWindow:
    """ Application window class.
        Класс окна приложения. """
    def __init__(self, hwnd: int, title: str):
        self.hwnd = hwnd
        self.title = title

    def __str__(self):
        return f"Window (title='{self.title}', hwnd={self.hwnd})"

    def is_foreground(self) -> bool:
        """ Whether the window is active.
            Является ли окно активным. """
        return win32gui.GetForegroundWindow() == self.hwnd

    def set_foreground(self) -> None:
        """ Make the window active and maximize it on top of others.
            Сделать окно активным и развернуть его поверх других. """
        win32gui.ShowWindow(self.hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(self.hwnd)

    def minimize(self) -> None:
        """ Minimize a window.
            Свернуть окно. """
        win32gui.ShowWindow(self.hwnd, win32con.SW_MINIMIZE)

    def maximize(self) -> None:
        """ Maximize a window.
            Развернуть окно. """
        win32gui.ShowWindow(self.hwnd, win32con.SW_MAXIMIZE)

    def close(self):
        """ Закрыть окно.
            Close a window. """
        win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)


class WindowsScanner:
    @staticmethod
    def get_all_ui_windows(with_titles_only=True) -> list:
        """ Get a list of 'Window' instances for all windows in the interface.
            Получить список экземпляров 'Window' для всех окон в интерфейсе. """
        def enum_window_callback(hwnd):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if with_titles_only:
                    if not title:
                        return
                windows.append(AppWindow(hwnd, win32gui.GetWindowText(hwnd)))

        windows = []
        win32gui.EnumWindows(enum_window_callback)  # Populate list 'win_list'
        # Return a list of instances of the Window class
        return windows

    @staticmethod
    def get_process_windows(pid: int) -> list:
        """ Get a list of 'Window' instances for all windows in the process.
            Получить список экземпляров 'Window' для всех окон процесса. """
        def enum_window_callback(hwnd, pid):
            _, current_pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == current_pid and win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(pid)
                windows.append(AppWindow(hwnd, title))

        windows = []
        win32gui.EnumWindows(enum_window_callback, pid)
        return windows


if __name__ == "__main__":

    # Run this code to test the capabilities of the Window class.
    # All windows with titles will be selected.
    # Windows without titles refer to the operation of system programs and
    # Manipulation with them can lead to a failure in the graphical interface of the system.
    # Then the selected windows will be brought to the foreground one by one, on top of other windows.
    # The try-except construct will allow you to skip errors when trying to execute
    # action for those windows for which it is disabled.

    # Запустите этот код для проверки возможностей класса Window.
    # Будут выбраны все окна с заголовками.
    # Окна без заголовков относятся к работе системных программ и
    # манипуляции с ними могут привести к сбою в работе графического интерфейса системы.
    # Затем выбранные окна, по очереди будут выведены на передний план, поверх других окон.
    # Конструкция try-except позволит пропустить ошибки при попытке выполнить
    # действие для тех окон, для которых оно запрещено.

    from app_manager import AppProcessManager
    proc = AppProcessManager.get_processes_by(partial_name="chrome")

    wins = WindowHelper.get_process_windows(pid=proc[0].pid)
    for win in wins:
        print(win)
        try:
            win.set_foreground()
            time.sleep(2)
        except RuntimeError:
            print(f"Failed to install upper window {win}")




