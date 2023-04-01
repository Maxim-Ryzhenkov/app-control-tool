# coding: utf-8
import time

# Оба этих модуля не отдельные библиотеки,
# а компоненты библиотеки pywin32
import win32gui
import win32con


class Window:
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

    @staticmethod
    def get_all_ui_windows(with_titles_only=True) -> list:
        """ Get a list of 'Window' instances for all windows in the interface.
            Получить список экземпляров 'Window' для всех окон в интерфейсе. """
        win_list = []

        def _callback(hwnd, windows_list):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if with_titles_only:
                    if not title:
                        return True
                windows_list.append(Window(hwnd, win32gui.GetWindowText(hwnd)))
            return True

        win32gui.EnumWindows(_callback, win_list)  # populate list
        # Вернуть список экземпляров класса Window
        return win_list


if __name__ == "__main__":
    wins = Window.get_all_ui_windows()

    for win in wins:
        print(win)
        try:
            win.set_foreground()
            time.sleep(2)
        except Exception:
            print(f"Не удалось установить верхним окно {win}")




