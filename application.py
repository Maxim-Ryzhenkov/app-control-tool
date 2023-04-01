# coding: utf-8

import time
import win32gui
import win32process
from datetime import datetime
from win32com.client import Dispatch
from pathlib import WindowsPath as Path
from psutil import Process

from app_manager import AppProcessManager
from app_window import Window
from app_version import ApplicationVersion


class Application:
    """
        Класс абстрактного приложения. Он умеет:
            - узнавать версию приложения из его исполняемого файла;
            - запускать приложение;
            - завершать его работу, через proc.terminate();
            - контролировать состояние приложения, запущено не запущено;
            - управлять окном приложения: сделать верхним, нижним, свернуть, развернуть.
    """

    def __init__(self, executable_path: Path, window_title: str):
        """
        :param executable_path: The path to the application's executable file.
        :param window_title: The title text of the application's main window. Needed to identify the window.
        """
        self.executable_path = executable_path
        if not self.executable_path.exists():
            raise FileNotFoundError(f"File not found: {executable_path}.")
        self.name = self.executable_path.name
        self.version = self._get_version_number()
        self.main_window_title = window_title
        self.process = None
        self.window = None

    def __str__(self):
        return f"Application (name='{self.name}', exe='{self.executable_path}', " \
               f"started='{datetime.fromtimestamp(self.process.create_time())}')"

    def attach_to_process(self, process: Process) -> None:
        """ Подключить класс приложения к уже запущенному процессу.
            Процесс нужно предварительно найти с помощью класса AppProcessManager.
            При подключении

            Процесс приложения должен быть запущен и только одном экземпляре,
            иначе выбрасывается исключение.
            """
        if process.exe() != self.executable_path:
            raise OSError(f'Путь к исполняемому файлу у приложения и процесса не совпадают. '
                          f'Процесс ({process.exe()}). Приложение ({self.executable_path})')
        self.window = [window for window in self._get_all_app_windows() if self.main_window_title in window.title]

    def start_application(self, timeout=20):
        """ Launch the application and wait for its window to appear in the interface.
            Запустить приложение и дождаться появления его окна в интерфейсе. """
        self.process = AppProcessManager.run_application(executable_path=self.executable_path, timeout=timeout)
        self._wait_app_window(timeout=timeout)

    def terminate(self) -> None:
        """ End the application process.
            Завершить процесс приложения. """
        self.process.terminate()
        self.process.wait()

    def terminate_all_instances(self):
        """ Terminate all running instances of the application.
            Завершить все запущенные экземпляры приложения. """
        apps = AppProcessManager.get_processes_by(self.name)
        AppProcessManager.terminate_processes(apps)

    def _get_all_app_windows(self) -> list:
        """ Получить все существующие на данный момент в UI окна,
            созданные процессом приложения.
            Результат вернется в виде списка экземпляров класса 'Window'.
        """

        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == self.process.pid:
                    hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return [Window(hwnd, win32gui.GetWindowText(hwnd)) for hwnd in hwnds]

    def _get_app_main_window(self) -> object:
        """ Получить главное окно приложения. """
        windows = self._get_all_app_windows()
        if windows:
            main_window = [w for w in windows if self.main_window_title in w.title]
            main_window = main_window[1] if main_window else []
        else:
            main_window = []
        return

    def _wait_app_window(self, timeout=10):
        """ Wait until the application window with the specified title appears in the interface.
            Подождать пока окно приложения с указанным заголовком появится в интерфейсе. """
        existing_windows = [window for window in self._get_all_app_windows() if self.main_window_title in window.title]
        ex_win = existing_windows.copy()
        start_quantity = len(existing_windows)
        start_time = time.time()
        while len(existing_windows) != start_quantity + 1:
            time.sleep(1)
            if time.time() - start_time > timeout:
                OSError(f"The application window does not found. Waiting timed out for {timeout} seconds.")
            existing_windows = [window for window in self._get_all_app_windows() if
                                self.main_window_title in window.title]
        print(f"Application window started after {time.time() - start_time} second.")
        self.window = list(set(existing_windows) - set(ex_win))[0]

    def _get_version_number(self) -> object:
        """ Получить из исполняемого файла версию приложения и
            вернуть ее в виде объекта ApplicationVersion. """
        information_parser = Dispatch("Scripting.FileSystemObject")
        version_string = information_parser.GetFileVersion(str(self.executable_path))
        return ApplicationVersion(version_string)
