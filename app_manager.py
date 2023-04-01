# coding: utf-8

import os
import time
import psutil
from pathlib import WindowsPath as Path


class AppProcessManager:
    """
        A class with helper methods for interacting with processes in the system.
        Класс с вспомогательными методами для взаимодействия с процессами в системе.
        Он умеет:
            - получать список запущенных процессов по имени приложения
            - запускать приложение и возвращать его процесс
            - завершать указанные процессы
    """

    @staticmethod
    def get_processes_by(partial_name: str) -> list:
        """ Find all processes with the given name and return a list.
            Найти все процессы с заданным именем и вернуть список. """
        procs = [proc for proc in psutil.process_iter()
                 if partial_name.lower() in proc.name().lower()
                 and proc.status() == psutil.STATUS_RUNNING]
        if procs:
            procs.sort(key=lambda x: x.create_time())
        return procs

    @staticmethod
    def run_application(executable_path: Path, timeout: int = 20) -> psutil.Process:
        """ Run the executable, wait for it to run, and return a handle to the running process.
            Запустить исполняемый файл, дождаться запуска и вернуть объект запущенного процесса.
        """
        running_processes = AppProcessManager.get_processes_by(executable_path.name)
        start_quantity = len(running_processes)
        start_time = time.time()
        os.startfile(executable_path)
        while len(running_processes) != start_quantity + 1:
            time.sleep(1)
            if time.time() - start_time > timeout:
                OSError(f"The application did not start. Waiting timed out for {timeout} seconds.")
            running_processes = AppProcessManager.get_processes_by(executable_path.name)
        print(f"Application started after {time.time() - start_time} second.")
        return running_processes[-1]

    @staticmethod
    def terminate_processes(processes: list) -> None:
        """ Terminate the list of processes.
            Завершить список процессов. """
        for proc in processes:
            proc.terminate()
            proc.wait()


if __name__ == "__main__":

    # Запустите код для проверки работы метода.
    # Он вернет список процессов, содержащих указанное имя.
    # Процессы в списке отсортированы по времени запуска,
    # от запущенных раньше к запущенными последним.
    # Чтобы получить объект процесса, который был запущен последним возьмите processes[-1]
    processes = AppProcessManager.get_processes_by(partial_name='Chrome')
    for p in processes:
        print(p)
