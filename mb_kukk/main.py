import tkinter as tk
from ui import Application
from utilities import config, logger
import sys
import win32api
import win32event


def start(config, logger):
    # создание уникального мьютекса с именем программы для предотвращение запуска 2-го экземпляра программы
    mutex = win32event.CreateMutex(
        None, False, f"Global\\{config['prog_name']}")
    if win32api.GetLastError() == 183:
        logger.error('Попытка повторно запуска программы')
        sys.exit(1)
    # ___________________________________________________________________________

    try:
        root = tk.Tk()
        Application(root, config)
        logger.info('Program started')
        root.mainloop()

    except Exception as er:
        tb = er.__traceback__
        logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')

    finally:
        win32api.CloseHandle(mutex)
        logger.info('Program closed')


if __name__ == "__main__":
    start(config, logger)
