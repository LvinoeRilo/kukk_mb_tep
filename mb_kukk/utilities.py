from loguru import logger
import json
from pathlib import Path
import time
from functools import wraps
import sqlite3

import subprocess


# Инициализация логгера
logger.add("logs/state.log", rotation="1 MB",
           format="{time:DD-MM-YY HH:mm:ss}|{level}|{message}|{exception}",
           compression="zip")


# конфиг
class Cnf:
    def __init__(self, config_file):
        self.__config_file = config_file
        self.__data = {}
        try:
            # parse the json file and forms dictionary for further reading
            with open(self.__config_file, encoding='utf-8') as config:
                self.__data = json.load(config)
        except Exception as er:
            logger.error(er)
            raise er

    # returns the value of the json file
    def __getitem__(self, key, default=None):
        return self.__data.get(key, default)

    # returns dictionary with all configs
    @property
    def config(self):
        return self.__data


# декораток проверки времени выполнения
def timer_decorator(method):
    """Декоратор для измерения времени выполнения метода."""

    @wraps(method)
    def wrapper(*args, **kwargs):
        start_time = time.time()  # Засекаем время начала выполнения
        result = method(*args, **kwargs)  # Выполняем метод
        end_time = time.time()  # Засекаем время окончания выполнения
        elapsed_time = end_time - start_time  # Вычисляем время выполнения
        print(
            f"Метод {method.__name__} выполнился за {elapsed_time:.4f} секунд.")
        return result  # Возвращаем результат метода

    return wrapper

# сканирует базу и вытаскивает нужные шифры для дальнейшего опроса
# записывает в нужный конфиг файл


class TagScanner:

    def __init__(self, file_path):
        self.__write_config(file_path, self.scan_tags())

    def scan_tags(self):
        try:
            path = Path.cwd() / config['local_db']['db_path']
            with sqlite3.connect(path) as con:
                curs = con.cursor()
                res = curs.execute('''select Shifr_gt, Shifr_gt1,
                                    Shifr_gt2, Shifr_T, Shifr_P
                                    from table1''').fetchall()
                tags = set([el for lst in res for el in lst if el])
                return list(tags)
        except Exception as er:
            logger.error(er)
        finally:
            con.close()

    @staticmethod
    def __write_config(file_path, tags):
        with open(file_path, 'r', encoding='UTF-8') as file:
            res = json.load(file)
        res['tags'] = tags
        with open('new.json', 'w', encoding='UTF-8') as file:
            json.dump(res, file, indent=4, ensure_ascii=False)


# проверка доступности адреса аналог ping
def check_ping(host: str) -> bool:
    ping = subprocess.run(["ping", "-n", "1", host],
                          stdout=subprocess.PIPE)
    if not ping.returncode:
        return True
    return False


# Конфиг файл
config = Cnf("config.json")


# import re

# formula = '''cipher['FC7011'] + cipher['жопа026A'] + cipher['FI3026_A']'''

# pattern = r"cipher\['([а-яА-Яa-zA-Z0-9_]+)'\]"

# result = re.findall(pattern, formula)

# print(result)
