import pyodbc

import json
from datetime import datetime
from utilities import logger


class Anal:
    '''Класс для обновления лаб.анализа объекта D_502_KUG'''

    def __init__(self):
        # Иницализация конфига и стартовых атрибутов

        # формирование словаря параметров для строки соединяния с БД
        with open('config.ini') as cnfg:
            reader = [i.strip() for i in cnfg.readlines() if i.strip()]
            self.config = {s.split('=')[0]: s.split('=')[1] for s in reader}

        self.sqlstr = 'DRIVER=%s;SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (
            self.config['DRIVER'], self.config['SERVER_LAB'],
            self.config['DATABASE_LAB'], self.config['UID_LAB'],
            self.config['PWD_LAB'])

        # Кол объекта и ресурса
        self.k_obkt = 2019
        self.k_res = 1368
        # ________________
        # пар_коды для опроса

        self.par_cods = (40, 22, 23, 26, 27, 625, 668,
                         626, 50, 53, 78, 62, 41, 43)

        # пустой словарь и процедура set_par для дальнейшего его заполнения
        self.dict_of_par_cods = {}
        self.set_par()

    def set_par(self):

        try:
            # Попытка соединения с ежедневной базой анализов
            con = pyodbc.connect(self.sqlstr, readonly=True)
            curs = con.cursor()

            sql = f'''select top 1 * from analyse_day where
                                          (k_obkt = {self.k_obkt}) and
                                          (k_res = {self.k_res})'''

            # проверка наличия актуальных анализов, если актуальные анализы отсутствуют,
            # словрь dict_of_par_cods заполняется предыдущими значениями используя json файл
            record_set = list(curs.execute(sql))

            if not record_set:
                with open('D_502_kug.json', encoding='utf-8') as file:
                    self.dict_of_par_cods = json.load(file)
                print('there are no new analisys')
                with open('report_error.txt', 'a', encoding='utf-8') as log:
                    log.write(
                        f"{datetime.today().strftime('%d.%m.%Y %H:%M:%S')}: there are no new analisys, old are get\n")

            # заполнение словаря значениями par_cods из базы анализов
            else:
                for i in self.par_cods:
                    k = curs.execute(
                        sql + f' and (par_cod = {i}) order by dtt desc')
                    k_v = list(k.fetchone())
                    # заполнение словаря значениями par_cods из базы анализов, в случае некорректных значений - 0
                    self.dict_of_par_cods.update(
                        {k_v[18].strip(): float(k_v[20]) if in_numer(k_v[20]) else 0})

                    # Формирование json файла
                    with open('D_502_kug.json', 'w', encoding='UTF-8') as js:
                        json.dump(self.dict_of_par_cods, js,
                                  indent=4, ensure_ascii=False)

                with open('report_error.txt', 'a', encoding='utf-8') as log:
                    log.write(
                        f"{datetime.today().strftime('%d.%m.%Y %H:%M:%S')}: Analysis updated \n")

        except Exception as err:
            with open('report_error.txt', 'a', encoding='utf-8') as log:
                log.write(
                    f"{datetime.today().strftime('%d.%m.%Y %H:%M:%S')}: Error in module lab_anal - {err}. \n")

        finally:
            print('done')
            con.close()

    def __iter__(self):
        yield from self.dict_of_par_cods.values()

    def __getitem__(self, value):
        return self.dict_of_par_cods[value]

# with open('D_502_kug.json', encoding = 'utf-8') as file:
# lab = json.load(file)
# consts = (16.04,28.05,30.07,42.08,44.1,56.11,72.1,86.2,2.02,34.082,32,28.01,28.01,44.01)
# formula = map(lambda x: x[0] * x[1], zip(lab.values(), consts))
# print(sum(formula) / 100/22.4)
