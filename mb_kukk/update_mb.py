import sqlite3
from utilities import logger, timer_decorator, config
from lab_analysis import LabAnalys
from datetime import datetime, timedelta
import shutil


class Update_mb:
    __slots__ = ('watch', 'tags', 'curr_date_time')

    # инициализация основных атрибутов и функций
    def __init__(self, first_start):
        # определение текущего номера вахты при инициализации
        self.curr_date_time = datetime.now()
        if first_start:
            self.restore_data()
        self.watch = self.current_watch()
        self.check_if_periods_change()
        self.update_tags(config['tags'])
        self.update_mb()

        # Обновление времени в базе
        self.update_db_time()

    # обновление шифров из data.db
    def update_tags(self, tags):
        try:

            with sqlite3.connect(config['tag_db']['db_path']) as con:
                curs = con.cursor()
                sql_string = f'''SELECT shifr,tzn FROM
                {config['tag_db']['table_name']}
                WHERE shifr IN ({','.join(['?']*len(tags))})'''
                res = curs.execute(sql_string, tags).fetchall()
                self.tags = dict(res)

        except Exception as er:
            tb = er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')

        finally:
            try:
                con.close()
            except Exception as er:
                tb = er.__traceback__
                logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')

    # Обнуление вахтовых значений базы данных мат балласа
    def reset_to_zero(self):
        try:
            with sqlite3.connect(config['local_db']['db_path']) as con:
                curs = con.cursor()
                for date in range(1, 32):
                    for watch in range(1, 4):
                        curs.execute(f'''UPDATE table1 SET
                                    Sz_N{date}_{watch}V = 0,
                                    M_N{date}_{watch}V = 0,
                                    P_{date} = 0''')
        except Exception as er:
            tb = er.__traceback__
            logger.error(
                f'error_on_line: {tb.tb_lineno} |Ошибка в модуле reset_to_zero: {er} ')

        finally:
            logger.info('Значения вахт обнулены')

    # обновление значений базы материального баланса
    def update_mb(self):
        try:
            # подключение базы через крнтекстный менеджер
            with sqlite3.connect(config['local_db']['db_path']) as con:
                # Фабрика строк
                con.row_factory = sqlite3.Row
                curs = con.cursor()
                curs.execute(f'''select * from
                {config["local_db"]['table_name']}''')

                tag = self.find_tag

                for r in curs.fetchall():

                    try:
                        match r['Prizn']:

                            case 'NEFT_OUT_18':
                                # формула береться напрямую из базы, второй аргумент в eval обеспечивает безопасность
                                formula = eval(r['K'], {'tag': tag})
                                dens = self.count_density(
                                    r['P_nom'], tag('TI3135'), curs, r['Shifr'])

                                if formula >= 0:

                                    curs.execute(f'''update table1 set tzn =
                                    {formula} where Shifr = '{r['Shifr']}' ''')

                                    self.count_av_hour(
                                        r['sz_hour'], formula, r['n'], curs, r['Shifr'])
                                    self.count_av_mass_hour(
                                        r['sz_hour_m'], formula, r['n'], dens, curs, r['Shifr'])
                                else:
                                    curs.execute(
                                        f'''update table1 set tzn = -99999 where shifr = '{r['Shifr']}' ''')

                            case 'UROV':

                                formula = tag(r['Shifr_gt'])
                                if formula >= 0:
                                    curs.execute(f'''update table1 set tzn = {formula}
                                                where shifr = '{r['Shifr']}' ''')
                                else:
                                    curs.execute(
                                        f'''update table1 set tzn = -99999 where shifr = '{r['Shifr']}' ''')

                            case _:
                                pass

                        curs.execute(
                            f'''update table1 set n = {r['n'] + 1} where not tzn = -99999 and Shifr = '{r['Shifr']}' ''')

                    except Exception as Er:
                        tb = Er.__traceback__
                        logger.error(f'error_on_line: {tb.tb_lineno} | {Er} ')

        except Exception as Er:
            tb = Er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {Er} ')

        finally:
            print(f'{self.curr_date_time.strftime(
                '%d.%m.%Y %H:%M')}: Значения обновлены')
            con.close()

    # Определение номера вахты
    def current_watch(self, dt='now'):
        # Получаем текущее время
        if dt == 'now':
            now = self.curr_date_time
        else:
            now = dt
        current_hour = now.hour
        # Определяем номер вахты

        # Определяем номер вахты
        if 0 < current_hour <= 8:  # 00:01 - 08:00 → вахта 1
            return 1
        elif 8 < current_hour <= 16:  # 08:01 - 16:00 → вахта 2
            return 2
        else:  # 16:01 - 00:00 → вахта 3
            return 3

    # высчитывание плотности
    @staticmethod
    def count_density(dens, temp, cursor, tag):
        if temp != -99999:
            result = (dens - (0.001828 - 0.00132 * dens) * (temp - 20))
            cursor.execute(
                f"update table1 set density = {result} where shifr = '{tag}' ")
            return result

    # высчитывание среднего значения
    @staticmethod
    def count_av_hour(av_h, formula, iteration, curs, tag):

        result = av_h + (formula - av_h) / (iteration + 1)
        curs.execute(f'''update table1 set sz_hour =
        {round(result, 3)} where shifr = "{tag}"''')

    # высчитывание среднего массового значения
    @staticmethod
    def count_av_mass_hour(av_m_h, formula, iteration, density, curs, tag):

        result = av_m_h + ((formula * density) - av_m_h) / (iteration + 1)
        curs.execute(f'''update table1 set sz_hour_m =
        {round(result, 3)} where shifr = "{tag}"''')

    # Вывод шифра из словаря __tags
    def find_tag(self, tag):
        if not hasattr(self, 'tags'):
            logger.error('Аттрибут tags не был создан')
            return -99999
        result = self.tags.get(tag, -99999)
        if result == -99999:
            logger.error(f'Шифр {tag} не найден, возращено {result}')
            return result
        return result

    # вывод даты и времени последнего обновления в datetime формате
    def last_datetime(self):
        try:
            with sqlite3.connect(config['local_db']['db_path']) as con:
                con.row_factory = sqlite3.Row
                curs = con.cursor()
                curs.execute('select lastupdate from update_time')
                last_time = curs.fetchone()
                return datetime.strptime(
                    last_time['lastupdate'], '%Y-%m-%d %H:%M:%S')

        except Exception as Er:
            tb = Er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {Er} ')

        finally:
            con.close()

    # Обновление времени в базе
    def update_db_time(self, dt='now'):
        # проверка аргумента dt: добавлять ли текущую дату и время либо пользовательскую
        if dt == 'now':
            cur_time = self.curr_date_time.strftime('%Y-%m-%d %H:%M:%S')
        else:
            cur_time = dt
        # Обновление текущего времени и даты
        try:
            with sqlite3.connect(config['local_db']['db_path']) as con:
                curs = con.cursor()
                curs.execute(
                    f'''update update_time set lastupdate = '{cur_time}' ''')
        except Exception as er:
            tb = er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')
        finally:
            con.close()

    # проверка на -99999
    def no_minus_nines(self, *tags):
        return all(map(lambda x: x != -99999, tags))

    # проверка смены часа
    def hour_change(self, cur_time, restore=False):
        db_time_date = self.last_datetime()
        # обновление плотностей лаб.анализов

        if (db_time_date.hour != cur_time.hour) or (db_time_date.date() != cur_time.date()):
            if not restore:
                LabAnalys(config)

            try:
                with sqlite3.connect(config['local_db']['db_path']) as con:
                    con.row_factory = sqlite3.Row
                    curs = con.cursor()
                    cur_sz = f"Sz_N{db_time_date.day}_{self.watch}V"
                    cur_m = f"M_N{db_time_date.day}_{self.watch}V"
                    for row in curs.execute(f'''select shifr, prizn, tzn,
                                            sz_hour, sz_hour_m, {
                        cur_sz}, {cur_m}
                                            from table1''').fetchall():

                        if 'UROV' != row['Prizn']:
                            curs.execute(
                                f'''update table1 set {cur_sz} =  {row[cur_sz] + row['sz_hour']} where shifr = '{row['Shifr']}' ''')
                            curs.execute(
                                f'''update table1 set {cur_m} =  {row[cur_m] + row['sz_hour_m']} where shifr = '{row['Shifr']}' ''')
                        else:
                            curs.execute(f'''update table1 set {cur_sz} = {row['tzn']}
                                        where Shifr = '{row['Shifr']}' ''')

                    # Обнуление итератора и среднего часового значения если restore == False
                    if not restore:

                        curs.execute(
                            '''update table1 set n = 0, Sz_hour = 0, Sz_hour_m = 0''')

            except Exception as Er:
                tb = Er.__traceback__
                logger.error(
                    f'error_on_line: {tb.tb_lineno} | Ошибка при обновлении при переходе часа: {Er} ')
                logger.error(f' {Er}')

            finally:
                con.close()
                logger.info('Переход часа')

    # проверка смены дня
    def day_change(self, cur_time):

        if self.last_datetime().day != cur_time.day:
            try:
                shutil.copy(config['local_db']['db_path'],
                            f'db\\backup\\backup_{config['local_db']['database']}')
                # вставить сюда функцию для распечатки

            except Exception as er:
                tb = er.__traceback__
                logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')

            finally:
                logger.info('Переход суток. Бэкап текущей базы создан')

    # проверка смены месяца
    def month_change(self, cur_time):
        if (prev_d := self.last_datetime()).month != cur_time.month:

            arch_name = f'''Archive_{prev_d.year}_{prev_d.month:02}'''
            try:
                with sqlite3.connect(config['local_db']['db_path']) as con:
                    curs = con.cursor()
                    curs.execute(f'''create table {arch_name}
                    as select * from table1''')
                self.reset_to_zero()

            except Exception as er:
                logger.error(f'Ошибка создания архива {arch_name}: {er}')

            finally:
                logger.info(f'Переход месяца')
                logger.info(f'Архив {arch_name} создан')
                con.close()

    # общая проверка перехода часа, суток и месяца
    def check_if_periods_change(self):
        # определение текущего времени и даты
        cur_date = self.curr_date_time
        # проверка смены часа, суток и месяца
        self.hour_change(cur_date)
        self.day_change(cur_date)
        self.month_change(cur_date)

    # Восстановление данных при разрыве в 1 час и более
    def restore_data(self):
        from ui import Application
        app = Application()

        try:
            last_date = self.last_datetime()
            cur_date = self.curr_date_time

            # лейбл для анимации восстановления
            restore_label = 'Restore    '

            # проверка разбежки времени
            while (cur_date - last_date).days * 86400 + (
                    cur_date - last_date).seconds >= 0:

                # если разбежка меньше часа прервать восстановление
                if (cur_date - last_date).days * 86400 + (
                    cur_date - last_date).seconds < 3600 and (
                        cur_date.hour == last_date.hour):

                    logger.info(
                        f"Восстановление базы данных не требуется:\nТекущая дата - {cur_date}\nДата последнего обновления {last_date}")

                    break

                # анимация
                if restore_label == 'Restore....':
                    restore_label = 'Restore    '

                self.watch = self.current_watch(last_date)

                # методу hour_change передается zero_values = False во избежание обнуления средних значений
                self.hour_change(cur_date, restore=True)

                # проверка смены для и месяца
                if (last_date + timedelta(hours=1)).day > last_date.day:
                    self.day_change(cur_date)
                if (last_date + timedelta(hours=1)).month > last_date.month:
                    self.month_change(cur_date)

                # обновление лейбла интерфейса
                app.last_update_label['text'] = restore_label
                app.last_update_label.update_idletasks()
                restore_label = restore_label.replace(' ', '.', 1)

                # +1 час времени от последнего обновления в том числе в базе
                last_date += timedelta(hours=1)
                self.update_db_time(last_date.strftime('%Y-%m-%d %H:%M:%S'))

        except Exception as er:
            tb = er.__traceback__
            logger.error(
                f'error_on_line: {tb.tb_lineno} |Ошибка восстановления {er} ')

        finally:
            pass
