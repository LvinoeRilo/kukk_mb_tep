from utilities import logger, timer_decorator, config, check_ping
import sqlite3
import pyodbc


class LabAnalys:

    def __init__(self, config) -> None:
        self.config = config
        self.lab_zn = {}
        self.update_density()

    def update_density(self) -> None:
        try:
            with sqlite3.connect(self.config['local_db']['db_path']) as con:
                curs = con.cursor()
                curs.execute('select kod from table1')
                self.lab_codes = set([code[0] for code in curs if code[0]])
                self._fetch_lab_zn()

                for code, lb_zn in self.lab_zn.items():
                    curs.execute(f'''update table1
                                 set P_nom = {lb_zn}
                                where kod = '{code}' ''')
                    logger.info(f"плотность из {code} обновлена на {lb_zn}")

        except Exception as er:
            logger.info(er)

        finally:
            con.close()

    @timer_decorator
    def _fetch_lab_zn(self) -> None:

        try:

            auth = self.config['lab_db']
            # строка аутентификации
            sql_auth = ";".join([f"{k}={v}" for k, v in auth.items()])

            # проверка доступности базы лаб.анализов
            if not check_ping(config['lab_server']['db_lab']):
                raise pyodbc.Error

            with pyodbc.connect(sql_auth, readonly=True) as con:
                curs = con.cursor()

                for code in self.lab_codes:
                    obkt, res, par = code.split(',')
                    data = curs.execute(
                        f'''select top 1 par_zn
                        from analyse_day
                        where (k_obkt = '{obkt}') and
                        (k_res = '{res}') and
                        (par_cod = '{par}') ''').fetchval()
                    if data and data != 'None':
                        self.lab_zn.update({code: float(data)})

        except pyodbc.Error as er:
            logger.error(
                "Ошибка подключания к базе Лаб.Анализов")

        except Exception as er:
            logger.error(er)

        finally:
            try:
                con.close()
            except:
                pass


LabAnalys(config)
