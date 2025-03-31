import tkinter as tk
from tkinter import messagebox, ttk
import sqlite3
from os import startfile
from datetime import datetime
from tkcalendar import Calendar
from idlelib.tooltip import Hovertip
from pathlib import Path
from update_mb import Update_mb
from excel_generator import DailyReport
from utilities import logger
from time import sleep


def singleton(cls):
    instances = {}

    def wrapper(*args, **kwargs):
        if cls not in instances:
            instances[cls] = cls(*args, **kwargs)
        return instances[cls]
    return wrapper


@singleton
class Application:

    def __init__(self, root, config):
        self.first_start = True
        self.config = config
        self.main_win = root
        # Интервал опроса: (пример  1 сек = 1000)
        self.update_interval = 60000
        self.button_state = tk.IntVar(
            value=self.read_button_state(self.config))
        self.setup_ui()

    # запуск основных параметров интерфейса
    def setup_ui(self):
        self.main_win.title(self.config['prog_name'].upper())
        self.main_win.geometry('480x306')
        self.main_win.iconbitmap('ico\\mnpz.ico')
        self.main_win.resizable(False, False)

        self.setup_menu()
        self.setup_frames()
        self.setup_calendar()
        self.setup_buttons()
        self.setup_bottom_frame()

    # настройка верхнего меню
    def setup_menu(self):
        self.main_menu = tk.Menu(self.main_win)
        self.main_win.config(menu=self.main_menu, bg='#4d4d4d')

        self.file_menu = tk.Menu(tearoff=0)
        self.file_menu.add_command(label='Update')
        self.file_menu.add_command(
            label='Show logs', command=lambda: startfile('logs\\state.log'))
        self.file_menu.add_cascade(
            label='Show database', command=self.show_database)
        self.file_menu.add_separator()
        self.file_menu.add_command(label='Exit', command=self.main_win.destroy)

        # self.edit_menu = tk.Menu(tearoff=0)
        # self.edit_menu.add_cascade(
        #     label='Show database', command=self.show_database)

        self.main_menu.add_cascade(label='File', menu=self.file_menu)
        # self.main_menu.add_cascade(label='Options', menu=self.edit_menu)
        self.main_menu.add_command(label='Help', command=self.about)

    # основные фреймы и их расположение
    def setup_frames(self):
        self.frame_cal = tk.Frame(
            self.main_win, width=200, height=100, bg='#4d4d4d')
        self.frame_cal.place(relx=0, rely=0, anchor='nw')

        self.frame_buttons = tk.Frame(
            self.main_win, width=200, height=100, bg='#4d4d4d')
        self.frame_buttons.place(relx=0.87, rely=0.005)

        self.bottom_frame = tk.Frame(
            self.main_win, width=200, height=100, bg='#383838')
        self.bottom_frame.place(x=0, y=258, width=505)

    # настройки календаря
    def setup_calendar(self):
        date = datetime.now()
        self.cal = Calendar(self.frame_cal, selectmode='day', date_pattern='dd.mm.y',
                            year=date.year, month=date.month, day=date.day, font=("Arial", 16))
        self.cal.pack()

        for row in self.cal._calendar:
            for lbl in row:
                lbl.bind('<Double-1>', self.double_click_date)

    # основные кнопки excel, print, archive, update
    def setup_buttons(self):
        self.excel_ico = self.load_icon('ico\\excel.png')
        self.printer_ico = self.load_icon('ico\\printer.png')
        self.archive_ico = self.load_icon('ico\\archive.png')
        self.update_ico = self.load_icon('ico\\update.png')

        self.create_button(self.frame_buttons, self.excel_ico, 'Create Excel',
                           self.get_cal_for_report, 'Create an Excel file')
        self.create_button(self.frame_buttons, self.printer_ico,
                           'Print Report', self.print_selected_date_report, 'Print selected date excel file')
        self.create_button(self.frame_buttons, self.update_ico,
                           'Update', self.renew_labale_and_db, 'Update')
        self.create_button(self.frame_buttons, self.archive_ico,
                           'Open Archive', self.open_arch, 'Open_archive')

    # размер иконок
    def load_icon(self, path):
        icon = tk.PhotoImage(file=path)
        return icon.subsample(12, 12)

    # метод для настройки основных кнопок excel, print, archive, update
    def create_button(self, frame, icon, text, command, tooltip):
        button = tk.Button(frame, image=icon, text=text, font='arial 14', bg='#4d4d4d',
                           activebackground='#4d4d4d', cursor='hand2', border=0, command=command)
        button.pack(pady=10, anchor=tk.CENTER)
        Hovertip(button, tooltip, hover_delay=200)

    # настройка нижней панели интерфейса
    def setup_bottom_frame(self):

        style = ttk.Style()
        style.configure("WhiteText.TCheckbutton",
                        foreground="white", background='#383838', font=('arial', 14))

        self.timer_label = tk.Label(
            self.bottom_frame, font='arial 14', background='#383838', fg='white')
        self.timer_label.pack(side='left', padx=5)

        self.last_update_label = tk.Label(
            self.bottom_frame, text='Last update at: ', font='arial 14', background='#383838', fg='white')
        self.last_update_label.pack(side='right', padx=35)

        self.print_check = ttk.Checkbutton(
            self.bottom_frame, text='Print', style="WhiteText.TCheckbutton", onvalue=1, offvalue=0, variable=self.button_state, command=self.save_button_state)
        self.print_check.pack(side='left', padx=5)

        self.last_update_label.after_idle(self.update)
        self.timer_label.after_idle(self.timer)

    # Обработка двойного нажатия на дату календаря, для открытия отчета за данную дату
    def double_click_date(self, event):
        file = f'{self.cal.get_date()}_report.xlsx'
        if (path := Path().cwd() / f'reports/{file}').exists():
            startfile(path)
        else:
            messagebox.showerror('error', f'file: {file} doesnt exist')

    # Информация о программе
    def about(self):
        messagebox.showinfo(
            'Help', f'''{self.config['prog_name']}\n\n
            1) Double click on date to open selected date report\n
            2) \n
            3) ''')

    # Запуск расчетов м\б и обновление даты в интерфейсе
    def renew_labale_and_db(self):
        Update_mb(self.first_start)
        self.last_update_label['text'] = f'Last update at: {
            datetime.now().strftime("%H:%M:%S")}'
        self.first_start = False

    # запуск расчетов каждые __ сек
    def update(self):
        self.last_update_label.after(self.update_interval, self.update)
        self.renew_labale_and_db()

    # часы и таймер
    def timer(self):
        self.timer_label.after(1000, self.timer)
        self.timer_label['text'] = datetime.now().strftime('%H:%M:%S')

    # Открытие архивной базы
    def open_arch(self):
        try:
            month, year = self.cal.get_displayed_month()
            tb_name = f'Archive_{year}_{month:02}'
            self.show_database(tb_name)
        except Exception as er:
            messagebox.showerror('Error', f"no such table {tb_name}")
            tb = er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')

    # Формирование и отображение базы за выделенную дату в календаре
    def show_database(self, tb_name=None):
        table = tb_name or self.config['local_db']['table_name']

        with sqlite3.connect(self.config['local_db']['db_path']) as connect:
            curs = connect.cursor()
            cur_date = self.cal.get_date().split('.')
            cur_day = cur_date[0][1] if cur_date[0].startswith(
                '0') else cur_date[0]

            try:
                # формирование строки для sql запроса
                sql_data = f'''select ima as Наименование,shifr as Шифр,
                    tzn as 'Текущ.знач', sz_hour as 'Ср.знач', 
                    sz_hour_m as 'Ср.знач масса',p_nom as 'Пасп.плотн', 
                    density as 'Привед.плотн',

                    sz_n{cur_day}_1V as 'Вахта 1', 
                    sz_n{cur_day}_2V as 'Вахта 2', 
                    sz_n{cur_day}_3V as 'Вахта 3',

                    (sz_n{cur_day}_1V + sz_n{cur_day}_2V + sz_n{cur_day}_3V) as Сутки,

                    m_n{cur_day}_1V as 'Вахта 1 масса',
                    m_n{cur_day}_2V as 'Вахта 2 масса',
                    m_n{cur_day}_3V as 'Вахта 3 масса',

                    (m_n{cur_day}_1V + m_n{cur_day}_2V + m_n{cur_day}_3V) as 'Сутки масса',

                    {'+'.join([f'Sz_N{d}_{w}V' for d in range(1, int(cur_day) + 1)
                               for w in range(1, 4)])} as Месяц,
                               
                    {'+'.join([f'M_N{d}_{w}V' for d in range(1, int(cur_day) + 1)
                                   for w in range(1, 4)])} as 'Месяц масса'
                     
                    from {table}'''

                columns = ','.join(item[0]
                                   for item in curs.execute(sql_data).description)

                db_view = tk.Tk()
                db_view.title(table)
                db_view.geometry('700x500')

                # Формирование заголовков
                tree = ttk.Treeview(
                    db_view, columns=columns.split(','), show='headings')
                for heading in columns.split(','):
                    tree.heading(heading, text=heading)

                # Заполнение данными
                data_from_db = curs.execute(sql_data)
                for data in data_from_db.fetchall():
                    tree.insert('', tk.END, values=data)

                tree.pack(fill=tk.BOTH, expand=True)

                scrollbar = ttk.Scrollbar(
                    db_view, orient=tk.HORIZONTAL, command=tree.xview)
                tree.configure(xscroll=scrollbar.set)
                scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

                update_button = tk.Button(db_view, text='Update', font='Arial, 15', width=500,
                                          command=lambda: self.update_data(tree, curs, sql_data, table))
                update_button.pack()

                db_view.mainloop()

            except Exception as er:
                messagebox.showerror('Error', er)

    # Обновление базы в методе show_database
    def update_data(self, tree, curs, columns, table):
        for row in tree.get_children():
            tree.delete(row)
        data_from_db = curs.execute(columns)
        for data in data_from_db.fetchall():
            tree.insert('', tk.END, values=data)

    # чтения состояния печати принтера в таблице printer_settings, 1 вкл 0 выкл
    @staticmethod
    def read_button_state(config):
        try:
            with sqlite3.connect(config['local_db']['db_path']) as con:
                con.row_factory = sqlite3.Row
                curs = con.cursor()
                return curs.execute('''select print_state from print_settings''').fetchone()['print_state']
        except Exception as er:
            tb = er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')
        finally:
            con.close()

    # сохранение состояния печати принтера в таблице printer_settings, 1 вкл 0 выкл
    def save_button_state(self):
        state = str(not self.read_button_state(self.config))
        try:
            with sqlite3.connect(self.config['local_db']['db_path']) as con:
                curs = con.cursor()
                curs.execute(
                    f'''update print_settings set print_state = {state}''')
        except Exception as er:
            tb = er.__traceback__
            logger.error(f'error_on_line: {tb.tb_lineno} | {er} ')
        finally:
            con.close()

    # получение даты для формирования отчета excel
    def get_cal_for_report(self):
        date = datetime.strptime(self.cal.get_date(), '%d.%m.%Y')
        if date > datetime.now():
            messagebox.showinfo('Info', 'I can\'t predict the future')
        else:
            try:
                DailyReport(date)
            except Exception as er:
                messagebox.showerror(f'Error', er)

    # Печать отчета за выделенную дату
    def print_selected_date_report(self):
        date = datetime.strptime(self.cal.get_date(), '%d.%m.%Y')
        if date > datetime.now():
            messagebox.showinfo('Info', 'I can\'t predict the future')
        else:
            try:
                DailyReport(date, 1)
            except Exception as er:
                messagebox.showerror(f'Error', er)
