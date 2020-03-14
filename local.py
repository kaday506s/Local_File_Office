# -*- coding: utf-8 -*-
# Автор Михаил Кадай
__version__ = "1.1.4"
try:
    import wmi
except ImportError:
    raise Exception("- WMI Не установлен!")
try:
    from psutil import Process
except ImportError:
    raise Exception("- psutil Не установлен!")
try:
    from pynput.keyboard import Key, Listener
except ImportError:
    raise Exception("- pynput Не установлен!")

from sys import exit
from datetime import datetime
from json import load, dumps
from time import sleep
from os import stat as os_stat
import argparse


class App:
    wmi_sql_select_process = "SELECT * FROM Win32_Process where Name = '{}'"

    def __init__(self, ):
        # открытие файла Json для чтения настроек
        with open("Setting_find_format_file.json") as setting:
            self.data_work = load(setting, encoding="utf-8")

        self.result = []
        self.exit = 0

        print(f" * Мониторинг открытых файлов в "
              f"Microsoft Office Верчия скрипта {__version__} \n")

        # Парсинг Аргументов
        parser = argparse.ArgumentParser(description='HELP')

        parser.add_argument('-P', '--print',
                            action='store_true',
                            help='Печать результата работы в консоль')
        parser.add_argument('-S', '--save',
                            action='store_true',
                            help='Сохранить результат в файл JSON')
        parser.add_argument('-T', '--time',
                            type=int,
                            default=10,
                            help='ONLY INT VALUE - Время мониторинга (def-10s)')
        parser.add_argument('-Start', '--start',
                            action='store_true',
                            help='Начать работу скрипта')
        args = parser.parse_args()

        args_if_false = [value for _, value in args._get_kwargs() if value is False]
        if len(args_if_false) == 3:
            print('* Ппожалуйста выполните следующую команду: python local.py -h')

        self.data_save = args.save
        self.time = args.time
        self.print = args.print

        # старт программы
        if args.start:
            with Listener(on_release=self.on_release) as self.listener:
                self.listener.join(self.start())

    # Выход по нажатию ESC
    def on_release(self, key):
        if key == Key.esc:
            print('Выход из скрипта')
            self.listener.stop()
            self.exit = 1

    # Выход из програмы
    def sleep_time(self):
        for i in range(self.time):
            if self.exit == 1:
                exit(0)
            sleep(1)

    @classmethod
    def print_result(cls, result_data):
        print('* ------------------------------------------ *')
        for res in result_data:
            print(f"* Рабочий процесс : {res['name_process']}\n"
                  f"* Время замера : {res['working_time']}\n"
                  f"* Открытые Файлы: ")
            for file in res["data"]:
                print(f'   -- ФАЙЛ -{file["File"]}  '
                      f' -- Последние изменения Файла '
                      f' - {file["Last_Mod"]}')

    @classmethod
    def save_result_in_json(cls, result_data):
        with open("result.json", "w") as file_save:
            file_save.write(dumps(result_data, ensure_ascii=False))

    # сама програма ;)
    def start(self):
        while True:
            for process in self.data_work:
                # Запрос через WMI для получения открытых процессов
                # указанных в файле Setting_find_format_file.json
                if process["work"] is True:
                    for Win32_Process in wmi.WMI().query(
                         self.wmi_sql_select_process.format(process['name_process'])):
                        if Win32_Process:
                            data_save = {
                                "name_process": process['name_process'],
                                "working_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                "files_opens": [],
                                "data": []
                            }
                            # Поиск открытых файлов id процессом
                            opens_file_in_process = Process(pid=int(Win32_Process.ProcessId))\
                                .open_files()

                            for file in opens_file_in_process:
                                [
                                    data_save["files_opens"].append(
                                        file.path
                                    )
                                    for x in process['file_format_find']
                                    if file.path.endswith(x)
                                    and file.path not in data_save["files_opens"]
                                    and "~$" not in file.path
                                ]
                            # ToDo change
                            for file in data_save["files_opens"]:
                                data_save["data"].append(
                                    {"Last_Mod": datetime
                                        .fromtimestamp(os_stat(file).st_mtime)
                                        .strftime("%Y-%m-%d %H:%M"),
                                     "File": file})

                            self.result.append(data_save)

            # Аргумент для печати в консоль
            if self.print:
                self.print_result(self.result)

            # Аргумент сохранения информации
            if self.data_save:
                self.save_result_in_json(self.result)

            # Обнуление результатов
            self.result = []
            self.sleep_time()


if __name__ == "__main__":
    App()
