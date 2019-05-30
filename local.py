# -*- coding: utf-8 -*-
# Автор Михаил Кадай
__version__ = "1.1.3"
try:
    import wmi  # Библиотека для получения запущенных процессов в WINDOWS
except:
    # если WMI Не установлен
    print("- WMI Не установлен!")
    pass
try:
    from psutil import Process  # Библиотека для получения запущенных файлов определённым процессом
except:
    # если psutil Не установлен - установиться автоматически
    print("- psutil Не установлен!")
    pass
# Работа со строками
import re
import os
# Библиотека для работы с аргументами при запуске программы
import argparse
# Прослушка клавиатуры для прекращения работы скрипта
try:
    from pynput.keyboard import Key, Listener
except:
    # если pynput Не установлен - установиться автоматически
    print("- pynput Не установлен!")
from sys import exit
# Работы со временем
import datetime
# Библиотека для работы с json фаылами
from json import load, dumps
# Время сна скрипта в цикле
from time import sleep


class App:
    def __init__(self, ):
        # открытие файла Json
        with open("Setting_find_format_file.json") as setting:
            self.data_work = load(setting, encoding="utf-8")
        setting.close()
        self.result = []
        self.exit = 0

        print("Мониторинг открытых файлов в Microsoft Office Верчия скрипта 1.1.3 \n")
        # Парсинг Аргументов
        parser = argparse.ArgumentParser(description='HELP')
        parser.add_argument('-P', '--print', action='store_true', help='Печать результата работы в консоль')
        parser.add_argument('-S', '--save', action='store_true',  help='Сохранить результат в файл JSON')
        parser.add_argument('-V', '--version', action='store_true',  help='Версия скрипта')
        parser.add_argument('-T', '--time', type=int, default=10,
                            help='ONLY INT VALUE - Время мониторинга (def-10s)')
        parser.add_argument('-Start', '--start', action='store_true', help='Начать работу скрипта')
        args = parser.parse_args()

        # Если параметры не заданы
        pars = 0
        for _, value in args._get_kwargs():
            if value is False:
                pars += 1
            if pars == 4:
                print('Ппожалуйста выполните следующую команду: python local.py -h')
        # печать версии скрипта при заданном параметре -V/--version
        if args.version:
            print("Версия Скрипта: "+__version__)
        # Если задан аргумент сохранения
        if args.save:
            self.data_save = True
        else:
            self.data_save = False
        # Параметр для переодичности выполненния скрипта в секундах DEFAULT - 10 sec
        if args.time:
            self.time = args.time
        # Если задан аргумент печати в консоль
        if args.print:
            self.print = True
        else:
            self.print = False
        # старт программы
        if args.start:
            with Listener(on_release=self.on_release) as self.listener:
                self.listener.join(self.start())

    # Выход с програмы по средствам нажатия ESC
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

    # сама програма ;)
    def start(self):
        while True:
            if self.exit == 1:
                exit(0)
            for process in self.data_work:
                # Запрос через WMI для получения открытых процесов указанных в файле Setting_find_format_file.json
                if process["work"] is True:
                    for Win32_Process in wmi.WMI().query("SELECT * FROM Win32_Process where Name = '"+process['name_process']+"'"):
                        # Проверка запущен ли процесс
                        if len(str(Win32_Process)) > 0:
                            # Дата для сохранения процесса в результат выполнения программы

                            data_save = {
                                "name_process": process['name_process'],  # Имя процесса
                                "working_time": str(re.split(r"[.]", str(datetime.datetime.now()))[0]),  # Время
                                "file_opens": [],                         # Открытые файлы процессом
                                "data": []
                            }
                            # Поиск открытых файлов процессом
                            opens_file_in_process = Process(pid=int(Win32_Process.ProcessId)).open_files()
                            for file in opens_file_in_process:
                                [
                                    data_save["file_opens"].append(str(file.path))
                                    for x in process['file_format_find']
                                    if x in str(file.path)
                                    and file.path not in data_save["file_opens"]
                                    and "~$" not in str(file.path)
                                ]
                            for i in data_save["file_opens"]:
                                data_save["data"].append(
                                    {"Last_Mod": str(re.split(r"[.]",
                                                 str(datetime.datetime.fromtimestamp(os.stat(i).st_mtime)))[0]),
                                     "File": i})
                            # сохранение результатов
                            self.result.append(data_save)
            # Аргумент для печати в консоль
            if self.print is True:
                print('* ------------------------------------------ *')
                for i in self.result:

                    print("* Рабочий процесс : "+str(i['name_process']))
                    print("* Время замера : "+str(i['working_time']))
                    print("* Открытые Файлы: ")
                    for file in i["data"]:
                        print('   -- ФАЙЛ -'+str(file["File"])+" -- Последние изменения Файла - "+str(file["Last_Mod"]))
                    print("\n")

            # Аргумент сохранения информации - результат работы
            if self.data_save is True:
                with open("result.json", "w") as file_save:
                    file_save.write(dumps(self.result, ensure_ascii=False))
                file_save.close()
            # Обнуление результатов
            self.result = []
            # sleep(self.time)
            self.sleep_time()


# Запуск Программы
if __name__ == "__main__":
    App()
