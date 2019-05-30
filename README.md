Автор: Михаил Кадай 
Почта: kaday506@gmail.com
Требования для работы скрипта: 

Python 3.6 - Обязательное условие без него не работают библиотеки 
Сам Python. 
1)	Python 3.6 -  https://www.python.org/downloads/release/python-364/
Библиотеки работают только с Python3.6: 
1)  WMI – Для получения списка запущенных процессов в системе WINDOWS.
Установка (pip install WMI  /  pip3 install WMI  /  python3 –m pip install WMI) 

2)  win32compat – Требование WMI.
Установка (pip install win32compat / pip3 install win32compat / 
 python3 –m pip install win32compat)

3) psutil – Мониторинг открытых файлов в системе.
Установка (pip install psutil / pip3 install psutil / python3  –m pip install psutil)

4) win32core – Требование WMI.
Установка (pip install win32compat / pip3 install win32compat / 
 python3 –m pip install win32compat)

5) pynput – Прослушка клавиатуры для прекращения работы скрипта 
Установка (pip install pynput / pip3 install pynput / 
 python3 –m pip install pynput)



Описание скрипта: 
Для работы скрипта два файлы должны лежать в одной директории 
Файлы: 
1) Setting_find_format_file.json Настройки мониторинга указанных программ ( По-умолчанию вписаны Процессы WORD / ECXEL / PowerPoint )
Имеет вид:
// Имя ЕХЕ для мониторинга
"name_process": "WINWORD.EXE", 
// отслеживать работу данного ЕХЕ
 "work": true,
 // Форматы с которыми работает процесс 
"file_format_find": ["docx", "docm", "doc", "dotx", "dotm"] 

2) local.py – скрипт для мониторинга процессов ОС 
- Параметры запуска: 
1. –P / --print  - без аргументов , если параметр задан будет включена печать в консоль.
Формат печати в консоль:

* Process Working : WINWORD.EXE
* Time now : 2019-04-17 20:32:17
* Opens File: 
-- (file path work…)

2. –S/--save  - без аргументов , если параметр задан будет включена запись в файл в формате JSON.
3. –T/--time принимает Число – секунды с какой периодичности будет работать скрипт , если параметр не задан по умолчанию будет поставлено 10 секунд.
4. –Start/--start  - старт работы скрипта. 
5. –V/--version -  Вывод версии скрипта в консоль. 
	6. –h/--help  - Вывод справки аргумента. 


Пример запуска: 
python local.py --print --start -T 12
           - Для прекращения работы скрипта нажмите ESC******
