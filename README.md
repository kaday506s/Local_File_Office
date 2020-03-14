Автор: Михаил Кадай 
Почта: kaday506@gmail.com

Требования для работы скрипта: 
1)	Python 3.6 -  https://www.python.org/downloads/release/python-364/ 

2) win32core – Требование WMI.
Установка (pip install win32core)
 
3)  win32compat – Требование WMI.
Установка (pip install win32compat)

4)  WMI – Для получения списка запущенных процессов в системе WINDOWS.
Установка (pip install WMI) 

5) psutil – Мониторинг открытых файлов в системе.
Установка (pip install psutil)

6) pynput – Прослушка клавиатуры для прекращения работы скрипта 
Установка (pip install pynput)

* Команда для установки зависимотей :
pip install -r requirements.txt

Описание скрипта: 
Файлы: 
1) Setting_find_format_file.json Настройки мониторинга указанных программ ( По-умолчанию вписаны Процессы WORD / ECXEL / PowerPoint )
Имеет вид:
***
    // Имя ЕХЕ для мониторинга
    "name_process": "WINWORD.EXE", 
    // отслеживать работу данного процесса
     "work": true,
    // Форматы с которыми работает процесс 
    "file_format_find": [".docx", ".docm", ".doc", ".dotx", ".dotm"] 

2) local.py – скрипт для мониторинга процессов ОС 
- Параметры запуска: 
1. –P / --print  - без аргументов , если параметр задан будет включена печать в консоль.
Пример печати в консоль:
*** 
    * Process Working : WINWORD.EXE
    * Time now : 2019-04-17 20:32
    * Opens File: 
    -- (file path work…)

2. –S/--save  - для сохранения информации в формат json.
3. –T/--time принимает Число – секунды с какой периодичности будет работать скрипт , если параметр не задан по-умолчанию будет поставлено 10 секунд.
4. –Start/--start  - старт работы скрипта. 
5. –h/--help  - Вывод справки аргумента. 


Пример запуска: 
python local.py --print --start -T 12
    - Для прекращения работы скрипта нажмите ESC******
