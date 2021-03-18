# python_1c_updater
КРАТКАЯ ДОКУМЕНТАЦИЯ К ИСПОЛЬЗОВАНИЮ СКРИПТА UpdateSQLbase1C.py
Скрипт можно использовать только с базами находящимися на поддержке.
Скрипт использует модули os, time, psutil, pythoncom, win32com.client, pywinauto
Если нужна возможность завершать принудительно все запущенные платформы и процессы rphost, нужно запускать скрипт с повышенными правами и модули устанавливать аналогичным образом.
Скрипт нельзя использовать через Планировщик заданий Windows, т.к. используется модуль pywinauto для обнаружения и закрытия окон возникающих в процессе обновления. Для этого используются функции window0 - window4, все функции написаны с помощью утилиты Swapy и возможно на вашей ОС их придётся переделывать. Скрипт писался и тестировался на ОС Windows server 2012 R2.
Скрипт достаточно примитивен, но позволяет сэкономить много времени на ручном обновлении баз.

Для использования необходимо исправить переменные l_fld - строка пути к каталогу, где будут файлы логов (по умолчанию D:\DST\1Clog), pathZup - строка пути к каталогу содержащему файлы обновлений для конфигурации Зарплата и управление персоналом (по умолчанию D:\DST\HRM). pathBuh - строка пути к каталогу содержащему файлы обновлений для конфигурации Бухгалтерия предприятия (по умолчанию D:\DST\BUH). Конфигурация Управление торговлей в данной версии не тестировалось и не использовалось, функционал можно добавить по аналогии с другими конфигурациями. В каталоге обновлений, кроме файла .cfu обязательно должен быть файл UpdInfo.txt. Список баз для обновления берётся из файла bases.txt который должен лежать в одном каталоге со скриптом обновления и иметь следующий формат:                                                                                                 

<имя сервера>,<имя базы>,<имя пользователя>,<пароль>

например:
SERVER-1C,MyHomeBUH,Buhgalter,1234qwer
SERVER-1C,MyHomeZUP,Admin,123-qwe-asd

У пользователя должны быть права на обновление конфигурации,тестировались только латинские имена, но проблем с кириллицей быть не должно.
При запуске сперва выводится список баз с версиями и типами конфигураций, потом предлагается подтвердить дальнейшее выполнение для возможности проверки наличия файлов в каталогах обновлений. Скрипт проверяет, чтобы версия обновления была старше версии текущей конфигурации и чтобы обновление было корректным для текущей версии, для этого используется файл UpdInfo.txt. При отказе от продолжения без обновления через скрипт на сервере 1С остаются COM-соединения к базе, которые надо удалять вручную через консоль администрирования серверов 1С. Функционала сброса COM-соединений через скрипт нет, они сбрасываются платформой при обновлений, для подтверждения используются функции window2 и window3.
