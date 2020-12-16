REM Данный скрипт задает агенту адрес сервера администрирования Kaspersky
echo off

REM Путь по умолчанию где лежит агент
cd C:\Program Files (x86)\Kaspersky Lab\NetworkAgent

REM запускаем утилиту https://support.kaspersky.ru/12336  передаем адрес в качестве параметра
klmover.exe -address kas