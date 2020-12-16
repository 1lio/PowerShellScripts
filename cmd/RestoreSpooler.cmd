REM Перезапуск службы печати / очистка очереди

echo off

net stop spooler
del %systemroot%\system32\spool\printers\*.shd /F /S /Q
del %systemroot%\system32\spool\printers\*.spl /F /S /Q

net start spooler
