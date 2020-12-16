REM Поиск ip адреса по MAC
echo off
if "%1" == "" echo no MAC address & exit /b 1

REM Ищем в определенном диапазоне адресов (простым перебором)
for /L %%a in (1,1,254) do @start /b ping 192.168.0.%%a -n 2 > nul
ping 127.0.0.1 -n 3 > nul
arp -a | find /i "%1"