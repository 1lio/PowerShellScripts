echo off

REM бежим по сетевухам и собираем адреса (ipv4)
for /f "tokens=2" %%a in ('netsh interface ipv4 show addresses^|find "IP-"') do set LocIP=%%a& Goto extNetsh
:extNetsh

REM выводим в консоль результат
echo %LocIP%
pause
