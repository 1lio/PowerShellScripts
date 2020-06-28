@echo off
for /f "tokens=2" %%a in ('netsh interface ipv4 show addresses^|find "IP-"') do set LocIP=%%a& Goto extNetsh
:extNetsh
echo %LocIP%
pause
