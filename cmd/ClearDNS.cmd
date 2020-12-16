REM Очистка DNS

echo off
ipconfig /release
arp -d *
nbtstat -R
REM Чистим кэш
ipconfig /flushdns
REM Запрашиваем новый адрес
ipconfig /renew

REM Перезапускаем службу
net stop dnscache
net start dnscache
