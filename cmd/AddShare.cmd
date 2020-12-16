REM Добавление шары (сетевой диск)
echo off
net use t: /delete 
net use t: \\fs\temp$

