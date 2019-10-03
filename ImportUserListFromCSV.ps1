# Данный скрипт предназначен для импорта пользователей из CSV файла
# В качестве столбцов необходимо указать атрибуты пользователя

#Для работы скрипта необходим модуль AD
Import-Module ActiveDirectory

$defPass = "P@$$w0rd"
$csvPath = Import-CSV -Path ″C:\scripts\users.csv″
$csvPath | New-AdUser  -Path $org -Enabled $True -ChangePasswordAtLogon $true `
-AccountPassword (ConvertTo-SecureString $defPass -AsPlainText -force) -passThru