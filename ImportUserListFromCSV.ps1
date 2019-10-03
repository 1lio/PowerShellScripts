# Данный скрипт предназначен для импорта пользователей из CSV файла
# В качестве столбцов необходимо указать атрибуты пользователя
# Список все атрибутов можно посмотреть здесь https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/ee617253(v=technet.10)?redirectedfrom=MSDN

#Для работы скрипта необходим модуль AD
Import-Module ActiveDirectory

$defPass = "P@$$w0rd"
$csvPath = Import-CSV -Path ″C:\scripts\users.csv″
$csvPath | New-AdUser  -Path $org -Enabled $True -ChangePasswordAtLogon $true `
-AccountPassword (ConvertTo-SecureString $defPass -AsPlainText -force) -passThru