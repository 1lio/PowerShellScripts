#Данный скрипт создает службу Windows и выполняет указанный скрипт
 
# Название службы
$ServiceName = "BackupAdminFolder"
 
# Путь по которому лежит-PS скрипт
$ScripPath = "D:\PowerShell\BackupAdmin.ps1"

$NSSMPath = (Get-Command "C:\tools\nssm\win64\nssm.exe").Source
$PoShPath = (Get-Command powershell).Source
$args = '-ExecutionPolicy Bypass -NoProfile -File "{0}"' -f $PoShScriptPath
& $NSSMPath install $NewServiceName $PoShPath $args
& $NSSMPath status $NewServiceName

#Запускаем службу
Start-Service $ServiceName

# Проверка что сужба запущена
Get-Service $ServiceName