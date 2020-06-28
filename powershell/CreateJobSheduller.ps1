# Данный скрипт создает задание в планировщике Windows
# Задача будет лежать в разделе: Microsoft\Windows\PowerShell\SheduledJobs.

$JOB_NAME = "Разервное копирование папки"
$SCRIPT_PATH = "D:\PowerShell\BackupAdmin.ps1"

# В данном случае создается задание которое выполняется ежедневно в 4 вечера
$TRIGGER = New-JobTrigger -Daily -At 4:00PM

# Учетные данные для выполняния задания
$CREED = Get-Credential test\test

# Как опцию можно укзать запуск скрипта в повышенными привелегиями
$OPTION = New-ScheduledJobOption -RunElevated

# Решистрация задания в планировщике 
Register-ScheduledJob -Name $JOB_NAME -FilePath $SCRIPT_PATH -Trigger $TRIGGER -Credential $CREED -ScheduledJobOption $OPTION