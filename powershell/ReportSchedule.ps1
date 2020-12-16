# Задача на инвентаризацию

# Для работы с Ad Необходим импортнуть модуль или в ролях сервира добавить компонент WMI
import-module activedirectory

# Получить список всех компьютеров в домене и экспортировать в файл
get-ADcomputer -Filter * | 
Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,DC=dc"} |
Sort-Object name | Select-Object name | Export-csv C:\Invent\AllComputers.csv -NoTypeInformation

# Опросить компы из списка
import-csv c:\Invent\AllComputers.csv | foreach {
$a=$_.name

    # Проверяем что компьютер доступен
    if ((Test-connection $a -count 2 -quiet) -eq "True")
    {

    # Создание задачи в планировщике на запуск скрипта отчета о окмпьютере
    # Время запуска
    $Trigger= New-ScheduledTaskTrigger -At 10:00am -Daily
    
    # Запуск от системы
    $User= "NT AUTHORITY\SYSTEM"
    # Запустить скрипт лежащего по сетевому пути
    $Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "\\Share\scripts\PCReport.ps1"
    Register-ScheduledTask -TaskName "StartupScript_PS" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force
    }
}