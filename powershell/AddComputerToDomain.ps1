<# Скрип ввода компьютера в домен.
   @autor Sukhov Viachelav 2019 #>

# Запрос прав админа
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host "Скрипт запущен без прав администратора!" -ForegroundColor Red
    Write-Host "Запрос прав..." -ForegroundColor Red

    #запуск нового экземпляра данного скрипта, с запросом прав администратора
    Start-Process powershell $PSCommandPath -verb runAs
    Break
}

try {
    # установка разрешение выполнения PowerShell скриптов
    Set-ExecutionPolicy Unrestricted
    Write-Host "Запрос разрешения на выполение скрипта. Успех!" -ForegroundColor Green
}
catch {

    Write-Host "Запрос разрешения на выполение скрипта. Неудача!" -ForegroundColor Red
    Write-Host "Выполните команду: Set-ExecutionPolicy Unrestricted" -ForegroundColor Cyan
    powershell.exe -NoExit | Out-Null
}

Write-Host "Данный скрипт вводит компьютер в домен!" -ForegroundColor Cyan
$next = Read-Host "Продолжить? [Y]/[N]"

if ($next -eq 'y') {
    Write-Host "Запускаю процедуру присоединения к домену" -ForegroundColor Cyan

    $nameComuter = Read-Host "Введите имя компьютера"
    $domainName = Read-Host "Введите имя домена"

    Write-Host "Идет подключение к домену..." -ForegroundColor Cyan
    $creed = $Host.UI.PromptForCredential("Логин", "Введите данные администратора", "", "NetBiosUserName")

    if ($env:COMPUTERNAME -eq $nameComuter) {
        # Ввод в домен без переименования компьютера
        Add-Computer  -ComputerName $env:COMPUTERNAME -DomainName $domainName.ToUpper() -Credential $creed -Restart
    }
    else {
        # Ввод в домен с новым именем
        Add-Computer  -ComputerName $env:COMPUTERNAME -DomainName $domainName.ToUpper()  -NewName $nameComuter -Credential $creed -Restart
    }

    powershell.exe -NoExit | Out-Null
}
else {
    # завершение скрипта
    exit
}