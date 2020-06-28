<# Скрипт выводит компьютер из текущего домена и вводит обратно
   @autor Sukhov Viachelav 2019 #>

# Директория с временным файлом
$TEMP_PATH = "C:\temp\"
$TEMP_PATH_FILE = "C:\temp\ReConnectToDomain.temp"

# Запрос прав администратора
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host "Скрипт запущен без прав администратора!" -ForegroundColor Red
    Write-Host "Запрос прав..." -ForegroundColor Red

    #запуск нового экземпляра данного скрипта, с запросом прав администратора
    Start-Process powershell $PSCommandPath -verb runAs
    Break
}

# Установка разрешение выполнения PowerShell скриптов
Try {
    Set-ExecutionPolicy Unrestricted
    Write-Host "Запрос разрешения на выполение скрипта. Успех!" -ForegroundColor Green
}
Catch {
    Write-Host "Запрос разрешения на выполение скрипта. Неудача!" -ForegroundColor Red
    Write-Host "Выполните команду: Set-ExecutionPolicy Unrestricted" -ForegroundColor Cyan
    powershell.exe -NoExit | Out-Null
}

# Проверка были ли созданы временные файлы
$RESUME_SCRIPT = Test-Path $PSCommandPath

if($RESUME_SCRIPT -eq $false) {

    # Создаем временный файл с данными о текущем домене

     Try{     
              
        # Создание директории и временного файла
        New-Item -Path  $TEMP_PATH -ItemType Directory -Force | Out-Null
        New-Item -Path  $TEMP_PATH_FILE -ItemType File -Force | Out-Null
       
        # Запись данных о домене и администраторе
        $USERDOMAIN = (Get-WmiObject win32_computersystem).Domain

        "Domain: $USERDOMAIN" > $TEMP_PATH_FILE | Out-Null
        Write-Host "Создание директории. Успех!"-ForegroundColor Green
        } Catch{
    
         Write-Host "Ошибка! Нет доступа к деритории $TEMP_DIRECTORY" -ForegroundColor   Cyan
         powershell.exe -NoExit | Out-Null
       }

    #Продолжение выполненения скрипта после перезагрузки
    Write-Host "Создаю задачу на продолжение скрипта после перезапуска" -ForegroundColor   Cyan

    Set-Location HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce
    
    
    # Remove-ItemProperty –Path HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce –Name "NextStepConnectToDomain"
    New-ItemProperty . NextStepConnectToDomain -propertytype String -value "powershell $PSCommandPath"
       
    # Процесс вывода из домена  
    Write-Host "Запуск процесса вывода из домена"-ForegroundColor Green      
    Remove-Computer -UnjoinDomainCredential $env:USERDOMAIN -PassThru -Verbose -Restart    
    powershell.exe -NoExit | Out-Null

} else {
    # Присоединение к домену

    Write-Host "Инициирую обратное подключение к домену " -ForegroundColor   Green
    $next = Read-Host "Продолжить? [Y]/[N]"

    if ($next -eq 'y') {
    Write-Host "Запускаю процедуру присоединения к домену" -ForegroundColor Cyan

     # Читаем файл как одну строку
    $file = [System.IO.File]::ReadAllText($TEMP_PATH_FILE)

    # Индекс начала
    $start = $file.IndexOf("Domain: ")
    
    # Читаем выделенный текст
    $name = $file.Substring($start)
    $name = $name.Trim("Domain: ") 
    $name = $name.TrimStart()
    $name = $name.TrimEnd()
    
    Write-Host "Ожидаю ввода пароля..." -ForegroundColor Cyan
    $creed = $Host.UI.PromptForCredential("Логин", "Введите данные администратора", "", "NetBiosUserName")

    Write-Host "Добавление в домен $name..." -ForegroundColor Cyan
    # Ввод в домен без переименования компьютера
      
    Remove-Item -Path $TEMP_PATH
    Add-Computer  -ComputerName $env:COMPUTERNAME -DomainName $name.ToUpper() -Credential $creed -Restart
    
    powershell.exe -NoExit | Out-Null
    } 
 
    else {
     # завершение скрипта
     exit
    }
}