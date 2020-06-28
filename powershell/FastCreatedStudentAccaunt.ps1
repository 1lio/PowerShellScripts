# Скрипт для автоматизации создания учетных записей
# Необходим в случаях когда необхоимо создать большое колличество учетных записей

#Для работы скрипта необходим модуль AD
Import-Module ActiveDirectory

# Данный скрипт создает 50 учетных записей студент

$domain = "mydomain"
$domaintRoot = "com"
$department = "Students"

# Параметры учетных записей
$username="student"
$password = "@$$w0rd"
$countUsers = 1..50

# Включена ли учетная запись
$isBlockedAccount = $true

# Запрашивать смену пароля при первом подключении
$isUpdatedPass = $flase

$org = "OU=$department, DC=$domain, DC=$domaintRoot"

foreach ($i in $countUsers){
   { New-AdUser -Name $username$i -Path $org -Enabled $isBlockedAccount -ChangePasswordAtLogon $isUpdatedPass  `
    -AccountPassword (ConvertTo-SecureString «$password» -AsPlainText -force) -passThru 
   }
}

# Чтобы дать учетным данны одинаковые права необходимо настроить одну учетную запись, а затем использовать ее в качестве шаблона
# с помощью параметра -Instance Например:

<# $template = Get-AdUser -Identity ″student″
  foreach ($i in $countUsers){
   { New-AdUser -Name $username$i -UserPrincipalName $username$i -Path $org -Instance `
    $template -Enabled $True -ChangePasswordAtLogon $true `
    -AccountPassword (ConvertTo-SecureString «$password» -AsPlainText -force) -passThru
   }
} #>
