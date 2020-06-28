# Скрипт для создания ярылка программы в "Мой мопьютер", раздел "Прочее"

# Имя и описание программы
$PROGRAMM_NAME = "Notepad++" 
$PROGRAMM_DESCRIPTION = "Прикладное ПО"

# Путь до программы
$PROGRAMM_PATH = "C:\Program Files (x86)\Notepad++\notepad++.exe"

# Номер иконки (Иконка по умолчанию - 0)
$PROGRAMM_ICON = 0

##########################################################################################################

# Формат ключ вида {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
  $CHARS = "ABCDEF1234567890".ToCharArray()

# Можно оформить одним фором, но я был  пьян. ДР друга
  $PROGRAMM_HKEY = '{'
  1..8  | ForEach {  $PROGRAMM_HKEY += $CHARS | Get-Random }
  $PROGRAMM_HKEY += '-'
  1..4  | ForEach {  $PROGRAMM_HKEY += $CHARS | Get-Random }
  $PROGRAMM_HKEY += '-'
  1..4  | ForEach {  $PROGRAMM_HKEY += $CHARS | Get-Random }
  $PROGRAMM_HKEY += '-'
  1..4  | ForEach {  $PROGRAMM_HKEY += $CHARS | Get-Random }
  $PROGRAMM_HKEY += '-'
  1..12 | ForEach {  $PROGRAMM_HKEY += $CHARS | Get-Random }
  $PROGRAMM_HKEY += '}'

  Write-Host $PROGRAMM_HKEY -ForegroundColor 'Yellow'
  
  # Для монтируем путь HKCR
  New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
  # Переходи на путь глобальных индификаторов
  Set-Location "HKCR:\CLSID\"

  # Создаем раздел
  New-Item $PROGRAMM_HKEY -Type Directory
  Set-Location "HKCR:\CLSID\$PROGRAMM_HKEY\"

  # Правим значение по умолчанию даем имя программы
  Set-ItemProperty -Path . -Name '(default)' -Value $PROGRAMM_NAME
   
  # Создаем параметр для описания при наведениии InfoTip
  New-ItemProperty -Path . -Name "InfoTip" -Value "$PROGRAMM_DESCRIPTION"  -PropertyType "String"
   
  # Создаем каталог DefautIcon, Параметром по умолчанию указываем ссылку на ярлык и номер иконки
  New-Item "DefaultIcon" -Type Directory
  Set-ItemProperty -Path ".\DefaultIcon" -Name "(default)" -Value "$PROGRAMM_PATH,$PROGRAMM_ICON"

  # Создаем путь Shell\Open\Command  В качестве параметра указываем путь по умолчанию
  New-Item "Shell\Open\Command" -ItemType Catalog -Force
  Set-ItemProperty -Path ".\Shell\Open\Command" -Name "(default)" -Value "$PROGRAMM_PATH"
   
  # Включаем отображение в "Мой компьютер"
  Set-Location "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"
  New-Item $PROGRAMM_HKEY -Type Directory
  Set-ItemProperty -Path ".\$PROGRAMM_HKEY" -Name '(default)' -Value "$PROGRAMM_PATH"  