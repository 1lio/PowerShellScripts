# Данный скрипт формирует exel-отчет о компьютере или группе компьютеров
##################################################################################################

$PC_NAME = $env:COMPUTERNAME
$DOMAIN_NAME = $env:USERDOMAIN
$EXCEL_LIST = "Отчет"

##################################################################################################
# Функция конвертации RGB в OLE
##################################################################################################

Function ConvertToOLE($R, $G, $B){ 
    Return $R + ($G * 256) + ($B * 256 * 256) 
}

##################################################################################################

# Константные значения для цветов

$BLUE =  ConvertToOLE 83 141 213
$ORANGE =  ConvertToOLE 247 150 70
$GREEN = ConvertToOLE 146 206 80
$RED = ConvertToOLE 192 80 77
$GRAY =  ConvertToOLE 89 89 89
$BLACK = 0
$WHITE = ConvertToOLE 255 255 255

# Выравнивания
$CENTER = -4108

##################################################################################################
# Формирование Excel документа
##################################################################################################

# Создаем объект Exel и делаем его видимым | Равносильно запуску exel
$Excel = New-Object -ComObject Excel.Application

# Можно запустить с параметром false, чтобы пользователь не видел, если нужно скрыть выполнение
$Excel.Visible = $true

# Добавляем книгу/лист
$WorkBook = $Excel.Workbooks.Add()

# Фокусируемся на 1 листе и даем имя листа ("Отчет")
$Report = $WorkBook.Worksheets.Item(1)
$Report.Name = $EXCEL_LIST

##################################################################################################

# Вызов функций

$TD = 1 # Столбец
$TR = 1 # Строка

CreateHeader           # шапка отчета
CreateReportBase       # отчет о комплектующих на материнской плате
CreateReportVideo      # отчет о установленных дисплеях и видео адаптере
CreateReportData       # отчет о физических накопителях 
CreateReportLogicData  # отчет о локальных дисках 
CreateReportEthernet   # отчет о об сетевых устройствах
DicorateReport         # декорирование

##################################################################################################

# Шапка отчета
Function CreateHeader() {
    
    # Дата формирования отчета
    $DateCreate = Get-Date -Format d
    $Report.Cells.Item(1, 1) = "Отчет от $DateCreate"

    # Текст компьютер
    $Report.Cells.Item(1,2) = "Компьютер:"

    # Имя компьютера
    $Report.Cells.Item(1,3) = "$PC_NAME | $DOMAIN_NAME"

    # Пользователь
    $user = $env:USERNAME
    $user = $user.ToUpper()
    $Report.Cells.Item(3, 1) = "Пользователь:`n$user"
    
    # OS/ServicePack
    $os = (Get-WMIObject win32_operatingsystem).caption
    $sp = (Get-WMIObject win32_operatingsystem).csdVersion

    $Report.Cells.Item(3,2) = $os
    $Report.Cells.Item(3,3) = $sp
    
    # Office
    $Report.Cells.Item(4,2) = GetOffice

    # Тип лицензии
    $Report.Cells.Item(4,3) = GetOfficeType

    #Antivirus
    $AntivirusProduct = Get-WmiObject -Namespace "root\SecurityCenter2" -Query "SELECT * FROM AntiVirusProduct" 
    $Antivirus = $AntivirusProduct.displayName
    $Report.Cells.Item(5,2) = "$Antivirus"
}

#Office
Function GetOffice(){

   # Список программ в реестре
   $programms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName 
   $office = $programms -like "*Microsoft Office*"
   $office = $office[0]
   
   # Выйти из функции если не обнаружено продуктов Microsoft Office
   if($office -eq $null) { Return "Продукт Microsoft Office не установлен"}

   # Обрезать системную информацию
   $office = $office.TrimStart("@{DisplayName=")
   $office = $office.TrimEnd("}")
       
   Return "$office"
}

# Тип активации MS Office (KMS, MAC, Retail)
Function GetOfficeType() {
   $office = GetOffice

   # Получаем номер office
   $office = $office.Split("Microsoft Office")[-1]   
   $officeVer = ""

   # Нумерация офисов 11-2003, 12-2007, 14-2010, 15-2013, 16-2016, 19-2019
   Switch ($office) {
    2003 {$officeVer = "Office11"}
    2007 {$officeVer = "Office12"}
    2010 {$officeVer = "Office14"}
    2013 {$officeVer = "Office15"}
    2016 {$officeVer = "Office16"}
    2019 {$officeVer = "Office19"}
   }
   
   # Встроенной утилитой ospp олучаем тип активации
   Set-Location "C:\Program Files\Microsoft Office\$officeVer"
   $a = cscript ospp.vbs /dstatus 
   $ar = $a -match "^LICENSE NAME"
   $ar = $ar[0]
       
   # Проверка на тип активации

   $isMAK = $ar.Contains("MAK")
   $isKMS = $ar.Contains("KMS")
   $isRetail = $ar.Contains("Retail")

   $OfficeVersion = ""
   
   if($isMAK -eq $true) {
     $OfficeVersion = "MAK"
   }
   
   if($isKMS -eq $true) {
     $OfficeVersion = "KMS"
   }

   if($isRetail -eq $true) {
     $OfficeVersion = "Retail"
   }

   Return $OfficeVersion
}

##################################################################################################
# Отчет о комплектующих на материнской плате
Function CreateReportBase(){
  
   # Заголовок
   $Report.Cells.Item(7,1) = "Системная плата"

   # ЦП
   $cpu = (Get-WMIObject Win32_Processor).name
   $Report.Cells.Item(9,1) = "ЦП"
   $Report.Cells.Item(9,2) = $cpu

   # Материнская плата
   $manufacturer = (Get-WMIObject Win32_BaseBoard).manufacturer 
   $product = (Get-WMIObject Win32_BaseBoard).Product 
   $socket =  (Get-Ciminstance win32_processor).socketdesignation

   $Report.Cells.Item(10,1) = "Системная плата"
   $Report.Cells.Item(10,2) = "$manufacturer $product"

   # Чипсет
   $chipset = (Get-WMIObject Win32_BaseBoard).product  
   $Report.Cells.Item(11,1) = "Чипсет"
   $Report.Cells.Item(11,2) = "Нет данных"

    # ОЗУ
    $memory  =(Get-WMIObject Win32_Physicalmemory)
    $mem = (Get-WMIObject win32_PhysicalMemoryArray)

    # Нужно узнать сколько планок
    $count = $memory.capacity.Length #тут узнаем что у меня 2 планки

    # Найти общий размер ОЗУ  
    $sum = 0
    for($x = 0; $x -le $count; $x++) {
     $sum += $memory.capacity[$x]
    }
    
    $maxSize = $mem.MaxCapacity
    # Общий объем в МБ
    $max = $maxSize / (1024*1024)
    $result = $sum / (1024*1024)

    $manufacter = $memory.Manufacturer
   
    $part = $memory.PartNumber  
    $speed = $memory.Speed  

    if($part[0] -eq $part[1]){
     $part = $part[0]
     $manufacter = $manufacter[0] + "x$count"
     $speed = $speed[0]
    }                    

    $Report.Cells.Item(12,1) = "Оперативная память"
    $Report.Cells.Item(12,2) = "$manufacter | $part | $result Мб | $speed Мгц | Мах $max Гб"
    
    # BIOS
    $bios = (Get-WMIObject Win32_BIOS).BiosVersion
    $Report.Cells.Item(13,1) = "BIOS"
    $Report.Cells.Item(13,2) = $bios
}


##################################################################################################
# Отчет о видеоадаптерах и мониторах
Function CreateReportVideo(){

    # Шапка Дисплей
    $Report.Cells.Item(15,1) = "Дисплей"

    # Видеоадаптер
    $videoadapter = (Get-WMIObject win32_VideoController).name
    $Report.Cells.Item(17,1) = "Видеоадаптер"
    $Report.Cells.Item(17,2) = $videoadapter

    # Мониторы
    $Monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi
    
    # Общее колличество мониторов
    $CountMonitors =  $Monitors.Length 
    
    $Report.Cells.Item(18,1) = "Общее колличество мониторов"
    $Report.Cells.Item(18,2) = $CountMonitors

    $i = 1;
    ForEach ($Monitor in $Monitors) {
       
        $Manufacturer = ($Monitor.ManufacturerName -notmatch 0 | ForEach{[char]$_}) -join ""
        $MSerial = ($Monitor.SerialNumberID -notmatch 0 | ForEach{[char]$_}) -join ""

         # Расшифровка популярных производителей 
         Switch ($Manufacturer) {
          "GSM" {$Manufacturer = "LG" }
          "ACR" {$Manufacturer = "Acer" }
          "SAM" {$Manufacturer = "Samsung" }
         }
     
    $Report.Cells.Item(18 + $i, 1) = "$Manufacturer"
    $Report.Cells.Item(18 + $i, 2) = "$MSerial"
    $i++
   }   
}

##################################################################################################
# Отчет хранилищах
Function CreateReportData(){

    #Шапка Хранение данных
    $Report.Cells.Item(22,1) = "Хранение данных"

    # Оптический накопитель
    $cdAdapter = Get-WmiObject Win32_CDROMDrive
    $cdAdapter = $cdAdapter.Caption 
  
    $Report.Cells.Item(24,1) = "Оптический накопитель"
    $Report.Cells.Item(24,2) = $cdAdapter
    
    # Дисковые накопитеи
    $disks = Get-WmiObject Win32_DiskDrive 
    
    $hardDiskCount = $Monitors.Length 

     
    $Report.Cells.Item(25,1) = "Колличество накопителей"
    $Report.Cells.Item(25,2) = $hardDiskCount

    $i = 1

    ForEach ($disk in $disks) {

     $Manufacturer = $disk.Caption 
     $Size = $disk.size  
     # Конвертация в ГБ
     $Size = $Size / 1073741824

     # Округление до 1 после запятой
     $Size = [Math]::Round($Size, 1)
    
     # Производитель диска   
     $Report.Cells.Item(25+$i,1) = $Manufacturer
     # Размер диска
     $Report.Cells.Item(25+$i,2) = $Size
     $i++
   }   
}


##################################################################################################
# Отчет логических дисках
Function CreateReportLogicData(){

    #Шапка Разделы
    $Report.Cells.Item(29,1) = "Логические разделы"

    # Дисковые накопитеи
    $disks = Get-WmiObject Win32_LogicalDisk 
         
    $i = 1

    ForEach ($disk in $disks) {

     $Name = $disk.Caption 
     $Size = $disk.Size 
     $Free = $disk.FreeSpace 
     
      
     # Конвертация в ГБ
     $Size = $Size / 1073741824
     $Free = $Free / 1073741824

     # Округление до 1 после запятой
     $Size = [Math]::Round($Size, 1)
     $Free = [Math]::Round($Free, 1)
        
     # Производитель диска   
     $Report.Cells.Item(30+$i,1) = $Name
     # Размер диска
     $Report.Cells.Item(30+$i,2) = "Свободно $Free из $Size"
     $i++
   }   

}

##################################################################################################
# IP/Mac адрес
Function CreateReportEthernet() {

    # IP-адреса адрес активной сетевой карты
    $ips = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True'
      
    $i = 1
    ForEach ($ip in $ips) {

    $Name = $ip.IPAddress 

    $Report.Cells.Item(3+$i,5) = $Name
    $i++    
    }
    
    # MAC адрес активной сетевой карты
    $mac = Get-WmiObject Win32_NetworkAdapter -Filter 'NetConnectionStatus=2'
    $macAdress = $mac.MACAddress      
    $Report.Cells.Item(4+$i,5) = $macAdress
}

##################################################################################################
# Дикорирование отчета

Function DicorateReport() {

# Форматирование
$Report.Cells.Item(1, 1).Interior.Color = $BLUE
$Report.Cells.Item(1, 2).Interior.Color = $ORANGE
$Report.Cells.Item(1, 3).Interior.Color = $GRAY

$Report.Cells.Item(3, 1).Characters(14,$DateCreate).Font.Color = $BLUE

$Report.Cells.Item(4, 5).Interior.Color = $GRAY
$Report.Cells.Item(5, 5).Interior.Color = $GRAY
$Report.Cells.Item(7, 5).Interior.Color = $ORANGE

$Range = $Report.Range('E4','E7')
$Range.Font.Size = 14
$Range.Font.Bold = $true
$Range.Font.Color = $WHITE
$Range.VerticalAlignment = $CENTER
$Range.HorizontalAlignment = $CENTER

$Range = $Report.Range('A1','C1')
$Range.Font.Size = 14
$Range.Font.Bold = $true
$Range.Font.Color = $WHITE
$Range.VerticalAlignment = $CENTER
$Range.HorizontalAlignment = $CENTER

$Range = $Report.Range('A3','A5')
$Range.Font.Size = 14
$Range.Font.Bold = $true
$Range.VerticalAlignment = $CENTER
$Range.HorizontalAlignment = $CENTER
$Range.Merge()

# Подгоняем размеры
$Range = $Report.Range('A1','C3')
$Range = $Report.UsedRange
$Range.EntireColumn.AutoFit() | Out-Null

$Range = $Report.Range('A1','A3')
$Range = $Report.UsedRange
$Range.EntireRow.AutoFit() | Out-Null

$Report.Cells.Item(7,1).Interior.Color = $BLUE
$Report.Cells.Item(15,1).Interior.Color = $BLUE
$Report.Cells.Item(22,1).Interior.Color = $BLUE
$Report.Cells.Item(29,1).Interior.Color = $BLUE


$Report.Cells.Item(7,1).Font.Color = $WHITE
$Report.Cells.Item(15,1).Font.Color = $WHITE
$Report.Cells.Item(22,1).Font.Color = $WHITE
$Report.Cells.Item(29,1).Font.Color = $WHITE

$Report.Cells.Item(7,1).Font.Bold = $true
$Report.Cells.Item(15,1).Font.Bold = $true
$Report.Cells.Item(22,1).Font.Bold = $true
$Report.Cells.Item(29,1).Font.Bold = $true

$Report.Cells.Item(7,1).VerticalAlignment = $CENTER
$Report.Cells.Item(15,1).VerticalAlignment = $CENTER
$Report.Cells.Item(22,1).VerticalAlignment = $CENTER
$Report.Cells.Item(29,1).VerticalAlignment = $CENTER


$Report.Cells.Item(7,1).HorizontalAlignment = $CENTER
$Report.Cells.Item(15,1).HorizontalAlignment = $CENTER
$Report.Cells.Item(22,1).HorizontalAlignment = $CENTER
$Report.Cells.Item(29,1).HorizontalAlignment = $CENTER

$Range = $Report.Range('A7','C7')
$Range.Merge()

$Range = $Report.Range('A15','C15')
$Range.Merge()

$Range = $Report.Range('A22','C22')
$Range.Merge()

$Range = $Report.Range('A29','C29')
$Range.Merge()

}

# Сохраняем в файл
$WorkBook.SaveAs("D:\reporst_$PC.xls"); 
$WorkBook.Close()
$WorkBook.Quit()