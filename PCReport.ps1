# Данный скрипт формирует exel-отчет о компьютере или группе компьютеров
# TODO: Доделать!
$pc = $env:COMPUTERNAME
$domain = $env:USERDOMAIN

# Цвета OLE
$BLUE = 11307632
$ORANGE = 3386879
$GREEN = 65280
$RED = 255
$GRAY = 6710886
$BLACK = 0
$WHITE = 16777215

# Выравнивания
$CENTER = -4108
    

# Создаем объект Exel и делаем его видимым | Равносильно запуску exel
$Excel = New-Object -ComObject Excel.Application

# Можно запустить с параметром false чтобы пользователь не видел, если скрипт выполняется локально
$Excel.Visible = $true

# Добавляем книгу/лист
$WorkBook = $Excel.Workbooks.Add()

# Фокусируемся на 1 листе и даем имя "Отчет"
$Report = $WorkBook.Worksheets.Item(1)
$Report.Name = "Отчет"

CreateDateD 1 1

# Первая ячейка Дата отчета
function CreateDateD($Row, $Column) {

$DateCreate = Get-Date -Format d
$Report.Cells.Item($Row, $Column) = "Отчет от $DateCreate"

}

# Просто тикст компьютер
$Report.Cells.Item(1,2) = "Компьютер:"


# Имя компьютера
$pc = (Get-WMIObject win32_operatingsystem).CSname
#TODO: Править
$domain = $env:USERDOMAIN
$Report.Cells.Item(1,3) = "$pc | $domain"

# Пользователь
$user = $env:USERNAME
$user.ToUpper()

$Report.Cells.Item(3,1) = "Пользователь:`n$user"

#OS
$os = (Get-WMIObject win32_operatingsystem).caption
$sp = (Get-WMIObject win32_operatingsystem).csdVersion
$Report.Cells.Item(3,2) = $os
$Report.Cells.Item(3,3) = $sp

#Office

#Antivirus

#Шапка Системная плата
$Report.Cells.Item(7,1) = "Системная плата"

# ЦП
$cpu = (Get-WMIObject Win32_Processor).name

$Report.Cells.Item(9,1) = "ЦП"
$Report.Cells.Item(9,2) = $cpu

# Мать
$manufacturer = (Get-WMIObject Win32_BaseBoard).manufacturer 
$product = (Get-WMIObject Win32_BaseBoard).Product 


$Report.Cells.Item(10,1) = "Системная плата"
$Report.Cells.Item(10,2) = "$manufacturer $product"

# Чипсет
$chipset = (Get-WMIObject Win32_BaseBoard).product  

$Report.Cells.Item(11,1) = "Чипсет"
$Report.Cells.Item(11,2) = "$chipset TODO: Найти возможности собирать информацию о чипсете"

# ОЗУ

AddMemory

function global:AddMemory() {
    $memory  =(Get-WMIObject Win32_Physicalmemory)
    $mem = (Get-WMIObject win32_PhysicalMemoryArray)

      # Нужно узнать сколько планок
    $count = $memory.capacity.Length #тут узнаем что у меня 2 планки
    $sum = 0

    # Найти сумму
    for($x = 0; $x -le $count; $x++) {
      $sum += $memory.capacity[$x]
    }

     # Перевести в МБ
   
    [UInt64]$gb = (1024*1024)


    $maxSize = $mem.MaxCapacity
   
    $max = $maxSize / $gb

    $result = $sum / 1048576

    $manufacter = $memory.Manufacturer
   
    $part = $memory.PartNumber  
    $speed = $memory.Speed                        

    $Report.Cells.Item(12,1) = "Оперативная память"
    $Report.Cells.Item(12,2) = "$manufacter | $part | $result Мб | $speed Мгц | Мах $max Гб"
}

# BIOS
$bios = (Get-WMIObject Win32_BIOS).BiosVersion

$Report.Cells.Item(13,1) = "BIOS"
$Report.Cells.Item(13,2) = $bios

#Шапка Дисплей
$Report.Cells.Item(15,1) = "Дисплей"

# Video
$videoadapter = (Get-WMIObject win32_VideoController).name

$Report.Cells.Item(17,1) = "Видеоадаптер"
$Report.Cells.Item(17,2) = $videoadapter

# Monitors
$monitor = (Get-WMIObject Win32_DesktopMonitor).name
$Report.Cells.Item(18,1) = "Монитор"
$Report.Cells.Item(18,2) = $monitor

#Шапка Хранение данных
$Report.Cells.Item(22,1) = "Хранение данных"

# Оптический накопитель
$cdAdapter = "TODO"
$Report.Cells.Item(24,1) = "Оптический накопитель"
$Report.Cells.Item(24,2) = $cdAdapter


# Дисковые накопитеи
$hardDisk = "TODO"
$Report.Cells.Item(25,1) = "Дисковый накопитель"
$Report.Cells.Item(25,2) = $hardDisk

#Шапка Разделы
$Report.Cells.Item(28,1) = "Логические разделы"

# Дисковые накопитеи
$diskSize = "TODO"
$diskWord = "TODO"
$Report.Cells.Item(30,1) = $diskWord
$Report.Cells.Item(30,2) = $diskSize

#Шапка Лицензии
$Report.Cells.Item(34,1) = "Лицензии"

$license = "TODO"
$licenseKey = "TODO"
$Report.Cells.Item(36,1) = $license
$Report.Cells.Item(36,2) = $licenseKey


# Красивости 
FormateHeader
FormateUsername
FormatedNamed

function global:FormateHeader() {

# Форматирование
$Report.Cells.Item(1, 1).Interior.Color = $BLUE
$Report.Cells.Item(1, 2).Interior.Color = $ORANGE
$Report.Cells.Item(1, 3).Interior.Color = $GRAY


$Range = $Report.Range('A1','C1')
$Range.Font.Size = 14
$Range.Font.Bold = $true
$Range.Font.Color = $WHITE
$Range.VerticalAlignment = $CENTER
$Range.HorizontalAlignment = $CENTER
}

function global:FormateUsername() {
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
}

function global:FormatedNamed() {


   $Report.Cells.Item(7,1).Interior.Color = $BLUE
   $Report.Cells.Item(15,1).Interior.Color = $BLUE
   $Report.Cells.Item(22,1).Interior.Color = $BLUE
   $Report.Cells.Item(28,1).Interior.Color = $BLUE
   $Report.Cells.Item(34,1).Interior.Color = $BLUE

   $Report.Cells.Item(7,1).Font.Color = $WHITE
   $Report.Cells.Item(15,1).Font.Color = $WHITE
   $Report.Cells.Item(22,1).Font.Color = $WHITE
   $Report.Cells.Item(28,1).Font.Color = $WHITE
   $Report.Cells.Item(34,1).Font.Color = $WHITE

   $Report.Cells.Item(7,1).Font.Bold = $true
   $Report.Cells.Item(15,1).Font.Bold = $true
   $Report.Cells.Item(22,1).Font.Bold = $true
   $Report.Cells.Item(28,1).Font.Bold = $true
   $Report.Cells.Item(34,1).Font.Bold = $true

   $Report.Cells.Item(7,1).VerticalAlignment = $CENTER
   $Report.Cells.Item(15,1).VerticalAlignment = $CENTER
   $Report.Cells.Item(22,1).VerticalAlignment = $CENTER
   $Report.Cells.Item(28,1).VerticalAlignment = $CENTER
   $Report.Cells.Item(34,1).VerticalAlignment = $CENTER

   $Report.Cells.Item(7,1).HorizontalAlignment = $CENTER
   $Report.Cells.Item(15,1).HorizontalAlignment = $CENTER
   $Report.Cells.Item(22,1).HorizontalAlignment = $CENTER
   $Report.Cells.Item(28,1).HorizontalAlignment = $CENTER
   $Report.Cells.Item(34,1).HorizontalAlignment = $CENTER


   $Range = $Report.Range('A7','C7')
   $Range.Merge()

   $Range = $Report.Range('A15','C15')
   $Range.Merge()

   $Range = $Report.Range('A22','C22')
   $Range.Merge()

   $Range = $Report.Range('A28','C28')
   $Range.Merge()
      
   $Range = $Report.Range('A34','C34')
   $Range.Merge()
 }


 # Сохраняем в файл