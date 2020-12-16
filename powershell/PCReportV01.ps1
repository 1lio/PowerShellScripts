# Опросить компы из списка
import-csv c:\Invent\AllComputers.csv | foreach {
$a=$_.name

if ((Test-connection $a -count 2 -quiet) -eq "True")
{
Write-Host "Проверка компьютера " -ForeGroundColor Green $a
"Компьютер" | out-file c:\Invent\Comp\$a.txt
Get-WmiObject -computername $a Win32_OperatingSystem | 
select-object csname, caption, Serialnumber, csdVersion | 
ft @{Label="Сетевое имя"; Expression={$_.CSname}},
@{label="Наименование"; Expression={$_.caption}},
@{label="Версия"; Expression={$_.csdVersion}},
@{label="Серийный номер"; Expression={$_.SerialNumber}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_ComputerSystemProduct | select-object UUID | 
ft UUID -autosize | out-file c:\Invent\Comp\$a.txt -append
"Процессор" | out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_Processor | select-object name, SocketDesignation, Description | 
ft @{label="Имя"; Expression={$_.name}},
@{label="Разъем"; Expression={$_.SocketDesignation}},
@{label="Описание"; Expression={$_.Description}} -auto -wrap | out-file c:\Invent\Comp\$a.txt -append
"Материнская плата" | out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_BaseBoard | select-object Manufacturer, Product, SerialNumber | 
ft @{label="Производитель"; Expression={$_.manufacturer}},
@{label="Модель"; Expression={$_.Product}},
@{label="Серийный номер"; Expression={$_.SerialNumber}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
"Жесткие диски" | out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_DiskDrive | select-object Model, Partitions, Size, interfacetype | 
ft @{Label="Модель"; Expression={$_.Model}}, 
@{Label="Количество разделов"; Expression={$_.Partitions}}, 
@{Label="Размер (гб)"; Expression={($_.Size/1GB).tostring("F00")}}, 
@{Label="Интерфейс"; Expression={$_.interfaceType}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
"Логические диски" | out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_LogicalDisk -Filter "DriveType=3" | select-object DeviceID, FileSystem, Size, FreeSpace | 
ft @{Label="Наименование"; Expression={$_.DeviceID}},
@{Label="Файловая система"; Expression={$_.FileSystem}}, 
@{Label="Размер (гб)"; Expression={($_.Size/1GB).tostring("F00")}},
@{Label="Свободное место (гб)"; Expression={($_.FreeSpace/1GB).tostring("F00")}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
"Оперативная память" | out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_Physicalmemory | Select-Object capacity, DeviceLocator | 
ft @{Label="Размер (мб)"; Expression={($_.capacity/1MB).tostring("F00")}},
@{Label="Расположение"; Expression={$_.DeviceLocator}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
"Видеокарта" | out-file c:\Invent\Comp\$a.txt -append
Get-WmiObject -computername $a Win32_videoController | 
Select-Object name, AdapterRAM, VideoProcessor | 
ft @{Label="Наименование"; Expression={$_.name}},
@{Label="Объем памяти (мб)"; Expression={($_.AdapterRAM/1MB).tostring("F00")}},
@{Label="Видеопроцессор"; Expression={$_.VideoProcessor}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
"Сетевая карта" | out-file c:\Invent\Comp\$a.txt -append
$OS=Get-WmiObject -computername $a Win32_OperatingSystem | foreach {$_.caption} 
if ($OS -eq "Microsoft Windows 2000 Professional") 
{
Get-WmiObject -computername $a Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" | 
Select-Object caption,MACaddress |
ft @{Label="Наименование"; Expression={$_.caption}},
@{Label="MAC адрес"; Expression={$_.MACAddress}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
}
else 
{
Get-WmiObject -computername $a Win32_NetworkAdapter -Filter "NetConnectionStatus>0" | 
Select-Object name, AdapterType, MACAddress | 
ft @{Label="Наименование"; Expression={$_.name}},
@{Label="MAC адрес"; Expression={$_.MACAddress}},
@{Label="Тип"; Expression={$_.AdapterType}} -auto -wrap | 
out-file c:\Invent\Comp\$a.txt -append
}
}
}