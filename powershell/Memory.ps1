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