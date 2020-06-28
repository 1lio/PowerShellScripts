$Monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi
Write-Host $Monitors.Length  -ForegroundColor Cyan

ForEach ($Monitor in $Monitors) {

    $Manufacturer = ($Monitor.ManufacturerName -notmatch 0 | ForEach{[char]$_}) -join ""
    $MSerial = ($Monitor.SerialNumberID -notmatch 0 | ForEach{[char]$_}) -join ""

   Switch ($Manufacturer) {
    "GSM" {$Manufacturer = "LG"}
    "ACR" {$Manufacturer = "Acer"}
    "SAM" {$Manufacturer = "Samsung"}
   }
   
   Write-Host "$Manufacturer Ser: $MSerial" -ForegroundColor Yellow
}   