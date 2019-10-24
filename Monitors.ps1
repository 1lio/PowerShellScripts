# Информация о мониторах которую смог выдрать из параметров монитора
# Задумывался как часть PCREPORT.ps1
function Get-MonitorDetails
{
  param
  (
    [Object]
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Data to process")]
    $InputObject
  )
  process
  {
    $Manufacturer = ($InputObject.ManufacturerName -notmatch 0 | ForEach-Object{[char]$_}) -join ""
    $MName = ($InputObject.UserFriendlyName -notmatch 0 | ForEach-Object{[char]$_}) -join ""
    $MSerial = ($InputObject.SerialNumberID -notmatch 0 | ForEach-Object{[char]$_}) -join ""

    return [pscustomobject]@{
      Manufacturer = $Manufacturer
      MName        = $MName
      MSerial      = $MSerial
    }
  }
}

Get-WmiObject WmiMonitorID -Namespace root\wmi | Get-MonitorDetails