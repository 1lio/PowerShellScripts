# Запустить от имени Администратора
# Выполнить Set-ExecutionPolicy Unrestricted -Force

# Запрос прав админа
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host "Скрипт запущен без прав администратора!" -ForegroundColor Red
    Write-Host "Запрос прав..." -ForegroundColor Red

    #запуск нового экземпляра данного скрипта, с запросом прав администратора
    Start-Process powershell -verb runAs
    Break
}

[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")| Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
$browse.SelectedPath = "C:\"
$browse.ShowNewFolderButton = $false
$browse.Description = "Выберите папку"

$loop = $true
while($loop) {
    if ($browse.ShowDialog() -eq "OK") {
        $loop = $false
        cd $browse.SelectedPath
        Get-ChildItem -Recurse | Where-Object { $_.PSIsContainer -eq $false }| ForEach-Object { Rename-Item $_.FullName -newname ($_.name -replace "_CompressPdf" -replace "_compressed")
        }
    } else {
        return
    }
}

$browse.Dispose()
