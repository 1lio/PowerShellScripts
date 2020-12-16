# Список всех компьютеров

# Для работы с Ad Необходим импортнуть модуль или в ролях сервира добавить компонент WMI
import-module activedirectory

# Получить список всех компьютеров в домене и экспортировать в файл
# Сохранить как
$ExportFile = "C:\Invent\AllComputers.csv"

get-ADcomputer -Filter * | Where-Object {$a=$_.name; $_.DistinguishedName -ne "CN=$a,DC=dc03"} |
Sort-Object name | Select-Object name | Export-csv $ExportFile -NoTypeInformation