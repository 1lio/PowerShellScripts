REM Скрипт удаления файлов в директории

echo off
setlocal enableextensions enabledelayedexpansion

REM Создаем переменную содержащий путь папки
set sTargetFolder=D:\Folder

REM Уникальные поддиректории которые чистить не нужно
set sExcludeFolders="1" "2"

del /Q /F %sTargetFolder%\* 

REM очищаем директории/поддиректории
for /f "tokens=*" %%i in ('dir /AD /B %sTargetFolder%') do (
	set /a bDelete = 1
	
	for %%j in (%sExcludeFolders%) do (
		if /i "%%i" equ "%%~j" set /a bDelete = 0
	)
	
	if !bDelete! equ 1  (
		del /S /Q /F "%sTargetFolder%\%%i\*" 
		RD /S /Q "%sTargetFolder%\%%i"
		MD "%sTargetFolder%\%%i"
	)
)

endlocal
exit /b 0