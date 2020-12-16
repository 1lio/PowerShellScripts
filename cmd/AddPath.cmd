REM Создать системную/пользовательскую переменную
echo off

REM Глобальная переменная. в PATH будет добавлен D:\Path
setx /M path "%PATH%;D:\Path"

REM Пользовательская переменная
setx MyPath "Path Value"
