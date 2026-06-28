@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0.."

if not exist "data.7z" (
    echo [!] Файл data.7z не найден в корне проекта.
    exit /b 1
)

if not exist ".secret" (
    echo [!] Файл .secret не найден в корне проекта.
    exit /b 1
)

set "SEVEN_ZIP=C:\Program Files\7-Zip\7z.exe"
if not exist "%SEVEN_ZIP%" (
    echo [!] 7-Zip не найден: %SEVEN_ZIP%
    exit /b 1
)

set /p PASSWORD=<.secret

if not exist "Data" mkdir "Data"
if not exist "templates" mkdir "templates"

echo [+] Расшифровываю data.7z в Data\ и templates\.people.xlsx...
"%SEVEN_ZIP%" x -p"%PASSWORD%" -o. data.7z -y
if errorlevel 1 (
    echo [!] Ошибка 7-Zip. Неверный пароль?
    exit /b 1
)

if exist ".people.xlsx" (
    if not exist "templates\.people.xlsx" (
        move /y ".people.xlsx" "templates\.people.xlsx" >nul
    ) else (
        del /f ".people.xlsx"
    )
)

echo [V] Файлы расшифрованы: Data\ и templates\.people.xlsx
exit /b 0
