@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0.."

if not exist ".secret" (
    echo [!] Файл .secret не найден в корне проекта.
    exit /b 1
)

if not exist "Data\" (
    echo [!] Папка Data\ не найдена.
    exit /b 1
)

set "SEVEN_ZIP=C:\Program Files\7-Zip\7z.exe"
if not exist "%SEVEN_ZIP%" (
    echo [!] 7-Zip не найден: %SEVEN_ZIP%
    exit /b 1
)

set /p PASSWORD=<.secret

echo [+] Шифрую Data\ в data.7z...
if exist data.7z del /f data.7z

"%SEVEN_ZIP%" a -t7z -mhe=on -p"%PASSWORD%" data.7z .\Data\*
if errorlevel 1 (
    echo [!] Ошибка 7-Zip.
    exit /b 1
)

echo [V] Создан data.7z
exit /b 0
