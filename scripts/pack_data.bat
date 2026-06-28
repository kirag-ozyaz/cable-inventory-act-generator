@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0.."

if not exist ".secret" (
    echo [!] Файл .secret не найден в корне проекта.
    exit /b 1
)

set "HAS_DATA=0"
set "HAS_PEOPLE=0"
if exist "Data\*" set "HAS_DATA=1"
if exist "templates\.people.xlsx" set "HAS_PEOPLE=1"

if "%HAS_DATA%"=="0" if "%HAS_PEOPLE%"=="0" (
    echo [!] Нет Data\ и templates\.people.xlsx для архивации.
    exit /b 1
)

set "SEVEN_ZIP=C:\Program Files\7-Zip\7z.exe"
if not exist "%SEVEN_ZIP%" (
    echo [!] 7-Zip не найден: %SEVEN_ZIP%
    exit /b 1
)

set /p PASSWORD=<.secret

echo [+] Шифрую Data\ и templates\.people.xlsx в data.7z...
if exist data.7z del /f data.7z

set "ARCHIVE_ARGS="
if "%HAS_DATA%"=="1" set "ARCHIVE_ARGS=.\Data\"
if "%HAS_PEOPLE%"=="1" set "ARCHIVE_ARGS=%ARCHIVE_ARGS% .\templates\.people.xlsx"

"%SEVEN_ZIP%" a -t7z -mhe=on -p"%PASSWORD%" data.7z %ARCHIVE_ARGS%
if errorlevel 1 (
    echo [!] Ошибка 7-Zip.
    exit /b 1
)

echo [V] Создан data.7z
exit /b 0
