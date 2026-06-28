@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo Настройка git hooks для шифрования Data/ и templates/.people.xlsx ^-^> data.7z...
git config core.hooksPath .githooks

if exist ".githooks\pre-commit" (
    git add --chmod=+x .githooks/pre-commit .githooks/post-checkout .githooks/post-merge .githooks/_common.sh 2>nul
)

echo.
echo [V] Готово!
echo    pre-commit    - шифрует Data/ перед коммитом
echo    post-merge    - расшифровывает после pull
echo    post-checkout - расшифровывает после clone/checkout
echo.
echo Нужны: 7-Zip и файл .secret в корне проекта.
echo Ручная упаковка:   scripts\pack\pack_data.bat
echo Ручная распаковка: scripts\pack\unpack_data.bat
echo.
pause
