@echo off
chcp 65001 >nul
echo Настройка автоматического шифрования...

:: Указываем git использовать хуки из папки .githooks
git config core.hooksPath .githooks

:: Делаем хуки исполняемыми (для Git Bash)
git update-index --chmod=+x .githooks/pre-commit
git update-index --chmod=+x .githooks/post-checkout
git update-index --chmod=+x .githooks/post-merge

echo.
echo [V] Готово! Теперь git будет автоматически:
echo    - шифровать файлы перед коммитом
echo    - расшифровывать после pull/clone
echo.
pause