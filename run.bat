@echo off
chcp 65001 > nul
echo Запуск приложения...

:: Устанавливаем зависимости, если нужно
pip install -r "%~dp0requirements.txt" --quiet

:: Запускаем приложение
python "%~dp0app.py"

if errorlevel 1 (
    echo.
    echo Ошибка запуска. Убедитесь, что Python установлен и доступен в PATH.
    pause
)
