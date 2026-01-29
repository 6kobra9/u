@echo off
chcp 65001 >nul
echo ========================================
echo  Запуск веб-додатку MSSQL
echo ========================================
echo.

cd /d "%~dp0"

echo Перевірка підключення до бази даних...
python test_connection.py
if errorlevel 1 (
    echo.
    echo ПОМИЛКА: Не вдалося підключитися до бази даних!
    echo Перевірте налаштування в config.py
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Запуск веб-сервера...
echo ========================================
echo.
echo Додаток буде доступний за адресою:
echo   http://127.0.0.1:5000
echo.
echo Для зупинки натисніть Ctrl+C
echo.

python app.py

pause
