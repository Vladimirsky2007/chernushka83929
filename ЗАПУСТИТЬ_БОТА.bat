@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Настройка Telegram-бота расписания ===
echo ВАЖНО: сначала отзови токен, отправленный в чат, через @BotFather и создай новый.
set /p BOT_TOKEN=Вставь НОВЫЙ токен бота: 
if "%BOT_TOKEN%"=="" (
  echo Токен не указан.
  pause
  exit /b 1
)
set /p ADMIN_IDS=Введи свой числовой Telegram ID: 
if "%ADMIN_IDS%"=="" set ADMIN_IDS=0
(
  echo BOT_TOKEN=%BOT_TOKEN%
  echo ADMIN_IDS=%ADMIN_IDS%
) > .env
set BOT_TOKEN=
set ADMIN_IDS=
echo.
echo Устанавливаю зависимости...
py -m pip install -r requirements.txt
if errorlevel 1 python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Не удалось установить зависимости. Проверь, установлен ли Python.
  pause
  exit /b 1
)
echo Запускаю бота...
python bot.py
pause
