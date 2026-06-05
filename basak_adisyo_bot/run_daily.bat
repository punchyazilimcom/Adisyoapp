@echo off
REM Başak Adisyo Bot — günlük çalıştırıcı (DÜN'ü işler).
REM Görev Zamanlayıcı bu .bat'i çağırır. Loglar logs\ altına tarihli yazılır.

cd /d "%~dp0"

REM Sanal ortamı aktive et (yoksa README'deki kurulum adımlarını izleyin)
if exist "venv\Scripts\activate.bat" (
    call "venv\Scripts\activate.bat"
) else (
    echo [UYARI] venv bulunamadi, sistem Python'u kullaniliyor.
)

if not exist "logs" mkdir "logs"

REM Tarihli log dosyasi (locale'den bagimsiz, PowerShell ile)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set LOGDATE=%%i

echo ============================================== >> "logs\run_%LOGDATE%.log"
echo Calisma: %DATE% %TIME% >> "logs\run_%LOGDATE%.log"

python -m src.main --branch all >> "logs\run_%LOGDATE%.log" 2>&1

echo Cikis kodu: %ERRORLEVEL% >> "logs\run_%LOGDATE%.log"
