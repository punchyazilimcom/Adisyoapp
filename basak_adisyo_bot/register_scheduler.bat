@echo off
REM Her gece 02:00'de "BasakAdisyoBot" görevini olusturur (DÜN'u isler).
REM Yonetici olarak calistirmaniz onerilir.

cd /d "%~dp0"

schtasks /Create ^
    /TN "BasakAdisyoBot" ^
    /TR "\"%~dp0run_daily.bat\"" ^
    /SC DAILY ^
    /ST 02:00 ^
    /RL HIGHEST ^
    /F

if %ERRORLEVEL%==0 (
    echo.
    echo [TAMAM] "BasakAdisyoBot" gorevi olusturuldu - her gece 02:00.
    echo.
    echo Gorevi hemen test etmek icin:   schtasks /Run /TN "BasakAdisyoBot"
    echo Gorevi IPTAL etmek icin:        schtasks /Delete /TN "BasakAdisyoBot" /F
    echo Gorevi gormek icin:             schtasks /Query /TN "BasakAdisyoBot"
) else (
    echo [HATA] Gorev olusturulamadi. Yonetici olarak calistirmayi deneyin.
)
