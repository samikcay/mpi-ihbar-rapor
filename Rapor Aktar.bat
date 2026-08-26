@echo off
rem Ihbar Site Raporu - SAP aktarim programini baslatir.
rem Bu dosyaya cift tiklamak yeterlidir.

cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 goto yok

rem Gerekli paketler kurulu mu? Degilse sessizce kur.
python -c "import win32com.client, tkcalendar, docx" >nul 2>&1
if errorlevel 1 (
    echo Gerekli bilesenler kuruluyor, lutfen bekleyin...
    python -m pip install --quiet pywin32 tkcalendar python-docx
)

where pythonw >nul 2>&1
if errorlevel 1 (
    python rapor.py
) else (
    start "" pythonw rapor.py
)
exit /b 0

:yok
echo.
echo Python bulunamadi.
echo Lutfen python.org adresinden Python 3 kurun ve
echo kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
echo.
pause
