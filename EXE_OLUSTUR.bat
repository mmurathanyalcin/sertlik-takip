@echo off
chcp 65001 >nul
echo ============================================
echo   Sertlik Takip - EXE Olusturma Sihirbazi
echo ============================================
echo.

REM Python kontrolu
python --version >nul 2>&1
if errorlevel 1 (
    echo HATA: Python bulunamadi.
    echo Once https://www.python.org/downloads/ adresinden Python kurun.
    echo Kurulum sirasinda "Add Python to PATH" kutusunu isaretlemeyi unutmayin.
    pause
    exit /b 1
)

echo [1/3] Gerekli kutuphaneler kuruluyor...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo HATA: Kutuphaneler kurulamadi.
    pause
    exit /b 1
)

echo.
echo [2/3] EXE dosyasi olusturuluyor (bu islem 1-2 dakika surebilir)...
python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "SertlikTakip" ^
    --collect-all tkinterdnd2 ^
    sertlik_takip.py

if errorlevel 1 (
    echo HATA: EXE olusturulamadi.
    pause
    exit /b 1
)

echo.
echo [3/3] Tamamlandi!
echo.
echo EXE dosyaniz hazir: dist\SertlikTakip.exe
echo Bu dosyayi istediginiz bilgisayara kopyalayip
echo cift tiklayarak calistirabilirsiniz (Python kurulu olmasi gerekmez).
echo.
pause
