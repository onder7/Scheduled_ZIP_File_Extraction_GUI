@echo off
echo ZIP Excel Islemci - Build Script
echo ============================

echo 1. Gerekli kutuphaneleri yukleniyor...
pip install pyinstaller pillow schedule

echo 2. Ikon olusturuluyor...
python create_icon.py

echo 3. EXE olusturuluyor...
pyinstaller --name="ZIP_Excel_Islemci" ^
    --noconsole ^
    --onefile ^
    --icon=program_icon.ico ^
    --hidden-import=PIL ^
    --hidden-import=PIL._tkinter_finder ^
    --hidden-import=schedule ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --collect-all schedule ^
    --add-data "C:\Windows\System32\msvcp140.dll;." ^
    --add-data "C:\Windows\System32\vcruntime140.dll;." ^
    zip_extractor.py

echo 4. Gereksiz dosyalar temizleniyor...
rmdir /s /q build
del ZIP_Excel_Islemci.spec

echo 5. Islem tamamlandi!
echo EXE dosyasi dist klasorunde olusturuldu.
pause
