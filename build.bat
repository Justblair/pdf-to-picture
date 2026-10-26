@echo off
REM Builds "PDF to Picture.exe" (a single file, no install needed) into the dist folder.
REM Needs Python 3.9+ from python.org or the Microsoft Store.

cd /d "%~dp0"

if not exist .venv (
    py -m venv .venv || python -m venv .venv || goto :nopython
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt || goto :failed

pyinstaller --noconfirm --onefile --windowed ^
    --name "PDF to Picture" ^
    --collect-all tkinterdnd2 ^
    --collect-all pymupdf ^
    pdf_to_picture.py || goto :failed

echo.
echo Done: dist\PDF to Picture.exe
pause
exit /b 0

:nopython
echo Python wasn't found. Install it from https://www.python.org/downloads/ and run this again.
pause
exit /b 1

:failed
echo Build failed - see the messages above.
pause
exit /b 1
