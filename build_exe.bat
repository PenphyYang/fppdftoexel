@echo off
setlocal

set PYTHON_EXE=E:\pyenv\pyenv-win\versions\3.8.10\python.exe
set APP_NAME=invoice_analysis_tool

if not exist "%PYTHON_EXE%" (
    echo Python not found: %PYTHON_EXE%
    exit /b 1
)

"%PYTHON_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --icon assets\invoice_tool_logo.ico ^
  --name %APP_NAME% ^
  invoice_pdf_to_excel_gui.py

echo.
echo Build complete. EXE path: dist\%APP_NAME%.exe
