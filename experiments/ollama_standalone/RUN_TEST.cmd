@echo off
chcp 65001 >nul
cd /d "%~dp0"
python ollama_test_ui.py
if errorlevel 1 (
    echo.
    echo [ERROR] Python or Gradio was not found, or the test program failed.
    pause
)
