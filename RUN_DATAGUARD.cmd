@echo off
chcp 65001 >nul
cd /d "%~dp0"
python dataguard_gradio_ui.py
if errorlevel 1 pause
