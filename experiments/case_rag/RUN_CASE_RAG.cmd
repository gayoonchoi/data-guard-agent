@echo off
chcp 65001 > nul
cd /d "%~dp0"
python --version
if errorlevel 1 (
    echo Python was not found. Please open this folder in PowerShell and check Python.
    pause
    exit /b 1
)
python case_rag_ui.py
if errorlevel 1 (
    echo Case RAG launch failed. Review the error shown above.
    pause
)
