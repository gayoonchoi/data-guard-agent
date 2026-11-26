@echo off
chcp 65001 >nul
cd /d "%~dp0experiments\ollama_standalone"
call RUN_TEST.cmd
