@echo off
chcp 65001 >nul
cd /d "%~dp0experiments\case_rag"
call RUN_CASE_RAG.cmd
