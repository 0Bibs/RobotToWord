@echo off
REM Abre o painel sem deixar uma janela preta de console atras.
cd /d "%~dp0"
start "" pythonw painel.py
