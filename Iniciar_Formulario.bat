@echo off
title Formulario de Adicionales y Retornos - Wasion
cd /d "%~dp0"
echo Iniciando Formulario de Adicionales y Retornos...
"C:\Users\Fernando.Carrasco\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.13_qbz5n2kfra8p0\python.exe" app.py
if %errorlevel% neq 0 (
    python app.py
)
pause
