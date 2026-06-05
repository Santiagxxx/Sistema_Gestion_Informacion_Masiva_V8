@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set PYTHON_CMD=py
) else (
    set PYTHON_CMD=python
)

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo No fue posible crear el entorno. Verifica que Python este instalado.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
echo Instalando o validando dependencias...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo No fue posible instalar las dependencias.
    pause
    exit /b 1
)

echo Abriendo el sistema...
python -m streamlit run app.py
pause
