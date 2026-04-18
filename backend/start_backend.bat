@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ======================================
echo  SFA / Facturacion - Dev Backend
echo ======================================
echo.

REM --- Crear venv si no existe ---
if not exist ".venv" (
    echo [1/3] Creando entorno virtual con py -3.13...
    py -3.13 -m venv .venv
    if errorlevel 1 goto :error
) else (
    echo [1/3] Entorno virtual ya existe.
)

REM --- Activar venv e instalar deps ---
call .venv\Scripts\activate.bat
if errorlevel 1 goto :error

echo [2/3] Instalando/actualizando dependencias...
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo [3/3] Arrancando servidor...
echo.
python run_dev.py %*

goto :eof

:error
echo.
echo ERROR: fallo la preparacion del entorno.
pause
exit /b 1
