@echo off
setlocal

cd /d "%~dp0"

rem Usar el mismo entorno virtual que CONTROL MATERIALES
set "PYTHON_EXE=C:\Users\NICO\OneDrive\Documentos\PALCA\Scripts Python PALCA\CONTROL MATERIALES\.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

"%PYTHON_EXE%" "%~dp0control_cotizacionesv2.py" %*
if errorlevel 1 (
  echo.
  echo Ocurrio un error al ejecutar el control de cotizaciones.
  pause
  exit /b 1
)

echo.
echo Proceso finalizado.
pause
