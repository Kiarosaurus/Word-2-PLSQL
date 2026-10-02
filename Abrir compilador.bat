@echo off
rem Abre el Compilador de reportes APEX desde la carpeta donde se encuentre.
rem La primera vez en un equipo nuevo crea .venv e instala python-docx.
setlocal
cd /d "%~dp0"

set "VENV_PY=.venv\Scripts\python.exe"
set "VENV_PYW=.venv\Scripts\pythonw.exe"

"%VENV_PY%" -c "import sys, docx, tkinter; sys.exit(sys.version_info < (3, 12))" >nul 2>&1
if not errorlevel 1 goto run

echo Preparando el compilador en este equipo (solo la primera vez)...
set "BASE_PY="
py -3 -c "import sys; sys.exit(sys.version_info < (3, 12))" >nul 2>&1
if not errorlevel 1 set "BASE_PY=py -3"
if not defined BASE_PY (
    python -c "import sys; sys.exit(sys.version_info < (3, 12))" >nul 2>&1
    if not errorlevel 1 set "BASE_PY=python"
)
if not defined BASE_PY goto nopython

%BASE_PY% -m venv --clear .venv
if errorlevel 1 goto failed
"%VENV_PY%" -m pip install --disable-pip-version-check --quiet -r requirements.txt
if errorlevel 1 goto failed

:run
start "" "%VENV_PYW%" "%~dp0launch_gui.pyw"
exit /b 0

:nopython
echo.
echo No se encontro Python 3.12 o superior.
echo Instalelo desde https://www.python.org/downloads/ marcando "Add python.exe to PATH"
echo y vuelva a abrir el compilador.
pause
exit /b 1

:failed
echo.
echo No se pudo preparar el entorno (.venv). Revise la conexion a Internet
echo o instale manualmente: python -m pip install -r requirements.txt
pause
exit /b 1
