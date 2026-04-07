@echo off
title Instalador - Sistema de Respaldo Mensual
color 0A
cls

echo.
echo  ============================================================
echo    INSTALADOR - SISTEMA DE RESPALDO MENSUAL
echo  ============================================================
echo.
echo  Este instalador realizara los siguientes pasos:
echo    1. Verificar si Python esta instalado
echo    2. Instalar Python si es necesario
echo    3. Copiar el programa a C:\RespaldoMensual\
echo    4. Crear acceso directo en el Escritorio
echo    5. Registrar tarea automatica (dia 1 de cada mes, 9 AM)
echo    6. Abrir el programa para configuracion
echo.
echo  Presiona cualquier tecla para comenzar...
pause > nul

:: ─────────────────────────────────────────────
:: PASO 1 — Verificar Python
:: ─────────────────────────────────────────────
echo.
echo  [1/5] Verificando Python...

python --version > nul 2>&1
if %errorlevel% == 0 (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo        OK - %%i encontrado.
    goto COPIAR
)

:: ─────────────────────────────────────────────
:: PASO 2 — Descargar e instalar Python
:: ─────────────────────────────────────────────
echo        Python no encontrado. Descargando instalador...
echo        (esto puede tardar unos minutos segun la conexion)
echo.

if not exist "%TEMP%\RespaldoInstall" mkdir "%TEMP%\RespaldoInstall"

powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile '%TEMP%\RespaldoInstall\python_installer.exe' -UseBasicParsing"

if not exist "%TEMP%\RespaldoInstall\python_installer.exe" (
    echo.
    echo  ERROR: No se pudo descargar Python.
    echo  Descargalo manualmente desde: https://python.org/downloads
    echo  Marca "Add Python to PATH" durante la instalacion.
    echo  Luego vuelve a ejecutar este instalador.
    echo.
    pause
    exit /b 1
)

echo        Descarga completa. Instalando Python...
echo        (no cierres esta ventana)
echo.

"%TEMP%\RespaldoInstall\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

echo        Esperando que termine la instalacion...
timeout /t 5 /nobreak > nul

:: Refrescar PATH sin reiniciar
for /f "tokens=*" %%i in ('powershell -Command "[System.Environment]::GetEnvironmentVariable(\"PATH\",\"Machine\") + \";\" + [System.Environment]::GetEnvironmentVariable(\"PATH\",\"User\")"') do set "PATH=%%i"

python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ADVERTENCIA: Python se instalo pero requiere reiniciar la sesion.
    echo  Por favor:
    echo    1. Cierra esta ventana
    echo    2. Cierra sesion de Windows y vuelve a entrar
    echo    3. Ejecuta este instalador de nuevo
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo        OK - %%i instalado correctamente.

:: ─────────────────────────────────────────────
:: PASO 3 — Copiar programa
:: ─────────────────────────────────────────────
:COPIAR
echo.
echo  [2/5] Copiando programa...

if not exist "C:\RespaldoMensual" mkdir "C:\RespaldoMensual"

set "SCRIPT_ORIGEN=%~dp0respaldo_mensual.py"

if not exist "%SCRIPT_ORIGEN%" (
    echo.
    echo  ERROR: No se encontro respaldo_mensual.py
    echo  Asegurate de que este en la misma carpeta que este instalador.
    echo.
    pause
    exit /b 1
)

copy /Y "%SCRIPT_ORIGEN%" "C:\RespaldoMensual\respaldo_mensual.py" > nul
echo        OK - Programa copiado a C:\RespaldoMensual\

:: ─────────────────────────────────────────────
:: PASO 4 — Acceso directo en el Escritorio
:: ─────────────────────────────────────────────
echo.
echo  [3/5] Creando acceso directo en el Escritorio...

powershell -Command "$py = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source; if (-not $py) { $py = (Get-Command python.exe -ErrorAction SilentlyContinue).Source }; $ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Respaldo Mensual.lnk'); $s.TargetPath = $py; $s.Arguments = 'C:\RespaldoMensual\respaldo_mensual.py'; $s.WorkingDirectory = 'C:\RespaldoMensual'; $s.IconLocation = 'shell32.dll,23'; $s.Description = 'Sistema de Respaldo Mensual'; $s.Save()"

if %errorlevel% == 0 (
    echo        OK - Acceso directo creado.
) else (
    echo        ADVERTENCIA: No se pudo crear el acceso directo.
    echo        Puedes ejecutar el programa directamente desde:
    echo        C:\RespaldoMensual\respaldo_mensual.py
)

:: ─────────────────────────────────────────────
:: PASO 5 — Registrar tarea automatica
:: ─────────────────────────────────────────────
echo.
echo  [4/5] Registrando tarea automatica...

python C:\RespaldoMensual\instalar_tarea.py
if %errorlevel% == 0 (
    echo        OK - Tarea registrada. Correra el dia 1 de cada mes a las 9:00 AM.
) else (
    echo        ADVERTENCIA: No se pudo registrar la tarea automatica.
    echo        Puedes hacerlo despues desde el programa en la pestana Configuracion.
)

:: ─────────────────────────────────────────────
:: PASO 6 — Abrir el programa
:: ─────────────────────────────────────────────
echo.
echo  [5/5] Abriendo el programa...

start "" python "C:\RespaldoMensual\respaldo_mensual.py"

timeout /t 2 /nobreak > nul

echo.
echo  ============================================================
echo.
echo    INSTALACION COMPLETADA
echo.
echo    El programa ya esta abierto en otra ventana.
echo.
echo    En el programa:
echo      1. Ve a la pestana "Configuracion"
echo      2. Ingresa la ruta de red de ULAPC46
echo         Ejemplo:  \\ULAPC46\Respaldos
echo      3. Presiona "Guardar"
echo      4. Presiona "Guardar" y listo
echo         La tarea automatica ya fue instalada por el instalador
echo.
echo  ============================================================
echo.
echo  Instalacion finalizada. Presiona cualquier tecla para cerrar.
pause > nul

exit /b 0
