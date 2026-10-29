<#
    Instalar-PanelAdministrador.ps1
    ------------------------------------------------------------
    Instala el Panel del Administrador en PC100 y lo deja arrancando
    solo cada vez que la PC se enciende (sin ventanas, como un
    servicio). A diferencia del agente de las laptops, esta carpeta
    SÍ puede quedar visible: es la herramienta del propio encargado,
    no hay nada que ocultarle a nadie en PC100.

    REQUISITOS:
      - Generar antes la carpeta "PanelAdministrador" con:
            pip install flask pyinstaller
            pyinstaller --onedir --noconsole --name PanelAdministrador panel_administrador.py
        (puedes correr ese paso en PC100 mismo, o copiar la carpeta
        dist\PanelAdministrador generada en otra PC con Windows)
      - Ejecutar este script COMO ADMINISTRADOR, en PC100.

    USO:
        .\Instalar-PanelAdministrador.ps1 -CarpetaApp .\PanelAdministrador

    Tras ejecutarlo, abre http://localhost:5000 en el navegador de PC100
    (la primera vez te pedirá la carpeta local de los respaldos y la
    contraseña del panel).
#>

param(
    [string]$CarpetaApp = ".\PanelAdministrador",
    [string]$CarpetaInstalacion = "C:\PanelRespaldo",
    [string]$NombreTarea = "PanelRespaldoWeb"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $CarpetaApp)) {
    Write-Error "No se encontró '$CarpetaApp'. Genera antes el .exe con pyinstaller (ver instrucciones arriba)."
    exit 1
}

$exe = Join-Path $CarpetaInstalacion "PanelAdministrador.exe"

Write-Host "Copiando archivos a $CarpetaInstalacion ..."
New-Item -ItemType Directory -Path $CarpetaInstalacion -Force | Out-Null
Copy-Item -Path "$CarpetaApp\*" -Destination $CarpetaInstalacion -Recurse -Force

if (-not (Test-Path $exe)) {
    Write-Error "No se encontró $exe tras copiar. Revisa que pyinstaller haya generado PanelAdministrador.exe"
    exit 1
}

Write-Host "Creando tarea programada '$NombreTarea' (arranca solo con Windows, sin ventanas)..."

$accion     = New-ScheduledTaskAction -Execute $exe
$disparador = New-ScheduledTaskTrigger -AtStartup
$principal  = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$ajustes    = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero) `
                -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $NombreTarea -Action $accion -Trigger $disparador `
    -Principal $principal -Settings $ajustes -Force | Out-Null

Write-Host "Iniciando el panel ahora mismo (no hace falta reiniciar PC100)..."
Start-ScheduledTask -TaskName $NombreTarea
Start-Sleep -Seconds 2

Write-Host ""
Write-Host "============================================================"
Write-Host "Listo. El panel ya está corriendo en segundo plano."
Write-Host "Ábrelo en el navegador: http://localhost:5000"
Write-Host "Desde otra PC de la oficina: http://<IP de PC100>:5000"
Write-Host ""
Write-Host "Si la tarea no arranca (revisa con: Get-ScheduledTask -TaskName $NombreTarea),"
Write-Host "confirma que ejecutaste este script como Administrador."
Write-Host "============================================================"
