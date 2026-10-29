<#
    Desplegar-RespaldoMensual.ps1
    ------------------------------------------------------------
    Instala RespaldoMensual.exe en muchas laptops a la vez, leyendo
    la lista de equipos desde un CSV.

    REQUISITOS EN TU PC (la que ejecuta este script):
      - Estar en el dominio / misma red que las laptops.
      - PowerShell Remoting activo en las laptops (WinRM). Si tu
        empresa usa Active Directory, normalmente ya está activo.
        Para activarlo tú mismo en una laptop (una vez, o por GPO):
            Enable-PSRemoting -Force
      - Una cuenta con permisos de administrador en esas laptops.
      - La carpeta "RespaldoMensual" (creada con pyinstaller --onedir)
        en la misma carpeta que este script.

    PREPARA UN ARCHIVO maquinas.csv así (usa comas), junto a este script:

        Maquina,Usuario
        LAPC 10,ULAPC 10
        LAPC 23,ULAPC 23
        LAPC 46,ULAPC 46

    USO:
        .\Desplegar-RespaldoMensual.ps1 `
            -Csv .\maquinas.csv `
            -Destino "\\192.168.1.100\Respaldos" `
            -DestinoAlt "\\PC100\Respaldos"

    Al terminar genera "resultado_despliegue.csv" con el estado de cada
    laptop (OK / ERROR y el motivo), para que sepas cuáles quedaron
    pendientes de revisar a mano.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Csv,

    [Parameter(Mandatory = $true)]
    [string]$Destino,

    [string]$DestinoAlt = "",

    # Usuario de la carpeta compartida de PC100 (ej. PC100\respaldo)
    # Si la carpeta no requiere usuario/contraseña, deja estos parámetros vacíos.
    [string]$UsuarioRed = "",
    [string]$PasswordRed = "",

    # Carpeta generada por "pyinstaller --onedir --noconsole --name RespaldoMensual ..."
    [string]$CarpetaApp = ".\RespaldoMensual",

    # Credenciales de administrador para las laptops (se piden una sola vez)
    [System.Management.Automation.PSCredential]$Credencial
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Csv)) {
    Write-Error "No se encontró el archivo $Csv. Crea el CSV con columnas Maquina,Usuario."
    exit 1
}
if (-not (Test-Path $CarpetaApp)) {
    Write-Error "No se encontró la carpeta '$CarpetaApp'. Genera primero el .exe con pyinstaller --onedir."
    exit 1
}

if (-not $Credencial) {
    Write-Host "Introduce una cuenta con permisos de administrador en las laptops:"
    $Credencial = Get-Credential
}

$equipos = Import-Csv -Path $Csv
$resultados = New-Object System.Collections.Generic.List[object]
$total = $equipos.Count
$i = 0

foreach ($eq in $equipos) {
    $i++
    $maquina = $eq.Maquina.Trim()
    $usuario = $eq.Usuario.Trim()
    Write-Host ""
    Write-Host "[$i/$total] === $maquina (usuario objetivo: $usuario) ===" -ForegroundColor Cyan

    $estado  = "OK"
    $detalle = ""

    try {
        # 1) Verificar que la laptop responde en la red
        if (-not (Test-Connection -ComputerName $maquina -Count 1 -Quiet)) {
            throw "No responde a ping. ¿Está encendida y en la red?"
        }

        # 2) Copiar la carpeta de la app por el recurso administrativo C$
        $destinoRemoto = "\\$maquina\C$\ProgramData\RespaldoMensual"
        Write-Host "  Copiando archivos a $destinoRemoto ..."
        if (-not (Test-Path $destinoRemoto)) {
            New-Item -ItemType Directory -Path $destinoRemoto -Force | Out-Null
        }
        Copy-Item -Path "$CarpetaApp\*" -Destination $destinoRemoto -Recurse -Force

        # 3) Ejecutar la instalación silenciosa en remoto (PowerShell Remoting)
        Write-Host "  Instalando y programando la tarea en $maquina ..."
        $salida = Invoke-Command -ComputerName $maquina -Credential $Credencial -ScriptBlock {
            param($maquinaN, $usuarioN, $destinoN, $destinoAltN, $usuarioRedN, $passwordRedN)
            $exe = "C:\ProgramData\RespaldoMensual\RespaldoMensual.exe"
            if (-not (Test-Path $exe)) {
                return "ERROR: no se encontró $exe tras la copia."
            }
            $argsList = @(
                "--instalar",
                "--maquina",           $maquinaN,
                "--destino",           $destinoN,
                "--usuario-objetivo",  $usuarioN
            )
            if ($destinoAltN)  { $argsList += @("--destino-alt",    $destinoAltN)  }
            if ($usuarioRedN)  { $argsList += @("--usuario-red",    $usuarioRedN)  }
            if ($passwordRedN) { $argsList += @("--password-red",   $passwordRedN) }
            & $exe @argsList
            return "OK"
        } -ArgumentList $maquina, $usuario, $Destino, $DestinoAlt, $UsuarioRed, $PasswordRed

        if ($salida -match "^ERROR") {
            throw $salida
        }
        Write-Host "  Listo." -ForegroundColor Green
    }
    catch {
        $estado  = "ERROR"
        $detalle = $_.Exception.Message
        Write-Host "  FALLÓ: $detalle" -ForegroundColor Red
    }

    $resultados.Add([pscustomobject]@{
        Maquina  = $maquina
        Usuario  = $usuario
        Estado   = $estado
        Detalle  = $detalle
        Fecha    = (Get-Date -Format "yyyy-MM-dd HH:mm")
    })
}

$resultados | Export-Csv -Path ".\resultado_despliegue.csv" -NoTypeInformation -Encoding UTF8

$ok  = ($resultados | Where-Object { $_.Estado -eq "OK" }).Count
$err = $total - $ok
Write-Host ""
Write-Host "============================================================"
Write-Host "Terminado: $ok de $total instaladas correctamente, $err con error."
Write-Host "Revisa el detalle en resultado_despliegue.csv"
Write-Host "============================================================"
