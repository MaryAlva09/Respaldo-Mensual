# Respaldo Mensual

Sistema automático de respaldo mensual para organizaciones con laptops Windows.  
Respalda archivos y correos de Thunderbird durante la ventana de los días 1–9, lunes a sábado, con reanudación de interrupciones.

---

## Archivos del proyecto

| Archivo | Descripción | Dónde va |
|---|---|---|
| `respaldo_mensual.py` | Agente de respaldo — corre en cada laptop | 120 laptops |
| `panel_administrador.py` | Panel web del administrador | Solo en PC100 |
| `Desplegar-RespaldoMensual.ps1` | Instala el agente en todas las laptops a la vez | Tu PC |
| `Instalar-PanelAdministrador.ps1` | Instala el panel en PC100 | PC100 |
| `maquinas_ejemplo.csv` | Plantilla con la lista de laptops | Tu PC |
| `HISTORIAL.md` | Historial completo de versiones | Referencia |

---

## Requisitos

- Windows 10 u 11 en todas las máquinas
- Python 3.11+ con `pip install flask pyinstaller` (solo para generar los `.exe`, una sola vez)
- Thunderbird instalado en las laptops si se quieren respaldar correos
- PC100 con una carpeta compartida en red (`\\192.168.1.X\Respaldos`)
- PowerShell Remoting activo en las laptops para el despliegue masivo

---

## Instalación rápida

### 1. Generar los .exe (una sola vez)
```
pip install flask pyinstaller
pyinstaller --onedir --noconsole --name RespaldoMensual respaldo_mensual.py
pyinstaller --onedir --noconsole --name PanelAdministrador panel_administrador.py
```

### 2. Instalar el panel en PC100
```powershell
.\Instalar-PanelAdministrador.ps1 -CarpetaApp .\PanelAdministrador
```
El panel se inicia en `http://127.0.0.1:5000` solo durante la configuración inicial. Define destino y contraseña, configura en **Configuración** el certificado TLS y la clave privada; al reiniciar, el acceso remoto será por HTTPS. El instalador también programa una copia diaria verificada del historial.

### 3. Desplegar el agente en las laptops
Rellena `maquinas_ejemplo.csv` con los nombres reales y ejecuta:
```powershell
.\Desplegar-RespaldoMensual.ps1 `
    -Csv .\maquinas.csv `
    -Destino "\\192.168.1.X\Respaldos" `
    -DestinoAlt "\\PC100\Respaldos"
```

### 4. Registrar las máquinas esperadas
Entra al panel → **Lista de máquinas esperadas** → pega los 120 nombres.

---

## Cómo funciona

**Agente (cada laptop):**
- Corre como SYSTEM en segundo plano, invisible para el usuario
- El día 1 de cada mes respalda todo el mes anterior (archivos + correos)
- Si no hay red o la laptop estaba apagada, reintenta hasta el día 9
- Guarda primero una copia local (no necesita red en ese momento) y luego la envía a PC100
- Si se interrumpe, retoma exactamente donde se quedó (checkpoint cada 15 segundos)
- Detecta automáticamente al usuario real (excluye cuentas de administrador)
- Usa la IP fija de PC100 como destino principal para evitar problemas con la resolución de nombres en redes públicas

**Panel del administrador (PC100):**
- Arranca por HTTP solo durante la configuración inicial y escucha en 127.0.0.1 hasta configurar TLS.
- Muestra el estado de cada máquina y permite respaldos manuales de meses anteriores.
- `/historial` permite consultar ejecuciones por máquina, usuario, mes o execution_id.
- `/traslados` registra discos externos; el registro no afirma verificación física del disco.
- El borrado local es manual, exige un traslado registrado y confirmación `ELIMINAR`.
- Una tarea diaria copia el historial a la ruta independiente configurada y verifica SHA-256.
- Las contraseñas usan hashes seguros; la contraseña de red no se persiste en `config.json`.

---

## Lo que se respalda por laptop

- `C:\Users\ULAPC XX\Downloads`
- `C:\Users\ULAPC XX\Documents`
- `C:\Users\ULAPC XX\Pictures`
- `C:\Users\ULAPC XX\Videos`
- `C:\Users\ULAPC XX\Desktop`
- Carpetas extra configurables (ej. `C:\Scans`)
- Correos de Thunderbird: Recibidos y Enviados de todas las cuentas IMAP — sin Papelera, Spam, Borradores ni mensajes marcados como borrados

---

## Historial de versiones

Ver `HISTORIAL.md` para el detalle completo de las 7 versiones del agente y las versiones del panel.


## Cumplimiento de BackupDesk SRS v1.1

- El agente procesa el mes calendario anterior durante los días 1–9, de lunes a sábado, desde las 09:00. Al iniciar, reanuda meses pendientes por orden cronológico y conserva checkpoints.
- Los archivos se seleccionan por creación o última modificación. Los mensajes EML se separan por cuenta, conservan carpetas personalizadas y usan fechas fiables del mensaje.
- La transferencia se verifica por tamaño y SHA-256 y admite un intento inicial más tres reintentos (5, 15 y 30 minutos).
- Cada ejecución tiene un `execution_id` propio publicado en `_historial` en PC100.
- Configura `history_backup_root` en una ubicación físicamente independiente. Sin ella, la tarea diaria deja un error registrado en el panel.
- El panel no borra automáticamente los respaldos. La eliminación manual requiere que exista un evento de traslado registrado.
- Para acceso remoto, configura certificado y clave TLS en **Configuración**; sin ellos el panel escucha únicamente en localhost.
