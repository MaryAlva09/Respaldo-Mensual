# Respaldo Mensual

Sistema automático de respaldo mensual para organizaciones con muchas laptops Windows.  
Respalda archivos y correos de Thunderbird (Gmail IMAP) el día 1 de cada mes, sin intervención del usuario.

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
Luego abre `http://localhost:5000` y completa la configuración inicial.

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
- Página web local en `http://localhost:5000` (o `http://IP-PC100:5000` desde la red)
- Muestra el estado de cada máquina: verde (respaldada), naranja (con errores), rojo (sin respaldo)
- Botón "Respaldar ahora" junto a cada máquina — la laptop lo ejecuta en los próximos 5 minutos
- Buscador, selector de mes, lista de máquinas esperadas
- Protegido con contraseña

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
