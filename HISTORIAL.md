# Historial del proyecto — Respaldo Mensual

Este documento resume cómo evolucionaron los dos programas del proyecto:
el **agente** que corre en cada laptop (`respaldo_mensual.py`) y el
**panel del administrador** que corre en PC100 (`panel_administrador.py`).

---

## 1. Agente de laptop (`respaldo_mensual.py`)

### Versión 1 — Primera versión
- Respaldo automático mensual: el día 1 de cada mes respalda todo el mes
  anterior (ej. el 1 de noviembre respalda octubre).
- Guarda la información en una carpeta con el nombre de la máquina.
- Al terminar, envía la copia a PC100.
- Incluye los correos de Thunderbird, excluyendo Papelera, Spam y Borradores.
- Interfaz gráfica básica para configurar y ejecutar manualmente.

### Versión 2 — Ventana de reintento y visibilidad de errores
- Ventana de reintento ampliada de 1 a 9 días: si la laptop estuvo apagada
  el día 1, el respaldo se completa igual dentro de esos 9 días.
- El programa revisa cada día si el mes ya se respaldó o sigue pendiente.
- Se agregó un registro de errores: qué falló y en qué máquina.
- Primera versión de una pantalla para ver qué máquinas ya se respaldaron
  y cuáles no.

### Versión 3 — Dos fases y puntos de control
Se detectaron dos problemas serios: el programa podía **ciclarse** cada hora
si una copia no terminaba a tiempo, y si la laptop se apagaba a medias, el
respaldo (sobre todo de correos grandes) **volvía a empezar desde cero**.

- Rediseño en dos fases: **recopilar** (copia local, no necesita red) y
  **enviar** (copia a PC100, con reintentos).
- Puntos de control (`checkpoints`) que permiten **retomar exactamente
  donde se quedó**, tanto en archivos como dentro de un mismo buzón de
  correo grande.
- Candado con "latido": evita que dos ejecuciones se encimen, y se libera
  solo si el proceso anterior murió (ej. apagón).
- Doble destino (IP fija + nombre de PC100), para cuando uno de los dos
  falla.

### Versión 4 — Escala a 120 máquinas
- Detección automática del usuario real de cada laptop, **excluyendo
  siempre las cuentas de administrador** (ej. "Propietario"). La detección
  se basa en recorrer el disco, no en qué cuenta tiene la sesión iniciada,
  así que funciona igual sin importar quién instale el programa.
- Modo oculto: carpeta de trabajo oculta, tarea programada en segundo
  plano ejecutada como **SYSTEM** (no depende de que alguien tenga sesión
  iniciada), prioridad de proceso baja para no afectar el rendimiento.
- Reparto de horario: cada laptop calcula su propio minuto dentro de la
  hora, para que las 120 no saturen la red al mismo tiempo.
- Botón de diagnóstico "Probar conexión".
- Instalación silenciosa por línea de comandos (`--instalar`), sin abrir
  ninguna ventana, pensada para desplegarse de forma masiva.

### Versión 5 — Corrección de red pública
- Se identificó que el problema real no era que PC100 estuviera en red
  pública, sino que **las laptops** lo estaban: en red pública, Windows no
  resuelve nombres de equipo, por lo que `\\PC100\...` dejaba de
  encontrarse.
- Solución: usar la **IP fija** de PC100 como destino principal (la
  resolución de nombres no afecta a una IP), dejando el nombre como
  respaldo.
- El diagnóstico de conexión se volvió de tres pasos: nombre, puerto y
  escritura, para saber exactamente dónde falla.

### Versión 6 — Interfaz más amigable
- Rediseño visual completo pensado para una persona no técnica: colores
  suaves, tarjetas, botones grandes con iconos, opciones avanzadas
  ocultas por defecto.
- Se agregó, por pedido, un panel de administrador protegido por
  contraseña **dentro del mismo programa** de la laptop.

### Versión 7 — Revisión del repositorio previo y separación del panel
Se revisó el repositorio anterior de la autora (`MaryAlva09/Respaldo-Mensual`)
y se encontraron fallas reales que se corrigieron en esta base:

- La contraseña del administrador estaba escrita en texto plano dentro
  del código, público en GitHub → ahora se guarda con hash, nunca en el
  código fuente.
- La tarea programada dependía de qué usuario tuviera la sesión iniciada
  (a veces el del administrador, a veces el del empleado, de forma
  inconsistente) → unificado, siempre como SYSTEM.
- El instalador dejaba un acceso directo visible en el escritorio y abría
  la ventana del programa al terminar → eliminado; la instalación es
  completamente silenciosa.
- Dependía de instalar Python en cada una de las 120 laptops → se adoptó
  empaquetar un `.exe` una sola vez (PyInstaller) y distribuirlo ya
  listo.

Además, en esta versión:
- El panel de administrador **se sacó por completo del agente de laptop**;
  ahora vive aparte, en PC100 (ver sección 2). El programa de cada
  laptop quedó con lo mínimo indispensable.
- Se agregó el mecanismo de **"solicitud de respaldo inmediato"**: el
  panel de PC100 deja una marca en la carpeta compartida, y una segunda
  tarea programada en la laptop (cada 5 minutos, de costo casi nulo)
  la detecta y dispara el respaldo sin esperar al horario normal.

---

## 2. Panel del administrador (`panel_administrador.py`)

Nace en la versión 7 del agente, como respuesta a la pregunta de si el
panel debía ser una ventana instalada o algo web.

### Versión 1 — Primera versión
- Se decidió construirlo como **página web local** (con Flask) en vez de
  una ventana de escritorio, para que varias personas puedan verlo a la
  vez desde cualquier navegador de la oficina, sin instalar nada además
  de en PC100.
- Pantalla de inicio de sesión con contraseña y configuración inicial
  (ruta local real detrás de la carpeta compartida de PC100).
- Tarjetas de estado por máquina (verde = respaldada, naranja = con
  errores, rojo = sin respaldo), mostrando fecha, archivos, correos y el
  detalle exacto del error.
- Botón **"Respaldar ahora"** junto a cada máquina (escribe la solicitud
  que el agente de esa laptop recoge en los siguientes minutos).
- Buscador de máquinas, selector de mes, y edición de la lista de
  "máquinas esperadas".

### Versión 1.1 — Arranque automático en PC100 (actual)
- Se agregó `Instalar-PanelAdministrador.ps1`, para que el panel quede
  corriendo en segundo plano desde que PC100 se enciende, sin ventanas
  ni consola visible, reiniciándose solo si llegara a fallar.
- Se activó `threaded=True` en el servidor web para que atienda varias
  pestañas o personas al mismo tiempo sin bloquearse.

---

## Archivos de apoyo del proyecto

- `Desplegar-RespaldoMensual.ps1` — instala el agente en muchas laptops a
  la vez, leyendo una lista de equipos desde un CSV.
- `Instalar-PanelAdministrador.ps1` — instala y deja corriendo el panel
  en PC100.
- `maquinas_ejemplo.csv` — ejemplo del formato que necesita el script de
  despliegue masivo.
