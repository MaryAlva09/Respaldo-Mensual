"""
instalar_tarea.py
Registra la tarea automatica de respaldo en Windows Task Scheduler.
Se llama desde instalar.bat con: python instalar_tarea.py [minutos]
- Sin argumentos: instala tarea permanente (dia 1 y 2 de cada mes, 9AM-6PM cada hora)
- Con numero: instala tarea de prueba que corre en N minutos
"""
import sys, os, subprocess, shutil, socket
from pathlib import Path
from datetime import datetime, timedelta

def get_pythonw():
    pw = shutil.which("pythonw")
    if pw:
        return pw
    py = shutil.which("python") or sys.executable
    pw2 = Path(py).parent / "pythonw.exe"
    if pw2.exists():
        return str(pw2)
    return py

def get_local_username():
    """
    Detecta la cuenta local de la PC (no el administrador/propietario).
    Estrategia:
    1. Derivar del hostname: ULAPC37 -> buscar usuario que contenga "37" o "lapc37"
    2. Buscar en C:/Users la carpeta que NO sea admin/propietario/publico
    3. Tomar la que tenga sesion activa segun query user (excluyendo admins conocidos)
    """
    ADMINS = {"administrator", "administrador", "propietario", "owner",
              "public", "publico", "default", "all users", "default user"}

    hostname = socket.gethostname().upper()  # ej: ULAPC37

    # --- Estrategia 1: derivar del hostname ---
    # ULAPC37 -> cuenta local suele llamarse LAPC37, ulapc37, lapc37, etc.
    digits = "".join(c for c in hostname if c.isdigit())  # "37"

    users_dir = Path("C:/Users")
    candidates = []
    try:
        for d in users_dir.iterdir():
            if not d.is_dir():
                continue
            if d.name.lower() in ADMINS:
                continue
            candidates.append(d)
    except Exception:
        pass

    # Si solo hay un candidato no-admin, ese es
    if len(candidates) == 1:
        return candidates[0].name

    # Si hay varios, buscar el que coincida con el hostname o sus digitos
    if candidates and digits:
        for d in candidates:
            name_lower = d.name.lower()
            if digits in name_lower:
                return d.name
        # Buscar por nombre similar al hostname
        for d in candidates:
            name_lower = d.name.lower()
            host_lower = hostname.lower()
            if host_lower in name_lower or name_lower in host_lower:
                return d.name

    # --- Estrategia 2: query user, excluir admins conocidos ---
    try:
        r = subprocess.run(["query", "user"], capture_output=True, text=True)
        for line in r.stdout.splitlines()[1:]:
            parts = line.split()
            if parts:
                username = parts[0].lstrip(">").strip()
                if username.lower() not in ADMINS and username != "":
                    return username
    except Exception:
        pass

    # --- Estrategia 3: el mas recientemente usado de los candidatos ---
    if candidates:
        return sorted(candidates, key=lambda d: d.stat().st_mtime, reverse=True)[0].name

    return os.environ.get("USERNAME", "")

def make_xml_permanent(python_path, script_path, username):
    """
    Crea XML con triggers cada hora de 9AM a 6PM los dias 1 y 2 de cada mes.
    RunOnlyIfNetworkAvailable garantiza que espera red disponible.
    ExecutionTimeLimit y StartWhenAvailable aseguran que corre aunque la PC
    haya estado apagada.
    """
    # Generar triggers: dia 1 de 9AM a 6PM cada hora, dia 2 igual
    triggers = ""
    # Dias 1 al 12 a las 9AM (cubre semana y media de dias habiles)
    # StartWhenAvailable hace que si la PC estaba apagada, corra al encender
    for dia in range(1, 13):
        dia_str = f"{dia:02d}"
        triggers += f"""    <CalendarTrigger>
      <StartBoundary>2026-04-{dia_str}T09:00:00</StartBoundary>
      <EndBoundary>2099-12-31T23:59:59</EndBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth><Day>{dia}</Day></DaysOfMonth>
        <Months><January/><February/><March/><April/><May/><June/>
                <July/><August/><September/><October/><November/><December/></Months>
      </ScheduleByMonth>
    </CalendarTrigger>
"""
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Respaldo mensual automatico - {username}</Description>
  </RegistrationInfo>
  <Triggers>
{triggers}  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{username}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <ExecutionTimeLimit>PT8H</ExecutionTimeLimit>
    <Enabled>true</Enabled>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <StartWhenAvailable>true</StartWhenAvailable>
  </Settings>
  <Actions>
    <Exec>
      <Command>{python_path}</Command>
      <Arguments>"{script_path}" --auto</Arguments>
    </Exec>
  </Actions>
</Task>"""

def make_xml_test(python_path, script_path, username, run_at_str):
    # EndBoundary = 1 hora despues del inicio
    from datetime import datetime, timedelta
    end_dt  = datetime.strptime(run_at_str, "%Y-%m-%dT%H:%M:%S") + timedelta(hours=1)
    end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%S")
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>PRUEBA respaldo - {username}</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <StartBoundary>{run_at_str}</StartBoundary>
      <EndBoundary>{end_str}</EndBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{username}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <DeleteExpiredTaskAfter>PT1H</DeleteExpiredTaskAfter>
    <Enabled>true</Enabled>
  </Settings>
  <Actions>
    <Exec>
      <Command>{python_path}</Command>
      <Arguments>"{script_path}" --auto</Arguments>
    </Exec>
  </Actions>
</Task>"""

def registrar_tarea(task_name, xml, xml_path, username):
    xml_path.write_text(xml, encoding="utf-16")
    # Registrar la tarea para el usuario local especifico
    result = subprocess.run(
        ["schtasks", "/Create", "/TN", task_name,
         "/XML", str(xml_path), "/F"],
        capture_output=True, text=True
    )
    xml_path.unlink(missing_ok=True)
    return result

def main():
    test_minutes = int(sys.argv[1]) if len(sys.argv) > 1 else 0

    pythonw     = get_pythonw()
    script_path = r"C:\RespaldoMensual\respaldo_mensual.py"
    xml_path    = Path(os.environ.get("TEMP", ".")) / "respaldo_tarea.xml"
    username    = get_local_username()

    print(f"Usuario local detectado: {username}")
    print(f"Python: {pythonw}")

    if not username:
        print("ERROR: No se pudo detectar el usuario local.")
        print("Edita instalar_tarea.py y escribe el nombre de usuario manualmente.")
        sys.exit(1)

    if test_minutes > 0:
        run_at  = datetime.now() + timedelta(minutes=test_minutes)
        run_str = run_at.strftime("%Y-%m-%dT%H:%M:%S")
        task_name = "RespaldoMensualPRUEBA"
        xml = make_xml_test(pythonw, script_path, username, run_str)
        print(f"Instalando tarea de prueba para las {run_at.strftime('%H:%M')}...")
    else:
        task_name = "RespaldoMensualAutomatico"
        xml = make_xml_permanent(pythonw, script_path, username)
        print("Instalando tarea automatica...")
        print("Triggers: dias 1 y 2 de cada mes, cada hora de 9AM a 6PM")

    result = registrar_tarea(task_name, xml, xml_path, username)

    if result.returncode == 0:
        if test_minutes > 0:
            print(f"\nOK - Tarea de prueba creada para las {run_at.strftime('%H:%M')}")
        else:
            print("\nOK - Tarea automatica registrada correctamente.")
            print(f"Usuario: {username}")
            print("Correra el dia 1 y 2 de cada mes de 9AM a 6PM.")
            print("Si la PC estaba apagada, corre en cuanto encienda dentro del horario.")
    else:
        print("\nERROR al registrar la tarea:")
        print(result.stdout)
        print(result.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
