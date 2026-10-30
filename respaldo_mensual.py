#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Respaldo Mensual v4
  python respaldo_mensual.py                     -> interfaz gráfica
  python respaldo_mensual.py --catchup            -> tarea programada (ejecución normal, cada hora)
  python respaldo_mensual.py --auto               -> respalda el mes anterior sin condiciones
  python respaldo_mensual.py --instalar [opciones] -> instala sin abrir ventanas (ver instalar())

Detección de usuario: el programa ignora el usuario administrador ("Propietario" y similares) y
respalda solo el perfil de usuario normal que encuentra en C:\\Users. Al correr como SYSTEM (tarea
programada), esto funciona sin importar qué usuario tenga la sesión iniciada.

FASE 1 (recopilar): copia lo del mes a una carpeta local oculta (no necesita red).
FASE 2 (enviar):    la envía a PC100 (reintenta; retoma archivos y correos a medias).
"""
import os, sys, json, glob, re, time, shutil, socket, hashlib
import queue, threading, subprocess, tempfile, mailbox, sqlite3
import configparser
import email.utils
from email.parser import BytesParser, HeaderParser
from email import policy
from datetime import datetime, date, timezone
from pathlib import Path
from email.generator import BytesGenerator
from email.header import decode_header, make_header

APP_DIR_NAME = "SyncAgent"              # nombre neutro: no debe llamar la atención en C:\ProgramData
TASK_NAME = "SyncAgentTask"
_PROGRAMDATA = os.environ.get("ProgramData")
if _PROGRAMDATA and os.name == "nt":
    CFG_DIR = Path(_PROGRAMDATA) / APP_DIR_NAME          # mismo lugar sin importar qué cuenta ejecute el programa
else:
    CFG_DIR = Path.home() / f".{APP_DIR_NAME.lower()}"   # equipo de pruebas (no Windows)
CFG_FILE = CFG_DIR / "config.json"
LOG_FILE = CFG_DIR / "log.txt"
LOCK_FILE = CFG_DIR / "running.lock"
PROG_DIR = CFG_DIR / "progress"
WINDOW_DAYS = 9
CK_SECONDS = 15

USER_SUBDIRS = ["Downloads", "Documents", "Pictures", "Videos", "Desktop"]
SKIP_USER_NAMES = {"public", "default", "default user", "all users"}
DEFAULTS = {
    "machine_name": socket.gethostname().upper(),
    "destination": r"\\192.168.1.100\Respaldos",     # usar IP fija de PC100 como principal
    "destination_alt": r"\\PC100\Respaldos",          # nombre como respaldo del destino
    "local_root": str(Path(os.environ.get("SystemDrive", "C:") + "\\") / "ProgramData" / APP_DIR_NAME / "data"),
    "delete_local_after_send": False,
    "sources": [],                                    # carpetas EXTRA (ej. scans compartidos), no las del usuario
    "thunderbird": "",                                 # vacío = detectar automático del usuario objetivo
    "target_user": "",                                 # vacío = autodetectar (excluye excluded_users)
    "excluded_users": ["propietario", "administrador", "administrator", "admin", "wdagutilityaccount"],
    "run_as_system": True,
    "net_user": "",   # usuario de la carpeta compartida de PC100, ej. PC100\respaldo
    "net_pass": "",   # contraseña (se guarda en config.json; para mayor seguridad usa cmdkey manualmente)
    "last_done": "", "last_status": "", "last_attempt": "",
}
for _p in (r"C:\Scans", r"C:\Scan", r"C:\Escaneos", r"C:\Escaner"):
    if Path(_p).is_dir():
        DEFAULTS["sources"].append(_p)

SKIP_MAIL = {
    "trash", "junk", "spam", "drafts", "draft", "deleted",
    "deleted items", "deleted messages", "bin", "papelera",
    "borradores", "correo no deseado", "elementos eliminados",
    "basura", "unsent messages", "bulk mail", "outbox",
    "bandeja de salida", "all mail", "todos"
}
SENT_MAIL = {
    "sent", "sent mail", "sent items", "enviados",
    "elementos enviados", "correo enviado"
}
SKIP_DIRS = {"appdata", "$recycle.bin", "node_modules", ".git", "__pycache__"}


class Stopped(Exception):
    pass


# ------------------------------------------------------------------ config
def load_cfg():
    cfg = json.loads(json.dumps(DEFAULTS))  # copia profunda
    try:
        cfg.update(json.loads(CFG_FILE.read_text(encoding="utf-8")))
    except Exception:
        pass
    return cfg


def save_cfg(cfg):
    CFG_DIR.mkdir(parents=True, exist_ok=True)
    CFG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def hide_path(p: Path):
    """Oculta una carpeta (atributos Oculto+Sistema). No se hace pasar por un archivo de Microsoft/Windows."""
    try:
        p.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            subprocess.run(["attrib", "+h", "+s", str(p)], capture_output=True)
    except Exception:
        pass


def lower_priority():
    if os.name == "nt":
        try:
            import ctypes
            h = ctypes.windll.kernel32.GetCurrentProcess()
            ctypes.windll.kernel32.SetPriorityClass(h, 0x00004000)   # BELOW_NORMAL_PRIORITY_CLASS
        except Exception:
            pass


def prev_month(today=None):
    t = today or date.today()
    return (t.year - 1, 12) if t.month == 1 else (t.year, t.month - 1)


def month_range(y, m):
    return datetime(y, m, 1), datetime(y + (m == 12), 1 if m == 12 else m + 1, 1)


def load_prog(tag):
    try:
        return json.loads((PROG_DIR / f"{tag}.json").read_text(encoding="utf-8"))
    except Exception:
        return {"files_done": False, "collect_done": False, "send_done": False,
                "files": 0, "mails": 0, "errors": [], "mail": {}}


def save_prog(tag, p):
    PROG_DIR.mkdir(parents=True, exist_ok=True)
    tmp = PROG_DIR / f"{tag}.tmp"
    tmp.write_text(json.dumps(p), encoding="utf-8")
    os.replace(tmp, PROG_DIR / f"{tag}.json")


def reset_month(cfg, tag):
    """'Empezar de nuevo': borra lo recopilado localmente y los puntos de control de ese mes."""
    shutil.rmtree(local_dir(cfg, tag), ignore_errors=True)
    for p in PROG_DIR.glob(f"{tag}*"):
        shutil.rmtree(p, ignore_errors=True) if p.is_dir() else p.unlink(missing_ok=True)


# ------------------------------------------------------------------ detección del usuario objetivo
def get_user_profiles():
    root = Path(os.environ.get("SystemDrive", "C:") + "\\") / "Users"
    if not root.is_dir():
        return []
    out = []
    for d in sorted(root.iterdir()):
        if d.is_dir() and d.name.lower() not in SKIP_USER_NAMES and not d.name.startswith("."):
            out.append(d)
    return out


def _match_user(target_raw: str, profile_name: str) -> bool:
    """
    Coincidencia flexible entre el nombre de usuario objetivo configurado y
    el nombre real de la carpeta en C:\\Users.

    Ejemplos que deben coincidir:
      "ULAPC 46"  ↔  "Ulapc46"   (el sistema guardó sin espacio)
      "ULAPC 46"  ↔  "ULAPC 46"  (coincidencia exacta sin espacios)
      "ulapc"     ↔  "Ulapc46"   (prefijo)
    """
    t = target_raw.lower()
    p = profile_name.lower()
    t_nsp = t.replace(" ", "")
    p_nsp = p.replace(" ", "")
    return (t in p) or (t_nsp in p_nsp) or (p in t) or (p_nsp in t_nsp)


def pick_target_users(cfg, log=None):
    profiles = get_user_profiles()
    target = cfg.get("target_user", "").strip()
    if target:
        cands = [d for d in profiles if _match_user(target, d.name)]
        if not cands and log:
            log(f"Usuario objetivo '{target}' no encontrado. "
                f"Perfiles en esta máquina: {[d.name for d in profiles]}. "
                f"Prueba con solo la parte numérica (ej. '46') o el nombre exacto de la carpeta.")
        return cands
    excluded = [x.strip().lower() for x in cfg.get("excluded_users", []) if x.strip()]
    cands = [d for d in profiles if not any(x and x in d.name.lower() for x in excluded)]
    if log:
        if not cands:
            log(f"ERROR: no se detectó un usuario distinto al administrador. "
                f"Perfiles vistos: {[d.name for d in profiles]}. "
                f"Define 'Usuario objetivo' manualmente.")
        elif len(cands) > 1:
            log(f"Aviso: {len(cands)} usuarios candidatos {[d.name for d in cands]}; "
                f"se respaldan todos. Define 'Usuario objetivo' si solo quieres uno.")
    return cands


def resolve_all_sources(cfg, log):
    users = pick_target_users(cfg, log)
    extra = [s for s in cfg.get("sources", []) if s.strip()]
    found = []
    for u in users:
        for sub in USER_SUBDIRS:
            p = u / sub
            if p.is_dir():
                found.append((f"{u.name}-{sub}", p))
    return users, extra, found


def find_thunderbird_profiles(user_home: Path):
    """
    Busca perfiles de Thunderbird bajo el perfil de un usuario.
    Cubre los tres lugares donde puede estar según la versión y configuración:
      1. AppData\\Roaming\\Thunderbird\\Profiles\\<perfil>   (instalación estándar)
      2. AppData\\Local\\Thunderbird\\Profiles\\<perfil>     (algunas instalaciones empresariales)
      3. Cualquier carpeta profiles.ini, para detectar rutas personalizadas
    Devuelve lista de rutas de perfil válidas (tienen ImapMail, Mail, o prefs.js).
    """
    candidates = []

    # Rutas estándar donde Thunderbird guarda los perfiles
    for base_rel in (
        Path("AppData") / "Roaming" / "Thunderbird" / "Profiles",
        Path("AppData") / "Local"   / "Thunderbird" / "Profiles",
    ):
        base = user_home / base_rel
        if base.is_dir():
            for p in base.iterdir():
                if p.is_dir():
                    candidates.append(p)

    # Leer profiles.ini por si hay rutas personalizadas
    for ini_rel in (
        Path("AppData") / "Roaming" / "Thunderbird" / "profiles.ini",
        Path("AppData") / "Local"   / "Thunderbird" / "profiles.ini",
    ):
        ini = user_home / ini_rel
        if not ini.exists():
            continue
        try:
            content = ini.read_text(encoding="utf-8", errors="replace")
            for line in content.splitlines():
                line = line.strip()
                if not line.lower().startswith("path="):
                    continue
                rel_path = line[5:].strip().replace("/", os.sep)
                # IsRelative=1 → relativa al directorio del .ini
                candidate = ini.parent / rel_path
                if candidate.is_dir():
                    candidates.append(candidate)
                # también intentar ruta absoluta tal cual
                candidate2 = Path(rel_path)
                if candidate2.is_dir():
                    candidates.append(candidate2)
        except Exception:
            pass

    # Filtrar: un perfil válido tiene prefs.js, ImapMail o Mail
    valid = []
    seen = set()
    for p in candidates:
        key = str(p.resolve())
        if key in seen:
            continue
        seen.add(key)
        has_imap = (p / "ImapMail").is_dir()
        has_mail = (p / "Mail").is_dir()
        has_prefs = (p / "prefs.js").is_file()
        if has_imap or has_mail or has_prefs:
            valid.append(p)

    return valid


def resolve_thunderbird(cfg, users):
    """
    Devuelve lista de (etiqueta, ruta_de_perfil) para respaldar.
    Si está configurado manualmente, lo usa directo.
    Si no, busca en todos los usuarios objetivo y devuelve todos los perfiles encontrados.
    """
    manual = cfg.get("thunderbird", "").strip()
    if manual:
        p = Path(manual)
        if p.is_dir():
            return [("manual", manual)]
        # El manual apunta a la carpeta Profiles; probar un nivel arriba
        if p.parent.is_dir() and (p.parent / "prefs.js").is_file():
            return [("manual", str(p.parent))]
        return [("manual-no-encontrado", manual)]

    out = []
    for u in users:
        profs = find_thunderbird_profiles(u)
        for prof in profs:
            out.append((u.name, str(prof)))

    return out


# ------------------------------------------------------------------
# VERSION 5: IDENTIFICACION DE CUENTAS THUNDERBIRD

def read_thunderbird_prefs(profile_path):
    """
    Lee las preferencias de Thunderbird.
    No modifica prefs.js.
    """
    prefs = {}
    prefs_file = Path(profile_path) / "prefs.js"

    if not prefs_file.is_file():
        return prefs

    pattern = re.compile(
        r'user_pref\("([^"]+)",\s*("(?:\\.|[^"\\])*"|[^)]+)\);'
    )

    try:
        content = prefs_file.read_text(
            encoding="utf-8",
            errors="replace"
        )

        for key, value in pattern.findall(content):
            value = value.strip()

            if value.startswith('"'):
                try:
                    value = json.loads(value)
                except Exception:
                    value = value.strip('"')

            prefs[key] = value

    except Exception:
        pass

    return prefs


def get_thunderbird_accounts(profile_path):
    """
    Identifica las cuentas configuradas en un perfil.
    Devuelve etiquetas anonimizadas y sus almacenes.
    """
    prefs = read_thunderbird_prefs(profile_path)

    account_ids = str(
        prefs.get("mail.accountmanager.accounts", "")
    ).split(",")

    accounts = []

    for account_id in account_ids:
        account_id = account_id.strip()

        if not account_id:
            continue

        server_id = prefs.get(
            f"mail.account.{account_id}.server"
        )

        if not server_id:
            continue

        server_type = prefs.get(
            f"mail.server.{server_id}.type", ""
        )

        if server_type not in ("imap", "pop3", "none"):
            continue

        directory = prefs.get(
            f"mail.server.{server_id}.directory"
        )

        if not directory:
            continue

        directory = Path(str(directory))

        if not directory.is_absolute():
            directory = Path(profile_path) / directory

        accounts.append({
            "label": f"Correo {len(accounts) + 1}",
            "server_id": server_id,
            "type": server_type,
            "directory": directory,
        })

    return accounts


def get_account_folders(account):
    """
    Identifica carpetas de enviados y recibidos.
    La clasificación definitiva se validará con
    la configuración de carpetas especiales.
    """
    base = account["directory"]

    return {
        "base": base,
        "exists": base.is_dir(),
    }


def thunderbird_summary(cfg, users):
    """
    Devuelve un texto descriptivo del estado de la detección de Thunderbird,
    para mostrarlo en la interfaz.
    """
    manual = cfg.get("thunderbird", "").strip()
    if manual:
        exists = Path(manual).is_dir()
        return ("manual", f"Perfil manual: {manual}" + ("" if exists else " ⚠ NO ENCONTRADO"))

    results = resolve_thunderbird(cfg, users)
    if not results:
        # Explicar por qué no se encontró nada
        msgs = []
        for u in users:
            tb_base = u / "AppData" / "Roaming" / "Thunderbird"
            if not tb_base.exists():
                msgs.append(f"{u.name}: Thunderbird no instalado (no existe {tb_base})")
            elif not (tb_base / "Profiles").exists():
                msgs.append(f"{u.name}: Thunderbird instalado pero sin carpeta Profiles")
            else:
                profs_dir = tb_base / "Profiles"
                subdirs = list(profs_dir.iterdir()) if profs_dir.is_dir() else []
                if not subdirs:
                    msgs.append(f"{u.name}: carpeta Profiles vacía")
                else:
                    msgs.append(f"{u.name}: perfiles encontrados ({[d.name for d in subdirs]}) "
                                f"pero ninguno tiene prefs.js, ImapMail ni Mail")
        if not msgs:
            msgs = ["No se detectaron usuarios objetivo"]
        return ("error", "Thunderbird no detectado → " + " | ".join(msgs))

    labels = [f"{lbl}: {Path(path).name}" for lbl, path in results]
    return ("ok", f"Thunderbird detectado ({len(results)} perfil/es): " + ", ".join(labels))


# ------------------------------------------------------------------ candado con latido
def acquire_lock():
    CFG_DIR.mkdir(parents=True, exist_ok=True)
    for _ in range(2):
        try:
            fd = os.open(str(LOCK_FILE), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode()); os.close(fd)
            return True
        except FileExistsError:
            try:
                if time.time() - LOCK_FILE.stat().st_mtime > 180:
                    LOCK_FILE.unlink(); continue
            except OSError:
                pass
            return False
    return False


def start_heartbeat():
    ev = threading.Event()

    def loop():
        while not ev.wait(20):
            try:
                os.utime(str(LOCK_FILE))
            except Exception:
                pass
    threading.Thread(target=loop, daemon=True).start()
    return ev


# ------------------------------------------------------------------ red / estado
def unc_host(path):
    p = path.strip()
    return p[2:].split("\\")[0] if p.startswith("\\\\") else None


def destinations(cfg):
    raw = [d.strip() for d in (cfg["destination"], cfg.get("destination_alt", "")) if d and d.strip()]
    if os.name == "nt":
        # Normalizar: convertir barras hacia adelante a backslash en rutas UNC de Windows.
        # Ocurre cuando la ruta se guardó con / en lugar de \ en el JSON.
        fixed = []
        for d in raw:
            if d.startswith("//") or d.startswith("\\\\"):
                d = d.replace("/", "\\")
            fixed.append(d)
        return fixed
    return raw


def dest_reachable(dest):
    host = unc_host(dest)
    if host:
        try:
            socket.create_connection((host, 445), 4).close()
        except Exception as e:
            return False, f"Sin red hacia {host} (puerto 445): {e}"
    try:
        Path(dest).mkdir(parents=True, exist_ok=True)
        return True, ""
    except Exception as e:
        return False, f"No se puede escribir en {dest}: {e}"


def test_destinations(cfg, log):
    """Diagnóstico detallado para el botón 'Probar conexión'."""
    results = []
    for dest in destinations(cfg):
        host = unc_host(dest)
        if host:
            try:
                ip = socket.gethostbyname(host)
                log(f"{dest}: el nombre '{host}' se resuelve a {ip}")
            except Exception as e:
                log(f"{dest}: [FALLA] no se pudo resolver el nombre '{host}' (red pública, DNS, o IP incorrecta): {e}")
                results.append((dest, False)); continue
            try:
                socket.create_connection((host, 445), 4).close()
                log(f"{dest}: puerto 445 (compartir archivos) abierto en {host}")
            except Exception as e:
                log(f"{dest}: [FALLA] sin conexión al puerto 445 de {host} (firewall de PC100 o perfil de red): {e}")
                results.append((dest, False)); continue
        try:
            p = Path(dest); p.mkdir(parents=True, exist_ok=True)
            t = p / "_prueba_conexion.tmp"; t.write_text("ok"); t.unlink()
            log(f"{dest}: escritura correcta")
            results.append((dest, True))
        except Exception as e:
            log(f"{dest}: [FALLA] no se pudo escribir (usuario/contraseña o permisos de la carpeta): {e}")
            results.append((dest, False))
    return results


def write_status(cfg, tag, res, dest):
    try:
        d = Path(dest) / "_estado"
        d.mkdir(parents=True, exist_ok=True)
        f = d / f"{cfg['machine_name'].strip()}.json"
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            data = {}
        data[tag] = {"status": res["status"], "files": res["files"], "mails": res["mails"],
                     "errors": res["errors"][:200], "when": datetime.now().strftime("%Y-%m-%d %H:%M")}
        f.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


# ------------------------------------------------------------------ FASE 1: recopilar (local)
def local_dir(cfg, tag):
    return Path(cfg["local_root"]) / cfg["machine_name"].strip() / tag


def collect(cfg, y, m, prog, log, stop):
    tag = f"{y}-{m:02d}"
    start, end = month_range(y, m)
    ts0, ts1 = start.timestamp(), end.timestamp()
    dest_root = local_dir(cfg, tag)
    lr = os.path.normcase(os.path.abspath(cfg["local_root"]))
    errors = []

    def err(msg):
        errors.append(msg); log("  ERROR: " + msg)

    users, extra, user_sources = resolve_all_sources(cfg, log)
    log(f"== Recopilando {tag} — usuario(s) objetivo: {[u.name for u in users] or 'NINGUNO'}")
    all_sources = [(Path(e).name or Path(e).drive.replace(":", ""), Path(e)) for e in extra] + user_sources

    if not prog["files_done"]:
        for label, sp in all_sources:
            if not sp.is_dir():
                log(f"(omitido, no existe) {sp}"); continue
            log(f"Archivos: {sp}")
            for root, dirs, files in os.walk(sp):
                dirs[:] = [d for d in dirs if d.lower() not in SKIP_DIRS
                           and os.path.normcase(os.path.abspath(os.path.join(root, d))) != lr]
                for f in files:
                    if stop():
                        raise Stopped()
                    p = Path(root) / f
                    try:
                        st = p.stat()
                        if not (ts0 <= st.st_mtime < ts1):
                            continue
                        out = dest_root / "Archivos" / label / p.relative_to(sp)
                        if out.exists() and out.stat().st_size == st.st_size:
                            continue
                        out.parent.mkdir(parents=True, exist_ok=True)
                        part = out.with_name(out.name + ".part")
                        shutil.copy2(p, part)
                        os.replace(part, out)
                        prog["files"] += 1
                    except Exception as e:
                        err(f"Archivo {p}: {e}")
        prog["files_done"] = not errors
        save_prog(tag, prog)
        log(f"  archivos nuevos copiados en total: {prog['files']}")


    tb_list = resolve_thunderbird(cfg, users)

    if tb_list:

        account_number = 0

        for ulabel, tb in tb_list:

            profile = Path(tb)

            for sub in ("ImapMail", "Mail"):

                base = profile / sub

                if not base.is_dir():
                    continue

                # Cada carpeta de cuenta/servidor representa
                # una cuenta de correo.
                accounts = sorted(
                    p for p in base.iterdir()
                    if p.is_dir()
                )

                for account_dir in accounts:

                    if stop():
                        raise Stopped()

                    #Ignorar carpetas especiales, no son cuentas de correo.
                    if account_dir.name.lower() in {
                        "local folders", "carpetas locales", "local folders.sbd"
                    }:
                        continue

                    account_number += 1

                    account_name = f"Correo {account_number}"

                    log(
                        f"Cuenta detectada: {account_name}"
                    )

                    backup_mail(
                        account_dir,
                        account_name,
                        dest_root / "Correos" / account_name,
                        start,
                        end,
                        prog,
                        tag,
                        log,
                        err,
                        stop
                    )

        prog["mails"] = sum(
            s.get("count", 0)
            for k, s in prog["mail"].items()
            if k.startswith("EML/")
        )

        log(f"  correos respaldados: {prog['mails']}")

    else:
        err("No se encontró perfil de Thunderbird del usuario objetivo.")

    prog["errors"] = errors
    prog["collect_done"] = not errors
    save_prog(tag, prog)


# ------------------------------------------------------------------ correo (streaming + puntos de control)
def clean_mail_filename(value, max_length=120):
    """Limpia nombres de asunto para que sean válidos en Windows."""
    value = str(value or "").strip()

    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", " ", value)
    value = value.rstrip(". ")

    # Evitar nombres reservados de Windows
    if value.upper() in {
        "CON", "PRN", "AUX", "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10))
    }:
        value = "_" + value

    return (value[:max_length].rstrip(". ") or "Sin asunto")


def get_mail_filename(message, raw):
    """Genera fecha_asunto.eml a partir del mensaje original."""

    try:
        subject = str(
            make_header(
                decode_header(message.get("Subject", "Sin asunto"))
            )
        )
    except Exception:
        subject = "Sin asunto"

    try:
        msg_date = email.utils.parsedate_to_datetime(
            message.get("Date", "")
        )

        if msg_date.tzinfo:
            msg_date = msg_date.astimezone().replace(tzinfo=None)

        date_text = msg_date.strftime("%Y-%m-%d")

    except Exception:
        date_text = "Sin-fecha"

    subject = clean_mail_filename(subject)

    # Huella para distinguir correos con el mismo asunto y fecha.
    message_id = message.get("Message-ID", "")
    fingerprint = hashlib.sha256(
        (message_id or "").encode("utf-8") + raw
    ).hexdigest()[:10]

    return f"{date_text}_{subject}_{fingerprint}.eml"


def get_mail_category(relative_path, filename):
    """Determina si un buzón corresponde a enviados o recibidos."""

    components = [
        str(x).lower().replace(".sbd", "")
        for x in relative_path.parts
    ]

    components.append(filename.lower())

    if any(
        component in SENT_MAIL
        for component in components
    ):
        return "Enviados"

    return "Recibidos"

def is_mbox(p: Path):
    if p.suffix.lower() in (".msf", ".dat", ".json", ".sqlite", ".html", ".txt", ".bak"):
        return False
    try:
        with open(p, "rb") as fh:
            return fh.read(5) == b"From "
    except Exception:
        return False


def _date_of(hdr, from_line):
    try:
        d = email.utils.parsedate_to_datetime(hdr.get("Date", ""))
        return d.astimezone().replace(tzinfo=None) if d.tzinfo else d
    except Exception:
        pass
    try:
        return datetime.strptime(from_line.decode("latin-1").split(None, 2)[2].strip(), "%a %b %d %H:%M:%S %Y")
    except Exception:
        return None


def _h(fh, pos, n):
    cur = fh.tell()
    fh.seek(max(0, pos - n))
    d = hashlib.sha1(fh.read(min(n, pos))).hexdigest()
    fh.seek(cur)
    return d


def filter_mbox(src, out_dir, ids_path, start, end,
                seen, st, save, stop):

    out_dir.mkdir(parents=True, exist_ok=True)
    ids_path.parent.mkdir(parents=True, exist_ok=True)

    n = 0
    new_ids = []
    buf = []

    with open(src, "rb") as fh:

        head = hashlib.sha1(fh.read(4096)).hexdigest()
        off = st.get("offset", 0)

        valid = (
            off > 0
            and st.get("version") == 2
            and st.get("head") == head
            and os.path.getsize(src) >= off
            and _h(fh, off, 128) == st.get("tail")
            and (
                st.get("ids_size", 0) == 0
                or (
                    ids_path.exists()
                    and ids_path.stat().st_size >= st["ids_size"]
        )
    )
)

        if valid:
            if ids_path.exists():
                with open(ids_path, "r+b") as f:
                    f.truncate(st.get("ids_size", 0))

                seen.update(
                    x.strip()
                    for x in ids_path.read_text(
                        encoding="utf-8"
                    ).splitlines()
                    if x.strip()
                )

            n = st.get("count", 0)

        else:
            off = 0
            ids_path.unlink(missing_ok=True)
            st.clear()
            n = 0

        def flush_ids():

            if new_ids:

                with open(ids_path, "a", encoding="utf-8") as f:

                    f.write("\n".join(new_ids) + "\n")
                    f.flush()
                    os.fsync(f.fileno())

                new_ids.clear()

        def checkpoint(boundary):
            flush_ids()

            st.update(
                version=2,
                offset=boundary,
                count=n,
                head=head,
                tail=_h(fh, boundary, 128),
                ids_size=(
                    ids_path.stat().st_size
                    if ids_path.exists() else 0
                )
            )

            save()

        def flush():

            nonlocal n

            if not buf:
                return

            # Quitar la línea separadora propia del formato MBOX.
            raw = b"".join(buf[1:])

            if not raw:
                return

            try:
                message = BytesParser(
                    policy=policy.default
                ).parsebytes(raw)

                # Respetar el filtro del mes que se respalda.
                d = _date_of(message, buf[0])

                if d is None or not (start <= d < end):
                    return

                # Evitar mensajes marcados como eliminados.
                try:
                    status = int(
                        message.get(
                            "X-Mozilla-Status", "0"
                        ).strip() or "0",
                        16
                    )

                    if status & 0x0008:
                        return

                except (ValueError, AttributeError):
                    pass

                # Evitar duplicados mediante Message-ID.
                mid = " ".join(
                    (message.get("Message-ID") or "").split()
                )

                if mid and mid in seen:
                    return

                # Si no hay Message-ID, utilizar huella del contenido.
                unique_id = mid or hashlib.sha256(raw).hexdigest()

                if unique_id in seen:
                    return

                filename = get_mail_filename(message, raw)
                output = out_dir / filename

                # Escritura temporal para evitar EML incompletos.
                part = output.with_name(output.name + ".part")

                with open(part, "wb") as destination:
                    generator = BytesGenerator(
                        destination,
                        policy=policy.default
                    )

                    generator.flatten(
                        message,
                        unixfrom=False
                    )

                    destination.flush()
                    os.fsync(destination.fileno())

                os.replace(part, output)

                seen.add(unique_id)
                new_ids.append(unique_id)
                n += 1

            except Exception as e:
                raise RuntimeError(
                    f"Error al convertir mensaje de {src}: {e}"
                ) from e

        try:

            fh.seek(off)

            pos = off
            prev_blank = True
            last_ck = time.time()
            message_since_checkpoint = 0

            
            while True:

                line = fh.readline()

                if not line:
                    break

                # Detectar el inicio de un nuevo correo MBOX.
                if prev_blank and line.startswith(b"From "):

                    # Procesar el correo anterior.
                    flush()

                    # Contabilizar un mensaje terminado.
                    messages_since_checkpoint += 1

                    # Guardar progreso cada 15 segundos
                    # o cada 25 mensajes procesados.
                    if (
                        time.time() - last_ck >= CK_SECONDS
                        or messages_since_checkpoint >= 25
                    ):

                        checkpoint(pos)

                        last_ck = time.time()
                        messages_since_checkpoint = 0

                        if stop():
                            raise Stopped()

                    # Comenzar el siguiente correo.
                    buf = [line]

                else:
                    buf.append(line)

                prev_blank = line.strip() == b""
                pos += len(line)


            flush()

            flush_ids()

            st.update(
                version=2,
                done=True,
                count=n,
                offset=pos,
                head=head,
                tail=_h(fh, pos, 128),
                ids_size=(
                    ids_path.stat().st_size
                    if ids_path.exists() else 0
                )
            )

            save()

        except Stopped:
            raise

    return n



def backup_mail(base, sub, out_base, start, end,
                prog, tag, log, err, stop):

    seen_cache = {}

    # Directorio independiente para identificadores de correos EML.
    ids_dir = PROG_DIR / f"{tag}_eml_ids"

    def ids_file(key):
        filename = hashlib.sha256(
        key.encode("utf-8")
    ).hexdigest()

        return ids_dir / f"{filename}.txt"


def seen_for(account_id):

    if account_id not in seen_cache:

        seen = set()

        prefix = f"EML/{account_id}/"

        for key, state in prog["mail"].items():

            # Solo consultar buzones de esta cuenta.
            if not key.startswith(prefix):
                continue

            path = ids_file(key)

            if not path.exists():
                continue

            # Recuperar únicamente los IDs confirmados
            # en el último checkpoint.
            saved_size = state.get("ids_size", 0)

            with open(path, "r+b") as f:
                f.truncate(saved_size)

            with open(path, "r", encoding="utf-8") as f:
                seen.update(
                    line.strip()
                    for line in f
                    if line.strip()
                )

        seen_cache[account_id] = seen

    return seen_cache[account_id]


    base = Path(base)

    for root, dirs, files in os.walk(base):

        if stop():
            raise Stopped()

        rel = Path(root).relative_to(base)

        dirs[:] = sorted(
            d for d in dirs
            if (
                d[:-4] if d.lower().endswith(".sbd") else d
            ).lower() not in SKIP_MAIL
        )

        if any(
            x.lower() in SKIP_MAIL
            for x in rel.parts
        ):
            continue

        for filename in sorted(files):

            p = Path(root) / filename

            if filename.lower() in SKIP_MAIL or not is_mbox(p):
                continue

            category = get_mail_category(rel, filename)

            account_id = hashlib.sha256(
                str(base.resolve()).lower().encode("utf-8")
            ).hexdigest()[:16]

            mailbox_id = hashlib.sha256(
                str(p.resolve()).lower().encode("utf-8")
            ).hexdigest()[:16]

            key = f"EML/{account_id}/{mailbox_id}"

            st = prog["mail"].setdefault(key, {})

            if st.get("done"):
                continue

            try:

                if (
                    not st.get("offset")
                    and p.stat().st_mtime < start.timestamp()
                ):
                    st.update(done=True, count=0, version=2)
                    save_prog(tag, prog)
                    continue

                destination = out_base / category

                log(
                    f"  {category}: {p.name}"
                    + (
                        " (retomando)"
                        if st.get("offset") else ""
                    )
                )

                count = filter_mbox(
                    p,
                    destination,
                    ids_file(key),
                    start,
                    end,
                    seen_for(account_id),
                    st,
                    lambda: save_prog(tag, prog),
                    stop
                )

                if count:
                    log(f"    {count} correos exportados")

            except Stopped:
                raise

            except Exception as e:
                err(f"Correo {key}: {e}")

    save_prog(tag, prog)


# ------------------------------------------------------------------ Credenciales de red para SYSTEM
def _host_from_dest(dest):
    """Extrae el host (nombre o IP) de una ruta UNC."""
    p = dest.strip()
    if p.startswith("\\\\"):
        return p[2:].split("\\")[0]
    return None


def save_net_credentials_for_system(host, user, password, log=None):
    """
    Guarda las credenciales en el almacén de Windows del usuario SYSTEM.
    La tarea programada corre como SYSTEM, que tiene su propio almacén de
    credenciales separado del usuario normal — por eso 'guardar con cmdkey'
    desde tu sesión normal NO sirve para el servicio.

    Técnica: crea una tarea temporal que corre cmdkey como SYSTEM y la
    ejecuta de inmediato. Borra la tarea al terminar.
    """
    if os.name != "nt":
        return True, ""
    if not host or not user:
        return False, "Host o usuario vacíos"
    task = "RespaldoCredTemp"
    cmd  = f'cmdkey /add:{host} /user:{user} /pass:{password}'
    try:
        # Crear tarea de un solo disparo como SYSTEM
        subprocess.run(
            ["schtasks", "/create", "/f", "/tn", task, "/sc", "once",
             "/st", "00:00", "/ru", "SYSTEM",
             "/tr", cmd],
            capture_output=True, check=True)
        # Ejecutar inmediatamente
        subprocess.run(["schtasks", "/run", "/tn", task], capture_output=True, check=True)
        # Esperar hasta 5 s a que termine
        for _ in range(10):
            time.sleep(0.5)
            r = subprocess.run(["schtasks", "/query", "/tn", task, "/fo", "csv"],
                               capture_output=True, text=True)
            if "Ready" in r.stdout or "Listo" in r.stdout:
                break
        # Borrar la tarea temporal
        subprocess.run(["schtasks", "/delete", "/f", "/tn", task], capture_output=True)
        if log:
            log(f"Credenciales de SYSTEM guardadas para {host} (usuario: {user})")
        return True, ""
    except subprocess.CalledProcessError as e:
        msg = f"No se pudieron guardar las credenciales: {e}"
        if log:
            log(f"  ⚠ {msg}")
        return False, msg


def mount_net_path(dest, user, password, log=None):
    """
    Establece la conexión de red a PC100 usando 'net use' antes de copiar.

    Maneja el error 1219 (conexión previa con distinto usuario) desconectando
    primero y luego reconectando con las credenciales correctas.
    """
    if os.name != "nt":
        return
    host = _host_from_dest(dest)
    if not host:
        return
    unc_base = f"\\\\{host}"

    def _connect():
        args = ["net", "use", unc_base]
        if user:
            args += [f"/user:{user}", password or ""]
        args += ["/persistent:no"]
        return subprocess.run(args, capture_output=True,
                              text=True, encoding="cp850", errors="replace", timeout=15)

    try:
        r = _connect()
        if r.returncode == 0:
            if log:
                log(f"Conexión con {unc_base} establecida correctamente")
            return
        out = (r.stdout + r.stderr).strip()
        # Error 1219: conexión previa con otro usuario → desconectar y reintentar
        if "1219" in out or "multiple" in out.lower() or "múltiples" in out.lower():
            if log:
                log(f"  Desconectando sesión previa en {unc_base} (error 1219)…")
            subprocess.run(["net", "use", unc_base, "/delete", "/yes"],
                           capture_output=True, timeout=10)
            time.sleep(1)
            r2 = _connect()
            if r2.returncode == 0:
                if log:
                    log(f"  Reconectado como {user or 'anónimo'}")
                return
            out = (r2.stdout + r2.stderr).strip()
        # Ya conectado sin problema (mensajes de éxito en español/inglés)
        if any(x in out.lower() for x in ("ya existe", "already", "se completó", "completed")):
            return
        if log and out:
            log(f"  net use: {out[:250]}")
    except Exception as e:
        if log:
            log(f"  net use falló: {e}")


# ------------------------------------------------------------------ FASE 2: enviar a PC100
def send(cfg, tag, log):
    src = local_dir(cfg, tag)
    if not src.exists():
        return True, None, ""
    why = "Sin destinos configurados"
    user     = cfg.get("net_user", "").strip()
    password = cfg.get("net_pass", "").strip()
    for dest in destinations(cfg):
        # Intentar establecer la conexión de red con credenciales antes de verificar acceso
        mount_net_path(dest, user, password, log)
        ok, why = dest_reachable(dest)
        if not ok:
            log(f"  {why}"); continue
        dst = Path(dest) / cfg["machine_name"].strip() / tag
        log(f"Enviando a {dst}")
        try:
            if shutil.which("robocopy"):
                # Escribir el log de robocopy en %TEMP% (siempre accesible, sea cual
                # sea el usuario que ejecuta la tarea: SYSTEM, admin o usuario normal)
                import tempfile
                rl = Path(tempfile.gettempdir()) / "respaldo_robocopy.log"
                # /Z = modo reiniciable: retoma archivos parciales si se corta la red
                with open(rl, "w", encoding="cp850", errors="replace") as lf:
                    rc = subprocess.call(
                        ["robocopy", str(src), str(dst),
                         "/E", "/COPY:DAT", "/R:3", "/W:10", "/Z",
                         "/NP", "/NDL", "/NFL", "/NJH"],
                        stdout=lf, stderr=lf)
                if rc >= 8:
                    try:
                        tail = rl.read_text(encoding="cp850", errors="replace")[-400:]
                    except Exception:
                        tail = "(no se pudo leer el log)"
                    why = f"robocopy código {rc}: {tail}"
                    log("  " + why); continue
            else:
                shutil.copytree(src, dst, dirs_exist_ok=True)
            return True, dest, ""
        except PermissionError as e:
            why = (f"Sin permisos en {dest}. "
                   f"Verifica que el usuario '{user or 'anónimo'}' tenga acceso de escritura "
                   f"a la carpeta compartida de PC100. Detalle: {e}")
            log("  " + why)
        except Exception as e:
            why = f"Envío a {dest}: {e}"; log("  " + why)
    return False, None, why


# ------------------------------------------------------------------ trabajo completo
def run_job(cfg, y, m, log, stop=lambda: False, collect_allowed=True):
    tag = f"{y}-{m:02d}"
    if not acquire_lock():
        log("Ya hay un respaldo en curso; esta ejecución se omite.")
        return None
    hb = start_heartbeat()
    try:
        prog = load_prog(tag)
        res = {"status": "local", "files": 0, "mails": 0, "errors": []}
        collected = prog["collect_done"]
        if not collected and collect_allowed:
            if prog["mail"] or prog["files_done"]:
                log(f"Retomando {tag}: {sum(1 for s in prog['mail'].values() if s.get('done'))} buzones ya completos")
            try:
                collect(cfg, y, m, prog, log, stop)
                collected = True
            except Stopped:
                save_prog(tag, prog)
                log("Detenido. Se retomará desde el último punto de control.")
                res["status"] = "stopped"; return res
        used = None
        if collected and not (prog["collect_done"] and prog["send_done"]):
            ok, used, why = send(cfg, tag, log)
            if ok and prog["collect_done"]:
                prog["send_done"] = True
            elif not ok:
                log("  Envío pendiente; se reintentará.")
        if prog["collect_done"] and prog["send_done"]:
            res["status"] = "ok"
        elif prog["errors"]:
            res["status"] = "error"
        res.update(files=prog["files"], mails=prog["mails"], errors=list(prog["errors"]))
        save_prog(tag, prog)
        if (y, m) == prev_month():
            cfg["last_status"], cfg["last_attempt"] = res["status"], date.today().isoformat()
            if res["status"] == "ok":
                cfg["last_done"] = tag
            save_cfg(cfg)
        if res["status"] in ("ok", "error"):
            for d in ([used] if used else destinations(cfg)):
                if dest_reachable(d)[0]:
                    write_status(cfg, tag, res, d); break
        if res["status"] == "ok" and cfg.get("delete_local_after_send"):
            shutil.rmtree(local_dir(cfg, tag), ignore_errors=True)
        log(f"== {tag}: {res['status'].upper()} (archivos={res['files']}, correos={res['mails']}, "
            f"errores={len(res['errors'])})")
        return res
    finally:
        hb.set()
        LOCK_FILE.unlink(missing_ok=True)


def my_minute_slot(name):
    """Minuto (0-54) de la hora que le toca a esta máquina, para no saturar la red con 120 equipos a la vez."""
    h = int(hashlib.sha1(name.encode("utf-8")).hexdigest(), 16)
    return h % 55


REQUEST_SUBDIR = "_solicitudes"


def request_file(dest, machine_name):
    return Path(dest) / REQUEST_SUBDIR / f"{machine_name.strip()}.json"


def check_backup_request(cfg, log):
    """Revisa si el panel del administrador pidió un respaldo inmediato para esta máquina.
    Se llama muy seguido (cada pocos minutos) y en el caso normal solo hace una prueba de
    existencia de archivo, así que es prácticamente gratis."""
    for dest in destinations(cfg):
        rf = request_file(dest, cfg["machine_name"])
        if not rf.exists():
            continue
        try:
            data = json.loads(rf.read_text(encoding="utf-8"))
        except Exception:
            data = {}
        try:
            rf.unlink()
        except Exception:
            pass
        mes = data.get("mes", "")
        if mes and len(mes) == 7 and mes[4] == "-":
            y, m = int(mes[:4]), int(mes[5:])
        else:
            y, m = prev_month()
        log(f"Solicitud recibida desde el panel del administrador: respaldar {y}-{m:02d} ahora.")
        run_job(cfg, y, m, log, collect_allowed=True)
        return True
    return False


def run_watch():
    """Modo liviano para la tarea que revisa cada pocos minutos si el administrador pidió
    un respaldo inmediato. No hace nada más (no cuenta días, no reparte horarios)."""
    CFG_DIR.mkdir(parents=True, exist_ok=True)

    def log(msg):
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")

    cfg = load_cfg()
    check_backup_request(cfg, log)


def run_auto(mode):
    CFG_DIR.mkdir(parents=True, exist_ok=True)
    lower_priority()

    def log(msg):
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")

    cfg = load_cfg()
    today = date.today()
    y, m = prev_month(today)
    tag = f"{y}-{m:02d}"
    if mode == "catchup":
        prog = load_prog(tag)
        started = prog["files_done"] or prog["collect_done"] or bool(prog["mail"])
        if not started:
            slot = my_minute_slot(cfg["machine_name"])
            if abs(datetime.now().minute - slot) > 2:
                return                                            # aún no es su turno en esta hora
        pend = [fp.stem for fp in PROG_DIR.glob("*.json")
                if fp.stem != tag and load_prog(fp.stem)["collect_done"] and not load_prog(fp.stem)["send_done"]]
        allow = today.day <= WINDOW_DAYS
        needs = cfg["last_done"] != tag and (allow or (prog["collect_done"] and not prog["send_done"]))
        if not needs and not pend:
            return
        if cfg["last_status"] == "error" and cfg["last_attempt"] == today.isoformat() and not pend:
            return
        for t in pend:
            run_job(cfg, int(t[:4]), int(t[5:]), log, collect_allowed=False)
        if needs:
            run_job(cfg, y, m, log, collect_allowed=allow)
    else:
        run_job(cfg, y, m, log)


def _principal_xml(cfg):
    return ("<UserId>SYSTEM</UserId><RunLevel>HighestAvailable</RunLevel>" if cfg.get("run_as_system", True)
            else "<RunLevel>HighestAvailable</RunLevel>")


def _exe_cmd_args(extra_arg):
    if getattr(sys, "frozen", False):
        return sys.executable, extra_arg
    py = sys.executable.replace("python.exe", "pythonw.exe")
    return (py if os.path.exists(py) else sys.executable), f'"{os.path.abspath(__file__)}" {extra_arg}'


def schedule_task(cfg):
    """Registra dos tareas:
    - La principal (TASK_NAME): hace el respaldo del mes, repartido en horario, del día 1 al 9.
    - La de vigilancia (TASK_NAME + 'Watch'): cada 5 minutos, solo para reaccionar rápido si el
      administrador pide un respaldo inmediato desde el panel. Revisar un archivo es casi gratis,
      así que correr tan seguido no sobrecarga la laptop."""
    cmd, args = _exe_cmd_args("--catchup")
    principal = _principal_xml(cfg)
    xml_main = f'''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
 <Triggers>
  <CalendarTrigger>
   <Repetition><Interval>PT1H</Interval><Duration>PT10H</Duration><StopAtDurationEnd>false</StopAtDurationEnd></Repetition>
   <StartBoundary>{datetime.now():%Y-%m-%d}T09:00:00</StartBoundary>
   <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
  </CalendarTrigger>
  <BootTrigger><Enabled>true</Enabled></BootTrigger>
 </Triggers>
 <Principals><Principal id="Author">{principal}</Principal></Principals>
 <Settings>
  <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
  <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
  <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
  <StartWhenAvailable>true</StartWhenAvailable>
  <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
  <Hidden>true</Hidden>
 </Settings>
 <Actions Context="Author"><Exec><Command>{cmd}</Command><Arguments>{args}</Arguments></Exec></Actions>
</Task>'''
    cmd_w, args_w = _exe_cmd_args("--vigilar")
    xml_watch = f'''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
 <Triggers>
  <TimeTrigger>
   <Repetition><Interval>PT5M</Interval></Repetition>
   <StartBoundary>{datetime.now():%Y-%m-%d}T00:00:00</StartBoundary>
  </TimeTrigger>
  <BootTrigger><Enabled>true</Enabled></BootTrigger>
 </Triggers>
 <Principals><Principal id="Author">{principal}</Principal></Principals>
 <Settings>
  <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
  <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
  <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
  <StartWhenAvailable>true</StartWhenAvailable>
  <ExecutionTimeLimit>PT4M</ExecutionTimeLimit>
  <Hidden>true</Hidden>
 </Settings>
 <Actions Context="Author"><Exec><Command>{cmd_w}</Command><Arguments>{args_w}</Arguments></Exec></Actions>
</Task>'''
    CFG_DIR.mkdir(parents=True, exist_ok=True)
    out = []
    for name, xml in ((TASK_NAME, xml_main), (TASK_NAME + "Watch", xml_watch)):
        fp = CFG_DIR / f"{name}.xml"
        fp.write_text(xml, encoding="utf-16")
        r = subprocess.run(["schtasks", "/create", "/f", "/tn", name, "/xml", str(fp)], capture_output=True, text=True)
        out.append((r.returncode, (r.stdout + r.stderr).strip()))
    code = 0 if all(c == 0 for c, _ in out) else 1
    return code, " | ".join(o for _, o in out)


# ------------------------------------------------------------------ instalación silenciosa (--instalar)
def install_silent(argv):
    """python respaldo_mensual.py --instalar --maquina "LAPC 46" --destino \\\\192.168.1.100\\Respaldos
       [--destino-alt \\\\PC100\\Respaldos] [--usuario-objetivo "ULAPC 46"] [--sin-system]"""
    cfg = load_cfg()
    flags = {"--maquina": "machine_name", "--destino": "destination", "--destino-alt": "destination_alt",
             "--local-root": "local_root", "--usuario-objetivo": "target_user",
             "--usuario-red": "net_user", "--password-red": "net_pass"}
    for i, a in enumerate(argv):
        if a in flags and i + 1 < len(argv):
            cfg[flags[a]] = argv[i + 1]
    if "--sin-system" in argv:
        cfg["run_as_system"] = False
    save_cfg(cfg)
    hide_path(Path(cfg["local_root"]))
    hide_path(CFG_DIR)
    # Guardar credenciales de red para la cuenta SYSTEM
    for dest in destinations(cfg):
        host = _host_from_dest(dest)
        if host and cfg.get("net_user"):
            ok, msg = save_net_credentials_for_system(host, cfg["net_user"], cfg["net_pass"])
            print(f"credenciales {host}: {'OK' if ok else msg}")
    code, out = schedule_task(cfg)
    print(f"maquina={cfg['machine_name']} schtasks_code={code} {out}")
    return 0 if code == 0 else 1


# ------------------------------------------------------------------ GUI
# Paleta de colores simple y amigable
C_BG = "#F2F6FC"          # fondo general, azul muy claro
C_CARD = "#FFFFFF"        # tarjetas blancas
C_PRIMARY = "#2563EB"     # azul: acción principal
C_PRIMARY_DK = "#1D4ED8"
C_GREEN = "#16A34A"       # éxito
C_GREEN_BG = "#DCFCE7"
C_ORANGE = "#D97706"      # pendiente / aviso
C_ORANGE_BG = "#FEF3C7"
C_RED = "#DC2626"         # error
C_RED_BG = "#FEE2E2"
C_TEXT = "#1F2937"
C_MUTED = "#6B7280"
FONT = ("Segoe UI", 10)
FONT_B = ("Segoe UI", 10, "bold")
FONT_H = ("Segoe UI", 15, "bold")


def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    cfg = load_cfg()
    root = tk.Tk()
    root.title("Respaldo Mensual")
    root.geometry("1060x740")
    root.configure(bg=C_BG)
    root.minsize(980, 600)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(".", background=C_BG, foreground=C_TEXT, font=FONT)
    style.configure("TFrame", background=C_BG)
    style.configure("Card.TFrame", background=C_CARD)
    style.configure("TLabel", background=C_BG, foreground=C_TEXT, font=FONT)
    style.configure("Card.TLabel", background=C_CARD, foreground=C_TEXT, font=FONT)
    style.configure("Muted.TLabel", background=C_BG, foreground=C_MUTED, font=FONT)
    style.configure("CardMuted.TLabel", background=C_CARD, foreground=C_MUTED, font=FONT)
    style.configure("Header.TLabel", background=C_BG, foreground=C_PRIMARY_DK, font=FONT_H)
    style.configure("CardTitle.TLabel", background=C_CARD, foreground=C_PRIMARY_DK, font=FONT_B)
    style.configure("TLabelframe", background=C_CARD, bordercolor="#D7E2F2")
    style.configure("TLabelframe.Label", background=C_CARD, foreground=C_PRIMARY_DK, font=FONT_B)
    style.configure("TEntry", padding=5, fieldbackground="#FFFFFF")
    style.configure("TSpinbox", padding=4, fieldbackground="#FFFFFF")
    style.configure("TCheckbutton", background=C_CARD, font=FONT)
    style.configure("TNotebook", background=C_BG, borderwidth=0)
    style.configure("TNotebook.Tab", font=FONT_B, padding=[16, 8])

    style.configure("TButton", font=FONT, padding=6)
    style.configure("Primary.TButton", font=FONT_B, padding=9, background=C_PRIMARY, foreground="white")
    style.map("Primary.TButton", background=[("active", C_PRIMARY_DK)])
    style.configure("Soft.TButton", font=FONT, padding=6, background="#E5EDFB", foreground=C_PRIMARY_DK)
    style.map("Soft.TButton", background=[("active", "#D3E1F8")])
    style.configure("Danger.TButton", font=FONT, padding=6, background=C_RED_BG, foreground=C_RED)
    style.map("Danger.TButton", background=[("active", "#FACACA")])
    style.configure("Link.TButton", font=("Segoe UI", 9), padding=2, background=C_BG, foreground=C_MUTED,
                     borderwidth=0)

    q = queue.Queue()
    cancel = {"v": False}

    # ---------- encabezado ----------
    header = ttk.Frame(root); header.pack(fill="x", padx=16, pady=(14, 4))
    ttk.Label(header, text="🗂️  Respaldo Mensual", style="Header.TLabel").pack(side="left")
    lbl_badge = tk.Label(header, text="", font=FONT_B, padx=12, pady=4)
    lbl_badge.pack(side="right")

    body = ttk.Frame(root); body.pack(fill="both", expand=True, padx=16, pady=6)

    canvas = tk.Canvas(body, bg=C_BG, highlightthickness=0)
    vsb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
    f = ttk.Frame(canvas)
    f.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=f, anchor="nw", width=760)
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    def _wheel(e):
        canvas.yview_scroll(int(-e.delta / 120), "units")
    canvas.bind_all("<MouseWheel>", _wheel)

    V = {k: tk.StringVar(value=cfg.get(k, "")) for k in
         ("machine_name", "destination", "destination_alt", "local_root", "thunderbird", "target_user",
          "net_user", "net_pass")}
    v_excl = tk.StringVar(value=", ".join(cfg.get("excluded_users", [])))
    v_del = tk.BooleanVar(value=cfg["delete_local_after_send"])
    v_sys = tk.BooleanVar(value=cfg.get("run_as_system", True))
    py_, pm_ = prev_month()
    v_year, v_month = tk.IntVar(value=py_), tk.IntVar(value=pm_)

    def card(parent, title):
        lf = ttk.Labelframe(parent, text=title, padding=14)
        lf.pack(fill="x", pady=8, padx=2)
        lf.columnconfigure(1, weight=1)
        return lf

    def field(parent, r, label, key, browse=False, help_=None):
        ttk.Label(parent, text=label, style="Card.TLabel").grid(row=r, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Entry(parent, textvariable=V[key]).grid(row=r, column=1, sticky="ew", pady=4)
        if browse:
            ttk.Button(parent, text="📁 Elegir…", style="Soft.TButton",
                       command=lambda: V[key].set(filedialog.askdirectory() or V[key].get())
                       ).grid(row=r, column=2, padx=6)
        if help_:
            ttk.Label(parent, text=help_, style="CardMuted.TLabel").grid(
                row=r + 1, column=0, columnspan=3, sticky="w", pady=(0, 6))

    # ---- tarjeta: esta máquina ----
    c1 = card(f, "💻  Esta máquina")
    field(c1, 0, "Nombre con el que se guarda en PC100:", "machine_name",
          help_="Ej. «LAPC 46». Así se llamará la carpeta de esta máquina en el servidor.")
    lbl_users = ttk.Label(c1, text="", style="CardMuted.TLabel")
    lbl_users.grid(row=2, column=0, columnspan=3, sticky="w")
    lbl_tb = ttk.Label(c1, text="", style="CardMuted.TLabel", wraplength=680, justify="left")
    lbl_tb.grid(row=3, column=0, columnspan=3, sticky="w", pady=(0, 4))

    # ---- tarjeta: a dónde enviar + credenciales ----
    c2 = card(f, "🌐  A dónde se envía (PC100)")
    c2.columnconfigure(1, weight=1)

    field(c2, 0, "Destino principal (recomendado: IP de PC100):", "destination", True)
    field(c2, 1, "Destino de respaldo (nombre de PC100):", "destination_alt")

    # Separador visual
    ttk.Separator(c2, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)
    ttk.Label(c2, text="🔑  Credenciales de la carpeta compartida", style="CardTitle.TLabel"
              ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 6))
    ttk.Label(c2, text="Usuario (ej. PC100\\respaldo o solo respaldo):", style="Card.TLabel"
              ).grid(row=5, column=0, sticky="w", pady=4, padx=(0, 8))
    ttk.Entry(c2, textvariable=V["net_user"]).grid(row=5, column=1, sticky="ew", pady=4)
    ttk.Label(c2, text="Contraseña:", style="Card.TLabel"
              ).grid(row=6, column=0, sticky="w", pady=4, padx=(0, 8))
    e_pass = ttk.Entry(c2, textvariable=V["net_pass"], show="●")
    e_pass.grid(row=6, column=1, sticky="ew", pady=4)
    v_show = tk.BooleanVar(value=False)
    def toggle_pass():
        e_pass.config(show="" if v_show.get() else "●")
    ttk.Checkbutton(c2, text="Mostrar", variable=v_show, command=toggle_pass,
                    style="TCheckbutton").grid(row=6, column=2, padx=6)
    ttk.Label(c2,
              text="Si la carpeta de PC100 requiere usuario y contraseña, ponlos aquí.\n"
                   "El botón 'Guardar credenciales en SYSTEM' los registra para que la\n"
                   "tarea programada (que corre en segundo plano) también tenga acceso.",
              style="CardMuted.TLabel", justify="left"
              ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(0, 4))

    lbl_cred_status = ttk.Label(c2, text="", style="CardMuted.TLabel")
    lbl_cred_status.grid(row=8, column=0, columnspan=3, sticky="w")

    def do_guardar_cred():
        c = current_cfg()
        save_cfg(c)
        u = c.get("net_user", "").strip()
        p = c.get("net_pass", "").strip()
        if not u:
            lbl_cred_status.config(text="⚠ Escribe el usuario antes de guardar.", foreground=C_ORANGE)
            return
        lbl_cred_status.config(text="Guardando credenciales en SYSTEM…", foreground=C_MUTED)
        def work():
            msgs = []
            for dest in destinations(c):
                host = _host_from_dest(dest)
                if host:
                    # 1. Guardar en el almacén de SYSTEM (para la tarea programada)
                    ok, msg = save_net_credentials_for_system(host, u, p, log)
                    msgs.append(f"{host}: {'✔ guardado en SYSTEM' if ok else '⚠ ' + msg}")
                    # 2. Guardar también para el usuario actual (para cuando se usa la GUI)
                    subprocess.run(["cmdkey", f"/add:{host}", f"/user:{u}", f"/pass:{p}"],
                                   capture_output=True)
                    # 3. Verificar que la conexión funciona ahora
                    mount_net_path(dest, u, p, log)
                    ok2, why = dest_reachable(dest)
                    msgs.append(f"  → Prueba de conexión: {'✔ OK' if ok2 else '✗ ' + why}")
            resumen = " | ".join(msgs)
            root.after(0, lambda: lbl_cred_status.config(
                text=resumen[:160],
                foreground=C_GREEN if all("✔" in m for m in msgs if "→" in m) else C_RED))
        threading.Thread(target=work, daemon=True).start()

    ttk.Button(c2, text="🔑 Guardar credenciales en SYSTEM", style="Soft.TButton",
               command=do_guardar_cred).grid(row=9, column=0, columnspan=2, sticky="w", pady=(4, 0))

    # ---- tarjeta: mes ----
    c3 = card(f, "📅  Mes a respaldar")
    mf = ttk.Frame(c3, style="Card.TFrame"); mf.grid(row=0, column=0, columnspan=3, sticky="w")
    ttk.Spinbox(mf, from_=2000, to=2100, width=6, textvariable=v_year).pack(side="left")
    ttk.Spinbox(mf, from_=1, to=12, width=4, textvariable=v_month).pack(side="left", padx=6)
    ttk.Label(mf, text="(por defecto, el mes anterior)", style="CardMuted.TLabel").pack(side="left", padx=6)

    # ---- tarjeta: carpetas extra ----
    c4 = card(f, "➕  Carpetas extra (ej. escáner compartido)")
    lb = tk.Listbox(c4, height=4, font=FONT, relief="flat", bg="#F8FAFD", highlightthickness=1,
                     highlightbackground="#D7E2F2")
    lb.grid(row=0, column=0, columnspan=2, sticky="ew")
    for s in cfg["sources"]:
        lb.insert("end", s)
    bf = ttk.Frame(c4, style="Card.TFrame"); bf.grid(row=0, column=2, sticky="n", padx=6)
    ttk.Button(bf, text="Agregar", style="Soft.TButton",
               command=lambda: (lambda d: d and lb.insert("end", d))(filedialog.askdirectory())).pack(fill="x", pady=2)
    ttk.Button(bf, text="Quitar", style="Soft.TButton",
               command=lambda: [lb.delete(i) for i in reversed(lb.curselection())]).pack(fill="x", pady=2)

    # ---- tarjeta: opciones avanzadas (plegable) ----
    adv_shown = tk.BooleanVar(value=False)
    adv_holder = ttk.Frame(f)

    def toggle_adv():
        if adv_shown.get():
            adv_holder.pack_forget(); adv_shown.set(False); btn_adv.config(text="⚙️  Mostrar opciones avanzadas ▾")
        else:
            adv_holder.pack(fill="x", pady=2, before=c6); adv_shown.set(True)
            btn_adv.config(text="⚙️  Ocultar opciones avanzadas ▴")

    btn_adv = ttk.Button(f, text="⚙️  Mostrar opciones avanzadas ▾", style="Soft.TButton", command=toggle_adv)
    btn_adv.pack(fill="x", pady=(10, 2))

    c5 = ttk.Labelframe(adv_holder, text="⚙️  Opciones avanzadas", padding=14)
    c5.pack(fill="x", pady=8); c5.columnconfigure(1, weight=1)
    field(c5, 0, "Usuario objetivo (vacío = detectar solo):", "target_user",
          help_="Déjalo vacío salvo que haya más de un usuario y debas precisar cuál, ej. «ULAPC».")
    ttk.Label(c5, text="Usuarios excluidos (coma):", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=4)
    ttk.Entry(c5, textvariable=v_excl).grid(row=2, column=1, sticky="ew", pady=4)
    field(c5, 3, "Carpeta local temporal (oculta):", "local_root", True)
    field(c5, 5, "Perfil de Thunderbird manual (opcional):", "thunderbird", True)

    def do_buscar_tb():
        """Busca perfiles de Thunderbird en los discos y los muestra en el log para copiar la ruta."""
        log("🔍 Buscando perfiles de Thunderbird en el disco…")
        def buscar():
            encontrados = []
            drives = [f"{chr(d)}:\\" for d in range(ord('C'), ord('F') + 1)
                      if os.path.exists(f"{chr(d)}:\\")]
            skip = {'windows', 'program files', 'program files (x86)',
                    '$recycle.bin', 'programdata', 'system volume information'}
            for drive in drives:
                for rdir, dirs, files in os.walk(drive):
                    dirs[:] = [d for d in dirs if d.lower() not in skip]
                    if 'prefs.js' in files and 'thunderbird' in rdir.lower():
                        encontrados.append(rdir)
                        log(f"  ✔ Encontrado: {rdir}")
                    if len(encontrados) >= 20:
                        break
            if not encontrados:
                log("  No se encontró ningún perfil. ¿Thunderbird está instalado en esta PC?")
            else:
                log(f"  → Copia la ruta correcta en el campo 'Perfil de Thunderbird manual' de arriba.")
        threading.Thread(target=buscar, daemon=True).start()

    ttk.Button(c5, text="🔍 Buscar Thunderbird en el disco", style="Soft.TButton",
               command=do_buscar_tb).grid(row=7, column=0, columnspan=2, sticky="w", pady=(4, 0))
    ttk.Checkbutton(c5, text="Borrar la copia local cuando se envíe correctamente", variable=v_del
                    ).grid(row=7, column=0, columnspan=2, sticky="w", pady=4)
    ttk.Checkbutton(c5, text="Ejecutar la tarea programada en segundo plano (oculta, recomendado)", variable=v_sys
                    ).grid(row=8, column=0, columnspan=2, sticky="w", pady=4)

    # ---- registro de actividad ----
    c6 = card(f, "📋  Actividad")
    bar = ttk.Progressbar(c6, mode="indeterminate"); bar.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
    logbox = tk.Text(c6, height=10, font=("Consolas", 9), bg="#F8FAFD", fg=C_TEXT, relief="flat",
                      highlightthickness=1, highlightbackground="#D7E2F2")
    logbox.grid(row=1, column=0, columnspan=3, sticky="ew")

    def log(msg):
        q.put(msg)
        CFG_DIR.mkdir(parents=True, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")

    def pump():
        try:
            while True:
                s = q.get_nowait()
                logbox.config(state="normal"); logbox.insert("end", s + "\n"); logbox.see("end")
                logbox.config(state="disabled")
        except queue.Empty:
            pass
        root.after(200, pump)

    def current_cfg():
        c = load_cfg()
        c.update({k: v.get().strip() for k, v in V.items()})
        c.update(sources=list(lb.get(0, "end")), delete_local_after_send=v_del.get(),
                 run_as_system=v_sys.get(),
                 excluded_users=[x.strip() for x in v_excl.get().split(",") if x.strip()])
        return c

    def refresh_users():
        c = current_cfg()
        users = pick_target_users(c)
        if users:
            lbl_users.config(text="✔  Usuario detectado: " + ", ".join(u.name for u in users),
                             foreground=C_GREEN)
        else:
            lbl_users.config(text="⚠  No se detecta ningún usuario (solo administradores en esta PC)",
                             foreground=C_RED)
        # Estado de Thunderbird
        kind, msg = thunderbird_summary(c, users)
        color = C_GREEN if kind == "ok" else (C_RED if kind == "error" else C_ORANGE)
        lbl_tb.config(text="📧 " + msg, foreground=color)

    def refresh_state():
        c = load_cfg(); y, m = prev_month(); t = f"{y}-{m:02d}"; p = load_prog(t)
        if c["last_done"] == t:
            lbl_badge.config(text=f"✔  {t}: AL DÍA", bg=C_GREEN_BG, fg=C_GREEN)
        elif p["collect_done"]:
            lbl_badge.config(text=f"⏳  {t}: FALTA ENVIAR", bg=C_ORANGE_BG, fg=C_ORANGE)
        elif p["files_done"] or p["mail"]:
            lbl_badge.config(text=f"↻  {t}: EN PROCESO", bg=C_ORANGE_BG, fg=C_ORANGE)
        else:
            lbl_badge.config(text=f"⚠  {t}: PENDIENTE", bg=C_RED_BG, fg=C_RED)

    def do_save(silent=False):
        save_cfg(current_cfg())
        if not silent:
            log("✔ Configuración guardada.")
        refresh_users()

    def _set_running(running: bool):
        """Cambia el aspecto de la ventana para que sea imposible no ver que está respaldando."""
        if running:
            root.title("⏳ Respaldando… — Respaldo Mensual")
            lbl_badge.config(text="⏳  RESPALDANDO AHORA…", bg=C_ORANGE_BG, fg=C_ORANGE)
            b_run.config(state="disabled", text="⏳  Respaldando…")
            b_reset.config(state="disabled")
            bar.start(8)
            # Hacer el log visible sin que el usuario tenga que hacer scroll
            canvas.yview_moveto(1.0)
        else:
            root.title("Respaldo Mensual")
            b_run.config(state="normal", text="▶  Respaldar ahora")
            b_reset.config(state="normal")
            bar.stop()
            refresh_state()

    def _run(y, m, allow=True):
        c = current_cfg(); save_cfg(c); cancel["v"] = False
        _set_running(True)
        log(f"{'─'*60}")
        log(f"▶  Iniciando respaldo {y}-{m:02d}…")
        log(f"{'─'*60}")

        def work():
            run_job(c, y, m, log, lambda: cancel["v"], collect_allowed=allow)
            root.after(0, lambda: _set_running(False))
        threading.Thread(target=work, daemon=True).start()

    def do_run():
        do_save(silent=True); _run(v_year.get(), v_month.get())

    def do_reset():
        tag = f"{v_year.get()}-{v_month.get():02d}"
        if messagebox.askyesno("Empezar de nuevo",
                                f"Esto borra lo recopilado localmente de {tag} y vuelve a empezar. ¿Continuar?"):
            reset_month(current_cfg(), tag); log(f"↻ Reiniciado {tag}."); do_run()

    def do_test():
        log("— Probando conexión —")
        threading.Thread(target=lambda: test_destinations(current_cfg(), log), daemon=True).start()

    def do_schedule():
        do_save(silent=True)
        if os.name != "nt":
            messagebox.showinfo("Programar", "Esto requiere el Programador de tareas de Windows."); return
        code, out = schedule_task(current_cfg()); log(f"Tarea programada ({code}): {out}")
        if code == 0:
            messagebox.showinfo("Programar", "Listo. El respaldo se repetirá solo cada mes, sin que tengas "
                                 "que abrir este programa.")
        else:
            messagebox.showerror("Programar", f"No se pudo crear la tarea (¿ejecutas como administrador?):\n{out}")

    # ---- barra de acciones ----
    actions = ttk.Frame(root, style="TFrame"); actions.pack(fill="x", padx=16, pady=(4, 10))
    b_run = ttk.Button(actions, text="▶  Respaldar ahora", style="Primary.TButton", command=do_run)
    b_run.pack(side="left", padx=(0, 8))
    b_reset = ttk.Button(actions, text="↻  Empezar de nuevo", style="Soft.TButton", command=do_reset)
    b_reset.pack(side="left", padx=4)
    ttk.Button(actions, text="⏹  Detener", style="Danger.TButton", command=lambda: cancel.update(v=True)
               ).pack(side="left", padx=4)
    ttk.Button(actions, text="📶  Probar conexión", style="Soft.TButton", command=do_test).pack(side="left", padx=4)
    ttk.Button(actions, text="🗓  Programar automático", style="Soft.TButton", command=do_schedule
               ).pack(side="left", padx=4)
    ttk.Button(actions, text="💾  Guardar", style="Soft.TButton", command=do_save).pack(side="left", padx=4)

    footer = ttk.Frame(root); footer.pack(fill="x", padx=16, pady=(0, 10))
    ttk.Label(footer, text="El estado de todas las máquinas se revisa desde el panel del encargado en PC100.",
              style="Muted.TLabel").pack(side="left")

    refresh_state(); refresh_users(); pump()
    root.mainloop()


def diagnostico_thunderbird():
    users = get_user_profiles()

    for user in users:
        profiles = find_thunderbird_profiles(user)

        for profile in profiles:
            print("\nPERFIL:", profile)

            accounts = get_thunderbird_accounts(profile)

            for account in accounts:
                print("CUENTA:", account["label"])
                print("TIPO:", account["type"])
                print("ALMACEN:", account["directory"])
                print("EXISTE:", account["directory"].is_dir())

            prefs = read_thunderbird_prefs(profile)

            print("\nPREFERENCIAS DE CARPETAS:")

            for key, value in prefs.items():
                if any(x in key.lower() for x in (
                    "fcc_folder",
                    "sent_folder",
                    "trash_folder",
                    "draft_folder",
                    "imapserver",
                    "directory"
                )):
                    print(key, "=", value)

def main():
    a = sys.argv[1:]
    if "--instalar" in a:
        sys.exit(install_silent(a))
    elif "--catchup" in a:
        run_auto("catchup")
    elif "--auto" in a:
        run_auto("auto")
    elif "--vigilar" in a:
        run_watch()
    elif "--diagnostico-tb" in a:
        diagnostico_thunderbird()    
    else:
        run_gui()


if __name__ == "__main__":
    main()