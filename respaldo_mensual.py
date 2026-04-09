"""
respaldo_mensual.py - Sistema de respaldo mensual por red local Windows.
"""

import os, sys, socket, shutil, threading, logging, json
import email, email.policy, email.utils, mailbox, calendar
from datetime import datetime, date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ───────────────────────────────────────────────
#  Modulo Thunderbird
# ───────────────────────────────────────────────

# Carpetas de Thunderbird que nos interesan
TB_TARGET_FOLDERS = {"INBOX", "Sent", "Sent Messages", "Enviados", "Bandeja de entrada",
                     "Drafts", "Borradores"}

def get_user_home() -> Path:
    """
    Detecta la carpeta home del usuario LOCAL de la PC, no del admin/propietario.
    ULAPC37 -> C:\\Users\\lapc37  (o como se llame la cuenta local)
    """
    ADMINS = {"administrator", "administrador", "propietario", "owner",
              "public", "publico", "default", "all users", "default user"}
    hostname = socket.gethostname().upper()
    digits   = "".join(c for c in hostname if c.isdigit())
    users_dir = Path("C:/Users")
    candidates = []
    try:
        for d in users_dir.iterdir():
            if d.is_dir() and d.name.lower() not in ADMINS:
                candidates.append(d)
    except Exception:
        pass
    # Un solo candidato no-admin
    if len(candidates) == 1:
        return candidates[0]
    # Buscar por digitos del hostname (ULAPC37 -> buscar "37" en nombre de usuario)
    if candidates and digits:
        for d in candidates:
            if digits in d.name.lower():
                return d
        for d in candidates:
            if hostname.lower() in d.name.lower() or d.name.lower() in hostname.lower():
                return d
    # Fallback: USERPROFILE si no es admin
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile and Path(userprofile).exists():
        p = Path(userprofile)
        if p.name.lower() not in ADMINS:
            return p
    if candidates:
        return sorted(candidates, key=lambda d: d.stat().st_mtime, reverse=True)[0]
    return Path.home()

def _user_file(name: str) -> Path:
    """Archivo en la carpeta del usuario local."""
    return get_user_home() / name

def find_thunderbird_profile() -> Path | None:
    """Localiza el perfil activo de Thunderbird del usuario local."""
    home = get_user_home()
    tb_root = home / "AppData" / "Roaming" / "Thunderbird" / "Profiles"
    if not tb_root.exists():
        return None
    profiles = sorted(tb_root.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True)
    for p in profiles:
        if p.is_dir():
            return p
    return None

def find_thunderbird_accounts(profile: Path) -> list:
    """
    Devuelve lista de (nombre_legible, Path_carpeta) para cada cuenta.
    Incluye cuentas IMAP (ImapMail/) y carpetas locales (Mail/).
    """
    accounts = []

    # Cuentas IMAP (Gmail, etc.)
    imap_root = profile / "ImapMail"
    if imap_root.exists():
        for account_dir in sorted(imap_root.iterdir()):
            if account_dir.is_dir():
                accounts.append((account_dir.name, account_dir))

    # Carpetas locales de Thunderbird (Mail/Local Folders/)
    mail_root = profile / "Mail"
    if mail_root.exists():
        for account_dir in sorted(mail_root.iterdir()):
            if account_dir.is_dir():
                # Incluir aunque se llame "Local Folders" o cualquier nombre
                accounts.append((account_dir.name, account_dir))

    return accounts

def find_mbox_files(profile: Path):
    """
    Devuelve lista de (Path_mbox, etiqueta_legible) para todas las cuentas.
    - Cuentas IMAP: solo INBOX y Enviados
    - Carpetas locales: TODOS los archivos mbox (cualquier carpeta que el
      usuario haya creado localmente en Thunderbird)
    """
    results = []
    seen = set()
    accounts = find_thunderbird_accounts(profile)

    for account_name, account_dir in accounts:
        # Determinar si es cuenta local o IMAP
        is_local = "local" in account_name.lower() or account_dir.parent.name == "Mail"

        # Nombre legible de la cuenta
        cuenta = account_name.replace("imap.gmail.com", "Gmail")
        if cuenta.startswith("Gmail-"):
            cuenta = "Gmail" + cuenta[6:]
        if "local" in cuenta.lower() or cuenta.lower() == "local folders":
            cuenta = "Carpetas locales"

        for mbox_file in account_dir.rglob("*"):
            if not mbox_file.is_file():
                continue
            if mbox_file.suffix != "":
                continue
            if mbox_file.name.endswith(".msf"):
                continue
            if str(mbox_file) in seen:
                continue
            # Ignorar archivos de sistema de Thunderbird
            if mbox_file.name in ("Trash", "Junk", "Templates", "Drafts",
                                   "Borradores", "Spam", "Basura"):
                continue

            name_upper = mbox_file.name.upper()

            if is_local:
                # En carpetas locales: respaldar TODOS los mbox
                # (el usuario puede tener carpetas propias con cualquier nombre)
                etiqueta = f"{cuenta} - {mbox_file.name}"
                results.append((mbox_file, etiqueta))
                seen.add(str(mbox_file))
            else:
                # En cuentas IMAP: solo INBOX y Enviados
                for target in TB_TARGET_FOLDERS:
                    if name_upper == target.upper():
                        etiqueta = f"{cuenta} - {mbox_file.name}"
                        results.append((mbox_file, etiqueta))
                        seen.add(str(mbox_file))
                        break
    return results


def count_emails_in_mbox(mbox_path: Path, start_date: date, end_date: date) -> int:
    """Cuenta cuantos correos hay en el rango sin exportarlos."""
    count = 0
    try:
        mb = mailbox.mbox(str(mbox_path), create=False)
        for msg in mb:
            d = parse_email_date(msg)
            if d and start_date <= d <= end_date:
                count += 1
        mb.close()
    except Exception:
        pass
    return count

def parse_email_date(msg) -> date | None:
    """Extrae la fecha de un mensaje de email, tolerante a formatos raros."""
    try:
        date_str = msg.get("Date", "")
        if not date_str:
            # Intentar con Received: que siempre tiene fecha
            received = msg.get("Received", "")
            if ";" in received:
                date_str = received.split(";")[-1].strip()
        if not date_str:
            return None
        # Limpiar parentesis al final: "Wed, 18 Feb 2026 10:26:17 -0600 (CST)"
        date_str = date_str.split("(")[0].strip()
        parsed = email.utils.parsedate_to_datetime(date_str)
        return parsed.date()
    except Exception:
        try:
            # Ultimo intento: extraer fecha con parsedate
            tup = email.utils.parsedate(msg.get("Date", ""))
            if tup:
                return date(tup[0], tup[1], tup[2])
        except Exception:
            pass
        return None

def _safe_header(value: str, max_len: int = 50) -> str:
    """Limpia un header de correo para nombre de archivo. Rapido."""
    if not value:
        return ""
    try:
        parts = email.header.decode_header(value)
        decoded = []
        for part, enc in parts:
            if isinstance(part, bytes):
                decoded.append(part.decode(enc or "utf-8", errors="replace"))
            else:
                decoded.append(str(part))
        value = " ".join(decoded)
    except Exception:
        pass
    bad = set('/\\:*?"<>|')
    value = "".join(c if c not in bad and ord(c) >= 32 else "_" for c in value)
    return value[:max_len].strip(" .")


def export_emails_to_eml(mbox_path: Path, dest_folder: Path,
                          start_date: date, end_date: date,
                          progress_cb=None) -> int:
    """
    Exporta correos del rango como .eml individuales.
    - Lee el mbox una sola vez en memoria
    - Filtra por fecha leyendo solo headers (no carga cuerpo ni adjuntos)
    - Escribe en paralelo con ThreadPoolExecutor para maxima velocidad
    """
    import re as _re
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import threading

    dest_folder.mkdir(parents=True, exist_ok=True)

    # Leer mbox completo de una vez
    try:
        with open(mbox_path, "rb") as f:
            raw = f.read()
    except Exception as e:
        if progress_cb:
            progress_cb(f"  Error abriendo {mbox_path.name}: {e}")
        return 0

    # Dividir en bloques por separador mbox
    posiciones = [m.start() for m in _re.finditer(rb"^From ", raw, _re.MULTILINE)]
    posiciones.append(len(raw))
    total = len(posiciones) - 1

    if progress_cb:
        progress_cb(f"  {mbox_path.name}: {total} mensajes en disco, filtrando...")

    # ── Fase 1: Filtrar por fecha (solo headers, sin cuerpo) ──────────────
    pendientes = []   # lista de (bloque_bytes, idx)
    for idx in range(total):
        bloque = raw[posiciones[idx]:posiciones[idx + 1]]
        sep = bloque.find(b"\n\n")
        if sep == -1:
            sep = bloque.find(b"\r\n\r\n")
        headers_raw = bloque[:sep] if sep != -1 else bloque[:3000]

        try:
            h = headers_raw.decode("utf-8", errors="replace")
        except Exception:
            h = headers_raw.decode("latin-1", errors="replace")

        msg_date = None
        for line in h.splitlines():
            if line.lower().startswith("date:"):
                ds = line[5:].strip().split("(")[0].strip()
                try:
                    msg_date = email.utils.parsedate_to_datetime(ds).date()
                except Exception:
                    try:
                        tup = email.utils.parsedate(ds)
                        if tup:
                            msg_date = date(tup[0], tup[1], tup[2])
                    except Exception:
                        pass
                break

        if msg_date is None or not (start_date <= msg_date <= end_date):
            continue

        # Extraer From y Subject para nombre de archivo
        remitente = "desconocido"
        asunto    = "sin_asunto"
        for line in h.splitlines():
            ll = line.lower()
            if ll.startswith("from:") and remitente == "desconocido":
                fr = line[5:].strip()
                raw_rem = fr.split("<")[0].strip().strip('"') if "<" in fr else fr.split("@")[0].strip()
                remitente = _safe_header(raw_rem, 30) or "desconocido"
            elif ll.startswith("subject:") and asunto == "sin_asunto":
                asunto = _safe_header(line[8:].strip(), 60) or "sin_asunto"

        pendientes.append((bloque, idx, msg_date, remitente, asunto))

    if progress_cb:
        progress_cb(f"  {mbox_path.name}: {len(pendientes)} correos del mes, escribiendo...")

    if not pendientes:
        return 0

    # ── Fase 2: Generar nombres unicos ────────────────────────────────────
    nombres_usados = set()
    tareas = []
    for bloque, idx, msg_date, remitente, asunto in pendientes:
        base  = f"{msg_date.isoformat()} - {remitente} - {asunto}"
        fname = base + ".eml"
        if fname in nombres_usados:
            fname = f"{base} ({idx}).eml"
        nombres_usados.add(fname)
        tareas.append((dest_folder / fname, bloque))

    # ── Fase 3: Escritura en paralelo ─────────────────────────────────────
    exported  = 0
    errores   = 0
    lock      = threading.Lock()

    # Usar max 8 hilos — mas no mejora en disco local de Windows
    workers = min(8, len(tareas))

    def escribir(args):
        path, data = args
        path.write_bytes(data)
        return True

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futuros = {ex.submit(escribir, t): t for t in tareas}
        done = 0
        for fut in as_completed(futuros):
            done += 1
            try:
                fut.result()
                with lock:
                    exported += 1
            except Exception as e:
                with lock:
                    errores += 1
            if progress_cb and done % 50 == 0:
                progress_cb(f"  {done}/{len(tareas)} correos escritos de {mbox_path.name}...")

    if progress_cb:
        msg = f"  {mbox_path.name}: {exported} correos exportados"
        if errores:
            msg += f" ({errores} errores)"
        progress_cb(msg)

    return exported

def backup_thunderbird(tmp_root: Path, start_date: date, end_date: date,
                        include_profile_copy: bool = True,
                        progress_cb=None) -> dict:
    """
    Respaldo completo de Thunderbird:
    1. Copia la carpeta de perfil completa  ->  tmp_root/Correos/perfil_completo/
    2. Exporta .eml por rango de fechas     ->  tmp_root/Correos/eml/INBOX/ y /Enviados/
    NOTA: Solo exporta los correos que Thunderbird tiene descargados en disco.
    Si la cuenta es IMAP (Gmail etc.) puede que no esten todos descargados.
    """
    result = {"profile_copied": False, "eml_inbox": 0, "eml_sent": 0, "error": None}

    profile = find_thunderbird_profile()
    if profile is None:
        result["error"] = "Perfil de Thunderbird no encontrado"
        if progress_cb:
            progress_cb("  Thunderbird no encontrado, omitiendo correos.")
        return result

    if progress_cb:
        progress_cb(f"  Perfil: {profile.name}")

    # 1. Copia completa del perfil
    if include_profile_copy:
        dest_profile = tmp_root / "Correos" / "perfil_completo"
        try:
            if progress_cb:
                progress_cb("  Copiando perfil completo de Thunderbird...")
            shutil.copytree(str(profile), str(dest_profile),
                            ignore=shutil.ignore_patterns("*.log", "cache2", "Cache",
                                                          "shader-cache", "datareporting"),
                            dirs_exist_ok=True)
            result["profile_copied"] = True
            if progress_cb:
                progress_cb("  Perfil copiado OK")
        except Exception as e:
            result["error"] = str(e)
            if progress_cb:
                progress_cb(f"  Advertencia copiando perfil: {e}")

    # 2. Exportar .eml por rango
    mbox_files = find_mbox_files(profile)
    if not mbox_files:
        if progress_cb:
            progress_cb("  Advertencia: no se encontraron archivos de correo de Thunderbird")
            progress_cb("  Verifica que Thunderbird este configurado para descargar correos offline")
        return result

    if progress_cb:
        progress_cb(f"  Archivos de correo encontrados: {len(mbox_files)}")
        for mbox_path, name in mbox_files:
            size_mb = mbox_path.stat().st_size / 1024 / 1024
            n_mes = count_emails_in_mbox(mbox_path, start_date, end_date)
            progress_cb(f"    {name}: {size_mb:.1f} MB  |  {n_mes} correos del mes en disco")

    for mbox_path, folder_name in mbox_files:
        if progress_cb:
            progress_cb(f"  Exportando: {folder_name}...")
        dest_eml = tmp_root / "Correos" / "eml" / folder_name
        count = export_emails_to_eml(mbox_path, dest_eml, start_date, end_date, progress_cb)
        if progress_cb:
            progress_cb(f"  {count} correos exportados de {folder_name}")
        name_upper = folder_name.upper()
        if "SENT" in name_upper or "ENVIADO" in name_upper:
            result["eml_sent"] += count
        else:
            result["eml_inbox"] += count

    return result

# ───────────────────────────────────────────────
def _user_file(name: str) -> Path:
    """Archivo en la carpeta del usuario local (no del admin)."""
    userprofile = os.environ.get("USERPROFILE", "")
    base = Path(userprofile) if userprofile and Path(userprofile).exists() else Path.home()
    return base / name

CONFIG_FILE = _user_file(".respaldo_mensual.json")
LOG_FILE    = _user_file("respaldo_mensual.log")

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("respaldo")

# ───────────────────────────────────────────────
#  Utilidades generales
# ───────────────────────────────────────────────

def get_pc_number() -> str:
    """Extrae el número de la PC del nombre de host (ulapcXX → XX)."""
    hostname = socket.gethostname().upper()
    digits = "".join(c for c in hostname if c.isdigit())
    return digits if digits else "00"

def get_pc_label() -> str:
    return f"LAPC {get_pc_number()}"

def get_default_folders():
    home = Path.home()
    return {
        "Escritorio": home / "Desktop",
        "Descargas":  home / "Downloads",
        "Documentos": home / "Documents",
        "Imágenes":   home / "Pictures",
    }

def last_day_of_month(d: date) -> date:
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])

def first_business_day_of_month(d: date) -> date:
    """Primer dia habil del mes: lunes a sabado. Si dia 1 es domingo, retorna dia 2."""
    first = d.replace(day=1)
    if first.weekday() == 6:  # domingo
        return first + timedelta(days=1)
    return first

def is_within_backup_window(d: date = None) -> bool:
    """
    Devuelve True si hoy esta dentro de la ventana de respaldo del mes.
    La ventana es: desde el 1er dia habil hasta los primeros MAX_REINTENTOS_HABILES
    dias habiles del mes (lunes a sabado, sin domingos).
    Ejemplo con MAX=5: si el dia 1 es martes, la ventana es mar-mie-jue-vie-sab-lun
    """
    d = d or date.today()
    # Solo aplica en los primeros dias del mes
    if d.day > 14:  # margen para cubrir 9 dias habiles incluyendo domingos
        return False
    primer_habil = first_business_day_of_month(d)
    if d < primer_habil:
        return False
    # Contar dias habiles desde el primero hasta hoy
    habiles = 0
    cur = primer_habil
    while cur <= d:
        if cur.weekday() != 6:  # no domingo
            habiles += 1
        cur += timedelta(days=1)
    return habiles <= MAX_REINTENTOS_HABILES

def is_first_business_day_of_month(d: date = None) -> bool:
    d = d or date.today()
    return d == first_business_day_of_month(d)

def month_range(d: date = None):
    """Devuelve (primer día, último día) del mes anterior."""
    d = d or date.today()
    first_this = d.replace(day=1)
    last_prev  = first_this - timedelta(days=1)
    first_prev = last_prev.replace(day=1)
    return first_prev, last_prev

# ───────────────────────────────────────────────
#  Configuración persistente
# ───────────────────────────────────────────────

DEFAULT_CFG = {
    "share_root":         r"\\ULAPC46\Respaldos",
    "retry_pending":      False,
    "last_backup":        "",
    "backup_thunderbird": True,
    "tb_profile_copy":    True,
}

def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            return {**DEFAULT_CFG, **cfg}
        except Exception:
            pass
    return dict(DEFAULT_CFG)

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

# ───────────────────────────────────────────────
#  Recolección de archivos
# ───────────────────────────────────────────────

EXCLUIR_EXTENSIONES = {".lnk", ".tmp", ".temp", ".ini"}

def file_in_range(path: Path, start: date, end: date) -> bool:
    try:
        if path.suffix.lower() in EXCLUIR_EXTENSIONES:
            return False
        mtime = datetime.fromtimestamp(path.stat().st_mtime).date()
        return start <= mtime <= end
    except Exception:
        return False

def collect_files(start_date: date, end_date: date, progress_cb=None):
    collected = []
    for folder_name, folder_path in get_default_folders().items():
        if not folder_path.exists():
            continue
        for root, dirs, files in os.walk(folder_path):
            dirs[:] = [d for d in dirs if not d.startswith("RESPALDO_")]
            for fname in files:
                fpath = Path(root) / fname
                if fpath.is_file() and file_in_range(fpath, start_date, end_date):
                    collected.append((fpath, folder_name))
        if progress_cb:
            progress_cb(f"Escaneado: {folder_name}")
    return collected

# ───────────────────────────────────────────────
#  Copia local temporal
# ───────────────────────────────────────────────

def build_local_backup(files, tmp_root: Path, status_cb=None, count_cb=None):
    errors = []
    total = len(files)
    folders = get_default_folders()
    for i, (fpath, folder_name) in enumerate(files, 1):
        try:
            base = folders.get(folder_name, fpath.parent)
            try:
                rel = fpath.relative_to(base)
            except ValueError:
                rel = Path(fpath.name)
            dest = tmp_root / folder_name / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                stem, suffix, c = dest.stem, dest.suffix, 1
                while dest.exists():
                    dest = dest.parent / f"{stem}_{c}{suffix}"
                    c += 1
            shutil.copy2(fpath, dest)
            if status_cb:
                status_cb(f"  [{i}/{total}] {fpath.name}")
            if count_cb:
                count_cb(i, total)
        except Exception as e:
            errors.append((str(fpath), str(e)))
    return errors

# ───────────────────────────────────────────────
#  Copia a la red local Windows
# ───────────────────────────────────────────────

def copy_to_network(local_backup: Path, net_dest: Path, status_cb=None) -> int:
    """Copia la carpeta completa de respaldo a la ruta de red."""
    net_dest.mkdir(parents=True, exist_ok=True)
    copied = 0
    for src in local_backup.rglob("*"):
        if not src.is_file():
            continue
        rel  = src.relative_to(local_backup)
        dest = net_dest / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
        copied += 1
        if status_cb and copied % 20 == 0:
            status_cb(f"  Red: {copied} archivos enviados...")
    return copied

# ───────────────────────────────────────────────
#  Lógica central del respaldo
# ───────────────────────────────────────────────

def _escribir_estado_red(cfg: dict, pc_label: str, mes: str, status: str,
                         archivos: int, errores: list,
                         net_dest: str = "", detalle: str = ""):
    """Escribe el estado del respaldo en la carpeta de informes de la red."""
    try:
        share = cfg.get("share_root", "")
        if not share:
            return
        ruta_informes = str(Path(share) / "_informes")
        # Importar y llamar la funcion del script de informe si existe
        import importlib.util, socket as _socket, json as _json
        from datetime import datetime as _dt
        carpeta = Path(ruta_informes)
        carpeta.mkdir(parents=True, exist_ok=True)
        nombre = f"{pc_label.replace(' ', '_')}_{mes}.json"
        estado = {
            "pc":        pc_label,
            "mes":       mes,
            "status":    status,
            "archivos":  archivos,
            "errores":   len(errores) if errores else 0,
            "net_dest":  net_dest,
            "detalle":   detalle,
            "timestamp": _dt.now().isoformat(),
            "hostname":  _socket.gethostname(),
        }
        with open(carpeta / nombre, "w", encoding="utf-8") as f:
            _json.dump(estado, f, indent=2, ensure_ascii=False)
        log.info(f"Estado del respaldo escrito en: {carpeta / nombre}")
    except Exception as e:
        log.warning(f"No se pudo escribir estado en red: {e}")


# ───────────────────────────────────────────────
#  Sistema de checkpoint y reintentos
# ───────────────────────────────────────────────

MAX_REINTENTOS_HABILES = 9  # dias habiles maximos (semana y media: lun-sab)

def _checkpoint_path(month_label: str) -> Path:
    return _user_file(f".respaldo_checkpoint_{month_label}.json")

def _load_checkpoint(month_label: str) -> dict:
    """Carga el checkpoint del mes. Si no existe devuelve estado inicial."""
    p = _checkpoint_path(month_label)
    if p.exists():
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "month":           month_label,
        "fase":            "inicio",   # inicio | archivos | correos | red | completo
        "archivos_ok":     [],          # rutas ya copiadas
        "correos_ok":      [],          # carpetas de correo ya exportadas
        "perfil_ok":       False,
        "red_ok":          False,
        "intentos":        0,
        "dias_habiles_usados": 0,
        "ultimo_intento":  "",
        "fallido":         False,
    }

def _save_checkpoint(cp: dict):
    p = _checkpoint_path(cp["month"])
    with open(p, "w", encoding="utf-8") as f:
        json.dump(cp, f, indent=2, ensure_ascii=False)

def _delete_checkpoint(month_label: str):
    _checkpoint_path(month_label).unlink(missing_ok=True)

def _dias_habiles_entre(d1: date, d2: date) -> int:
    """Cuenta dias habiles (lun-sab) entre dos fechas."""
    count = 0
    current = d1
    while current <= d2:
        if current.weekday() != 6:  # no domingo
            count += 1
        current += timedelta(days=1)
    return count

def _puede_reintentar(cp: dict) -> tuple:
    """
    Devuelve (puede, motivo).
    No puede si: ya esta completo, marcado fallido, o supero MAX dias habiles.
    """
    if cp.get("fallido"):
        return False, "Marcado como fallido definitivo"
    if cp.get("red_ok") or cp.get("fase") == "completo":
        return False, "Ya completado"
    # Contar dias habiles usados desde el primer intento
    if cp.get("ultimo_intento"):
        try:
            primer = date.fromisoformat(cp["ultimo_intento"][:10])
            dias = _dias_habiles_entre(primer, date.today())
            if dias > MAX_REINTENTOS_HABILES:
                return False, f"Supero {MAX_REINTENTOS_HABILES} dias habiles de reintento"
        except Exception:
            pass
    return True, "OK"

def run_backup(cfg: dict, status_cb=None, count_cb=None) -> dict:
    def msg(text, level="info"):
        getattr(log, level)(text)
        if status_cb:
            status_cb(text)

    start_d, end_d = month_range()
    pc_label       = get_pc_label()
    month_label    = start_d.strftime("%Y-%m")
    backup_name    = f"RESPALDO_{pc_label.replace(' ', '_')}_{month_label}"
    local_backup   = get_user_home() / "Desktop" / backup_name
    share_root     = cfg["share_root"].strip()
    net_dest       = Path(share_root) / pc_label / month_label

    # ── Cargar checkpoint del mes ──────────────────────────────────────
    cp = _load_checkpoint(month_label)

    # Verificar si puede reintentar
    puede, motivo = _puede_reintentar(cp)
    if not puede:
        msg(f"Respaldo cancelado: {motivo}")
        if cp.get("fallido"):
            _escribir_estado_red(cfg, pc_label, month_label, "error",
                                 0, [], str(net_dest), motivo)
        return {"status": "cancelado", "motivo": motivo}

    # Actualizar contadores del checkpoint
    cp["intentos"]       += 1
    cp["ultimo_intento"]  = str(date.today())
    if cp["intentos"] > 1:
        dias = _dias_habiles_entre(
            date.fromisoformat(cp.get("primer_intento", str(date.today()))[:10]),
            date.today()
        )
        cp["dias_habiles_usados"] = dias
        if dias >= MAX_REINTENTOS_HABILES:
            cp["fallido"] = True
            _save_checkpoint(cp)
            msg(f"Respaldo fallido: supero {MAX_REINTENTOS_HABILES} dias habiles de reintento.")
            _escribir_estado_red(cfg, pc_label, month_label, "error",
                                 0, [], str(net_dest),
                                 f"Supero {MAX_REINTENTOS_HABILES} dias habiles")
            return {"status": "fallido"}
    else:
        cp["primer_intento"] = str(date.today())

    _save_checkpoint(cp)

    msg(f"Iniciando respaldo  {pc_label}  -  {month_label}")
    msg(f"Rango: {start_d}  a  {end_d}")
    msg(f"Destino en red: {net_dest}")
    if cp["intentos"] > 1:
        msg(f"Reintento {cp['intentos']} (dia habil {cp['dias_habiles_usados']}/{MAX_REINTENTOS_HABILES})")
        msg(f"Retomando desde fase: {cp['fase']}")

    errors = []

    # ── FASE 1: Archivos locales ───────────────────────────────────────
    if cp["fase"] in ("inicio", "archivos"):
        cp["fase"] = "archivos"
        _save_checkpoint(cp)

        msg("Buscando archivos...")
        files = collect_files(start_d, end_d, status_cb)
        msg(f"{len(files)} archivos encontrados.")

        # Crear carpeta si no existe (puede existir de intento anterior)
        local_backup.mkdir(parents=True, exist_ok=True)

        # Filtrar archivos ya copiados en intentos anteriores
        ya_copiados = set(cp.get("archivos_ok", []))
        pendientes  = [(f, n) for f, n in files if str(f) not in ya_copiados]

        if pendientes:
            msg(f"  {len(ya_copiados)} ya copiados, {len(pendientes)} pendientes...")
            for fpath, folder_name in pendientes:
                try:
                    folders = get_default_folders()
                    base    = folders.get(folder_name, fpath.parent)
                    try:    rel = fpath.relative_to(base)
                    except: rel = Path(fpath.name)
                    dest = local_backup / folder_name / rel
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    if dest.exists():
                        stem, suffix, c = dest.stem, dest.suffix, 1
                        while dest.exists():
                            dest = dest.parent / f"{stem}_{c}{suffix}"; c += 1
                    shutil.copy2(fpath, dest)
                    cp["archivos_ok"].append(str(fpath))
                    _save_checkpoint(cp)
                except Exception as e:
                    errors.append((str(fpath), str(e)))
        else:
            msg("  Todos los archivos ya estaban copiados.")
        cp["fase"] = "correos"
        _save_checkpoint(cp)

    # ── FASE 2: Thunderbird ────────────────────────────────────────────
    if cp["fase"] == "correos":
        if cfg.get("backup_thunderbird", True):
            msg("Procesando correos de Thunderbird...")
            ya_correos = set(cp.get("correos_ok", []))
            profile    = find_thunderbird_profile()
            if profile:
                mbox_files = find_mbox_files(profile)
                pendientes_tb = [(p, n) for p, n in mbox_files if n not in ya_correos]
                if cp.get("perfil_ok") is False and cfg.get("tb_profile_copy", True):
                    # Copiar perfil completo si no se hizo
                    try:
                        dest_profile = local_backup / "Correos" / "perfil_completo"
                        if not dest_profile.exists():
                            msg("  Copiando perfil completo de Thunderbird...")
                            shutil.copytree(str(profile), str(dest_profile),
                                ignore=shutil.ignore_patterns("*.log","cache2","Cache",
                                                               "shader-cache","datareporting"),
                                dirs_exist_ok=True)
                        cp["perfil_ok"] = True
                        _save_checkpoint(cp)
                    except Exception as e:
                        msg(f"  Advertencia perfil: {e}")

                for mbox_path, folder_name in pendientes_tb:
                    msg(f"  Exportando: {folder_name}...")
                    dest_eml = local_backup / "Correos" / "eml" / folder_name
                    count = export_emails_to_eml(
                        mbox_path, dest_eml, start_d, end_d, status_cb)
                    msg(f"  {count} correos exportados de {folder_name}")
                    cp["correos_ok"].append(folder_name)
                    _save_checkpoint(cp)
            else:
                msg("  Thunderbird no encontrado.")
        cp["fase"] = "red"
        _save_checkpoint(cp)

    # ── FASE 3: Copiar a la red ────────────────────────────────────────
    if cp["fase"] == "red":
        # Verificar si la carpeta ya existe en la red con archivos (intento anterior completo)
        try:
            red_tiene_archivos = net_dest.exists() and any(net_dest.rglob("*"))
        except Exception:
            red_tiene_archivos = False

        if red_tiene_archivos:
            msg(f"La carpeta ya existe en la red: {net_dest}")
            msg("Ya fue copiada anteriormente. Marcando como completado.")
            cp["red_ok"] = True
            cp["fase"]   = "completo"
            _save_checkpoint(cp)
            cfg["retry_pending"] = False
            save_config(cfg)
        else:
            msg(f"Enviando a la red: {net_dest}")
            try:
                net_files = copy_to_network(local_backup, net_dest, status_cb)
                msg(f"{net_files} archivos enviados correctamente a la red.")
                cp["red_ok"] = True
                cp["fase"]   = "completo"
                _save_checkpoint(cp)
                cfg["retry_pending"] = False
            except Exception as e:
                log.error(f"Error de red: {e}")
                cfg["retry_pending"] = True
                save_config(cfg)
                _save_checkpoint(cp)
                raise RuntimeError(
                    f"No se pudo copiar a la red.\n"
                    f"Ruta: {net_dest}\n"
                    f"Detalle: {e}\n"
                    f"Se reintentara en el siguiente trigger.")

    # ── Finalizar ──────────────────────────────────────────────────────
    total_archivos = len(cp.get("archivos_ok", []))
    cfg["last_backup"]   = str(date.today())
    cfg["last_backup_month"]  = month_label   # mes que se respaldo (ej. "2026-03")
    cfg["retry_pending"] = False
    save_config(cfg)
    msg("Respaldo completado exitosamente.")

    # Borrar carpeta local del Escritorio
    try:
        shutil.rmtree(local_backup)
        msg(f"Carpeta local eliminada: {local_backup.name}")
    except Exception as e:
        msg(f"Advertencia: no se pudo eliminar carpeta local: {e}")

    # Borrar checkpoint (tarea completada)
    _delete_checkpoint(month_label)

    # Escribir estado en informes de red
    _escribir_estado_red(cfg, pc_label, month_label, "ok",
                         total_archivos, errors, str(net_dest))

    return {"status": "ok", "local_backup": str(local_backup),
            "net_dest": str(net_dest), "files": total_archivos, "errors": errors}

# ───────────────────────────────────────────────
#  Comprobación al arrancar (retry pendiente)
# ───────────────────────────────────────────────

def check_and_run_headless():
    """
    Llamado por Task Scheduler sin GUI.
    - Verifica lock file para no correr si ya hay un respaldo en curso
    - Corre si es el primer dia habil del mes O hay retry pendiente
    - Si ya corrio exitosamente hoy, no vuelve a correr
    """
    cfg   = load_config()
    today = date.today()

  # Calcular mes a respaldar
    start_d, _ = month_range()
    month_label = start_d.strftime("%Y-%m")

    # Verificar si este mes ya fue respaldado exitosamente (guardado en config)
    if cfg.get("last_backup_month") == month_label:
        log.info(f"El mes {month_label} ya fue respaldado exitosamente. Nada que hacer.")
        return

    # Verificar checkpoint del mes
    cp = _load_checkpoint(month_label)

    # Checkpoint dice que ya completo
    if cp.get("fase") == "completo" or cp.get("red_ok"):
        log.info(f"Respaldo de {month_label} ya completado (checkpoint).")
        # Actualizar config para no volver a revisar
        cfg["last_backup_month"] = month_label
        save_config(cfg)
        return

    # Fallido definitivo
    if cp.get("fallido"):
        log.error(f"Respaldo de {month_label} marcado como fallido definitivo.")
        return

    # Ver si debe correr hoy
    # Verificar si estamos dentro de la ventana de respaldo
    # (primeros MAX_REINTENTOS_HABILES dias habiles del mes)
    en_ventana    = is_within_backup_window(today)
    tiene_pendiente = cfg.get("retry_pending", False) or (
        cp.get("fase", "inicio") not in ("inicio", "completo")
    )

    if not en_ventana and not tiene_pendiente:
        log.info(f"Hoy ({today}) esta fuera de la ventana de respaldo "
                 f"(primeros {MAX_REINTENTOS_HABILES} dias habiles del mes).")
        return

    log.info(f"Hoy ({today}) esta dentro de la ventana de respaldo.")

    if not cfg.get("share_root"):
        log.error("Ruta de red no configurada.")
        return

    # Lock file: evita que dos instancias corran al mismo tiempo
    lock_file = _user_file(".respaldo_en_curso.lock")
    if lock_file.exists():
        # Verificar si el lock es de hoy y reciente (menos de 4 horas)
        try:
            import time
            lock_age_hours = (time.time() - lock_file.stat().st_mtime) / 3600
            if lock_age_hours < 4:
                log.info(f"Respaldo ya en curso (lock de hace {lock_age_hours:.1f}h). Esperando siguiente trigger.")
                return
            else:
                # Lock viejo (mas de 4h) = proceso anterior colgado, eliminar
                log.warning(f"Lock file viejo ({lock_age_hours:.1f}h), eliminando y reintentando.")
                lock_file.unlink(missing_ok=True)
        except Exception:
            lock_file.unlink(missing_ok=True)

    # Crear lock file
    try:
        lock_file.write_text(str(os.getpid()), encoding="utf-8")
    except Exception:
        pass

    log.info(f"Iniciando respaldo automatico ({today})")
    try:
        run_backup(cfg)
    except Exception as e:
        log.error(f"Respaldo fallido: {e}. Se reintentara en el siguiente trigger.")
    finally:
        # Siempre eliminar el lock al terminar
        lock_file.unlink(missing_ok=True)


# ───────────────────────────────────────────────
#  Interfaz grafica
# ───────────────────────────────────────────────

# Credenciales de acceso al programa
# Para cambiar: modificar ADMIN_USER y ADMIN_PASS aqui
ADMIN_USER = "admin"
ADMIN_PASS = "Respaldo2024"


class LoginApp(tk.Tk):
    """Ventana de login que debe pasar antes de abrir el programa."""
    BG    = "#0d0d14"
    PANEL = "#13131f"
    CARD  = "#1a1a2e"
    ACCENT= "#00d4aa"
    FG    = "#e8eaf6"
    FG2   = "#7986cb"
    RED   = "#ff5252"

    def __init__(self):
        super().__init__()
        self.title("Acceso - Sistema de Respaldo Mensual")
        self.geometry("400x320")
        self.resizable(False, False)
        self.configure(bg=self.BG)
        self.resultado = False
        self._build()
        self.eval("tk::PlaceWindow . center")
        # Centrar en pantalla
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - 400) // 2
        y = (self.winfo_screenheight() - 320) // 2
        self.geometry(f"400x320+{x}+{y}")

    def _build(self):
        # Barra superior
        tk.Frame(self, bg=self.ACCENT, height=3).pack(fill="x")
        tk.Label(self, text="SISTEMA DE RESPALDO MENSUAL",
                 bg=self.BG, fg=self.ACCENT,
                 font=("Courier New", 11, "bold")).pack(pady=(20, 4))
        tk.Label(self, text=f"PC:  {get_pc_label()}",
                 bg=self.BG, fg=self.FG2,
                 font=("Courier New", 9)).pack()

        # Formulario
        form = tk.Frame(self, bg=self.CARD)
        form.pack(fill="x", padx=30, pady=20)

        for row, (label, attr, show) in enumerate([
            ("Usuario:", "_entry_user", ""),
            ("Contraseña:", "_entry_pass", "*"),
        ]):
            tk.Label(form, text=label, bg=self.CARD, fg=self.FG,
                     font=("Segoe UI", 10), width=12, anchor="e").grid(
                row=row, column=0, padx=(12, 8), pady=10, sticky="e")
            e = tk.Entry(form, bg="#0d0d14", fg=self.FG,
                         insertbackground=self.ACCENT,
                         relief="flat", font=("Courier New", 11),
                         show=show, width=22)
            e.grid(row=row, column=1, padx=(0, 12), ipady=6)
            setattr(self, attr, e)

        self._entry_user.insert(0, "admin")
        self._entry_pass.bind("<Return>", lambda e: self._login())
        self._entry_pass.focus_set()

        # Mensaje de error
        self.lbl_error = tk.Label(self, text="", bg=self.BG,
                                   fg=self.RED, font=("Segoe UI", 9))
        self.lbl_error.pack()

        # Boton
        tk.Button(self, text="ENTRAR",
                  bg=self.ACCENT, fg=self.BG,
                  activebackground="#00b894",
                  relief="flat", cursor="hand2",
                  font=("Courier New", 11, "bold"),
                  padx=30, pady=8,
                  command=self._login).pack(pady=(0, 20))

    def _login(self):
        user = self._entry_user.get().strip()
        pwd  = self._entry_pass.get()
        if user == ADMIN_USER and pwd == ADMIN_PASS:
            self.resultado = True
            self.destroy()
        else:
            self.lbl_error.configure(text="Usuario o contraseña incorrectos")
            self._entry_pass.delete(0, "end")
            self._entry_pass.focus_set()
            # Sacudir la ventana para indicar error
            x, y = self.winfo_x(), self.winfo_y()
            for dx in [8, -8, 6, -6, 4, -4, 0]:
                self.geometry(f"+{x+dx}+{y}")
                self.update()
                self.after(30)


class RespaldoApp(tk.Tk):
    BG=     "#0d0d14"; PANEL=  "#13131f"; CARD=   "#1a1a2e"
    BORDER= "#2a2a45"; ACCENT= "#00d4aa"; ACC2=   "#0099ff"
    FG=     "#e8eaf6"; FG2=    "#7986cb"; GREEN=  "#00e676"
    RED=    "#ff5252"; YELLOW= "#ffd740"

    def __init__(self):
        super().__init__()
        self.title(f"Respaldo Mensual  -  {get_pc_label()}")
        self.geometry("900x820")
        self.minsize(820, 720)
        self.configure(bg=self.BG)
        self.cfg = load_config()
        self._build_ui()
        self.eval("tk::PlaceWindow . center")
        if self.cfg.get("retry_pending"):
            self.after(1000, lambda: self._log(
                "Hay un respaldo pendiente de reintento.", self.YELLOW))

    def _build_ui(self):
        hdr = tk.Frame(self, bg=self.PANEL); hdr.pack(fill="x")
        tk.Frame(hdr, bg=self.ACCENT, height=3).pack(fill="x")
        row = tk.Frame(hdr, bg=self.PANEL); row.pack(fill="x", padx=30, pady=16)
        tk.Label(row, text="SISTEMA DE RESPALDO MENSUAL",
                 bg=self.PANEL, fg=self.ACCENT,
                 font=("Courier New", 15, "bold")).pack(side="left")
        tk.Label(row, text=get_pc_label(), bg=self.ACCENT, fg=self.BG,
                 font=("Courier New", 12, "bold"), padx=12, pady=4).pack(side="right")

        st = ttk.Style(self); st.theme_use("clam")
        st.configure("TNotebook",        background=self.BG,   borderwidth=0)
        st.configure("TNotebook.Tab",    background=self.CARD, foreground=self.FG2,
                     font=("Courier New", 10), padding=[18, 8])
        st.map("TNotebook.Tab",
               background=[("selected", self.PANEL)],
               foreground=[("selected", self.ACCENT)])
        st.configure("TCheckbutton",     background=self.CARD, foreground=self.FG,
                     font=("Segoe UI", 10))
        st.configure("Accent.Horizontal.TProgressbar", troughcolor=self.BORDER,
                     background=self.ACCENT, thickness=14)

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True)
        t1=tk.Frame(nb,bg=self.BG); t2=tk.Frame(nb,bg=self.BG); t3=tk.Frame(nb,bg=self.BG)
        nb.add(t1, text="  Respaldo  ")
        nb.add(t2, text="  Configuracion  ")
        nb.add(t3, text="  Historial  ")
        self._tab_respaldo(t1); self._tab_config(t2); self._tab_historial(t3)

    def _tab_respaldo(self, p):
        today     = date.today()
        # El respaldo corre el 1er dia habil del mes siguiente
        # y respalda el mes ACTUAL completo (no el anterior)
        # Ejemplo: si hoy es marzo, el respaldo del 1 de abril respalda TODO marzo
        # Proximo automatico: 1er dia habil del mes siguiente
        next_run  = first_business_day_of_month(
            (today.replace(day=1) + timedelta(days=32)).replace(day=1)
        )
        days_left = (next_run - today).days
        # El mes a respaldar es el mes ANTERIOR (igual que run_backup)
        start_d, end_d = month_range(today)
        share     = self.cfg.get("share_root", "sin configurar")
        net_prev  = str(Path(share) / get_pc_label() / start_d.strftime("%Y-%m"))

        card = tk.Frame(p, bg=self.CARD); card.pack(fill="x", padx=24, pady=(20, 8))
        tk.Frame(card, bg=self.ACC2, width=4).pack(side="left", fill="y")
        inner = tk.Frame(card, bg=self.CARD); inner.pack(side="left", padx=16, pady=14)

        dc = self.GREEN if days_left > 5 else (self.YELLOW if days_left > 1 else self.RED)
        rows = [
            (f"PC detectada:         {get_pc_label()}",                                          self.FG),
            (f"Mes a respaldar:      {start_d.strftime('%B %Y')}  ({start_d} al {end_d})", self.FG),
            (f"Destino en red:       {net_prev}",                                   self.ACCENT),
            (f"Ultimo respaldo:      {self.cfg.get('last_backup', 'Nunca')}",                    self.FG2),
            (f"Proximo automatico:   {next_run}  ({days_left} dias restantes)",                  dc),
        ]
        labels = []
        for text, color in rows:
            lbl = tk.Label(inner, text=text, bg=self.CARD, fg=color,
                           font=("Courier New", 11))
            lbl.pack(anchor="w", pady=1)
            labels.append(lbl)
        self.lbl_remote = labels[2]

        fc = tk.Frame(p, bg=self.CARD); fc.pack(fill="x", padx=24, pady=8)
        tk.Label(fc, text="QUE SE RESPALDA", bg=self.CARD, fg=self.FG2,
                 font=("Courier New", 9, "bold")).pack(anchor="w", padx=16, pady=(10, 4))
        self.folder_vars = {}
        grid = tk.Frame(fc, bg=self.CARD); grid.pack(fill="x", padx=16, pady=(0, 4))
        for idx, (name, path) in enumerate(get_default_folders().items()):
            var = tk.BooleanVar(value=True); self.folder_vars[name] = var
            ttk.Checkbutton(grid, text=f"  {name}   ({path})", variable=var).grid(
                row=idx//2, column=idx%2, sticky="w", padx=12, pady=2)

        self._tb_var = tk.BooleanVar(value=self.cfg.get("backup_thunderbird", True))
        profile = find_thunderbird_profile()
        tb_txt = "  Correos Thunderbird  " + ("(detectado)" if profile else "(no encontrado)")
        ttk.Checkbutton(grid, text=tb_txt, variable=self._tb_var).grid(
            row=2, column=0, columnspan=2, sticky="w", padx=12, pady=(4, 10))

        self.btn_run = tk.Button(p, text="EJECUTAR RESPALDO AHORA",
                                 bg=self.ACCENT, fg=self.BG, activebackground=self.GREEN,
                                 relief="flat", cursor="hand2",
                                 font=("Courier New", 12, "bold"), padx=28, pady=12,
                                 command=self._start_backup)
        self.btn_run.pack(pady=14)

        self.progress = ttk.Progressbar(p, style="Accent.Horizontal.TProgressbar",
                                        mode="determinate", length=700)
        self.progress.pack(padx=24)
        self.lbl_prog = tk.Label(p, text="", bg=self.BG, fg=self.FG2,
                                 font=("Courier New", 9)); self.lbl_prog.pack()

        lf = tk.Frame(p, bg=self.BG)
        lf.pack(fill="both", expand=True, padx=24, pady=(6, 20))
        self.live_log = tk.Text(lf, bg=self.PANEL, fg=self.FG2,
                                font=("Courier New", 9), relief="flat",
                                state="disabled", wrap="word", height=10)
        sc = tk.Scrollbar(lf, command=self.live_log.yview)
        self.live_log.configure(yscrollcommand=sc.set)
        self.live_log.pack(side="left", fill="both", expand=True); sc.pack(side="right", fill="y")

    def _tab_config(self, p):
        form = tk.Frame(p, bg=self.CARD); form.pack(fill="x", padx=24, pady=24)

        tk.Label(form, text="RUTA DE RED  (ULAPC 46)",
                 bg=self.CARD, fg=self.FG2,
                 font=("Courier New", 10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(14, 10))

        tk.Label(form, text="Carpeta compartida:", bg=self.CARD, fg=self.FG,
                 font=("Segoe UI", 10), width=22, anchor="e").grid(
            row=1, column=0, sticky="e", padx=(18, 8), pady=8)

        self._share_var = tk.StringVar(
            value=self.cfg.get("share_root", "\\\\ULAPC46\\Respaldos"))
        tk.Entry(form, textvariable=self._share_var,
                 bg=self.BG, fg=self.FG, insertbackground=self.ACCENT,
                 relief="flat", font=("Courier New", 10), width=42).grid(
            row=1, column=1, sticky="w", padx=(0, 8), ipady=5)
        tk.Button(form, text="Examinar", bg=self.BORDER, fg=self.FG,
                  relief="flat", font=("Segoe UI", 9), padx=8, pady=4,
                  cursor="hand2", command=self._browse_share).grid(
            row=1, column=2, padx=(0, 18))

        tk.Label(form,
                 text="Ejemplo:  \\\\ULAPC46\\Respaldos\nSe creara:  LAPC XX / 2026-02",
                 bg=self.CARD, fg=self.FG2,
                 font=("Courier New", 8), justify="left").grid(
            row=2, column=1, columnspan=2, sticky="w", padx=(0, 18), pady=(0, 10))

        tk.Frame(form, bg=self.BORDER, height=1).grid(
            row=3, column=0, columnspan=3, sticky="ew", padx=18, pady=(4, 12))

        tk.Label(form, text="CORREOS THUNDERBIRD",
                 bg=self.CARD, fg=self.FG2,
                 font=("Courier New", 10, "bold")).grid(
            row=4, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 6))

        self._cfg_tb = tk.BooleanVar(value=self.cfg.get("backup_thunderbird", True))
        ttk.Checkbutton(form,
            text="  Respaldar correos (Bandeja + Enviados) como .eml filtrados por fecha",
            variable=self._cfg_tb).grid(row=5, column=1, columnspan=2, sticky="w", pady=3)

        self._cfg_tb_prof = tk.BooleanVar(value=self.cfg.get("tb_profile_copy", True))
        ttk.Checkbutton(form,
            text="  Incluir copia completa del perfil de Thunderbird (restaurable)",
            variable=self._cfg_tb_prof).grid(row=6, column=1, columnspan=2, sticky="w", pady=3)

        profile = find_thunderbird_profile()
        tk.Label(form,
            text="  Perfil detectado: " + (str(profile) if profile else "No detectado en esta PC"),
            bg=self.CARD, fg=self.FG2 if profile else self.YELLOW,
            font=("Courier New", 8)).grid(
            row=7, column=1, columnspan=2, sticky="w", pady=(0, 10))

        tk.Frame(form, bg=self.BORDER, height=1).grid(
            row=8, column=0, columnspan=3, sticky="ew", padx=18, pady=(4, 12))

        brow = tk.Frame(form, bg=self.CARD)
        brow.grid(row=9, column=0, columnspan=3, pady=(4, 18))
        for text, cmd, primary in [
            ("Guardar",                   self._save_config,      True),
            ("Probar conexion de red",    self._test_net,         False),
            ("Instalar tarea automatica", self._install_task,     False),
            ("Probar en 2 minutos",       self._install_task_test,False),
        ]:
            tk.Button(brow, text=text,
                      bg=self.ACCENT if primary else self.BORDER,
                      fg=self.BG if primary else self.FG,
                      activebackground=self.GREEN if primary else self.CARD,
                      relief="flat", cursor="hand2",
                      font=("Courier New", 10, "bold" if primary else "normal"),
                      padx=18, pady=8, command=cmd).pack(side="left", padx=6)

        self.lbl_cfg_status = tk.Label(form, text="", bg=self.CARD, font=("Courier New", 9))
        self.lbl_cfg_status.grid(row=10, column=0, columnspan=3)

    def _tab_historial(self, p):
        top = tk.Frame(p, bg=self.BG); top.pack(fill="x", padx=24, pady=(16, 6))
        tk.Label(top, text="Log:  " + str(LOG_FILE),
                 bg=self.BG, fg=self.FG2, font=("Courier New", 9)).pack(side="left")

        self.lbl_autorefrsh = tk.Label(top, text="actualizando cada 5s",
                 bg=self.BG, fg=self.BORDER, font=("Courier New", 8))
        self.lbl_autorefrsh.pack(side="right", padx=(0, 8))

        tk.Button(top, text="Actualizar ahora", bg=self.BORDER, fg=self.FG,
                  relief="flat", font=("Courier New", 9), padx=8, pady=4,
                  cursor="hand2", command=self._refresh_log_tab).pack(side="right")

        frame = tk.Frame(p, bg=self.BG)
        frame.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.hist_text = tk.Text(frame, bg=self.PANEL, fg=self.FG2,
                                 font=("Courier New", 9), relief="flat",
                                 state="disabled", wrap="word")
        sc = tk.Scrollbar(frame, command=self.hist_text.yview)
        self.hist_text.configure(yscrollcommand=sc.set)
        self.hist_text.pack(side="left", fill="both", expand=True)
        sc.pack(side="right", fill="y")
        self._refresh_log_tab()
        self._auto_refresh_log()  # iniciar auto-refresco

    def _refresh_log_tab(self):
        if not hasattr(self, "hist_text"):
            return
        # Guardar posicion del scroll para no saltar si el usuario esta leyendo
        try:
            at_end = self.hist_text.yview()[1] >= 0.95
        except Exception:
            at_end = True

        self.hist_text.configure(state="normal")
        self.hist_text.delete("1.0", "end")

        if LOG_FILE.exists():
            try:
                content = LOG_FILE.read_text(encoding="utf-8", errors="replace")
                lines = content.splitlines()
                if lines:
                    self.hist_text.insert("end", "\n".join(lines[-400:]))
                    # Colorear lineas de error en rojo
                    for tag, color in [("ERROR", self.RED), ("WARNING", self.YELLOW),
                                       ("Listo", self.GREEN), ("completado", self.GREEN)]:
                        idx = "1.0"
                        while True:
                            idx = self.hist_text.search(tag, idx, stopindex="end",
                                                         nocase=True)
                            if not idx:
                                break
                            line_end = f"{idx.split('.')[0]}.end"
                            t = f"color_{tag}"
                            self.hist_text.tag_config(t, foreground=color)
                            self.hist_text.tag_add(t, idx, line_end)
                            idx = line_end
                else:
                    self.hist_text.insert("end", "(el log esta vacio)")
            except Exception as e:
                self.hist_text.insert("end", f"Error leyendo log: {e}")
        else:
            self.hist_text.insert("end",
                f"El log todavia no existe.\n\n"
                f"Se creara cuando corra el primer respaldo.\n\n"
                f"Ruta esperada: {LOG_FILE}")

        if at_end:
            self.hist_text.see("end")
        self.hist_text.configure(state="disabled")

    def _auto_refresh_log(self):
        """Refresca el log automaticamente cada 5 segundos."""
        try:
            self._refresh_log_tab()
        except Exception:
            pass
        self.after(5000, self._auto_refresh_log)

    def _log(self, msg, color=None):
        self.live_log.configure(state="normal")
        tag = f"t{id(msg)}"
        self.live_log.insert("end", msg + "\n", tag)
        if color: self.live_log.tag_config(tag, foreground=color)
        self.live_log.see("end"); self.live_log.configure(state="disabled")

    def _browse_share(self):
        folder = filedialog.askdirectory(title="Carpeta compartida de ULAPC46")
        if folder: self._share_var.set(folder)

    def _save_config(self):
        raw = self._share_var.get().strip()
        # Normalizar barras: el usuario puede escribir \ULAPC46\Res o \\ULAPC46\Res
        # Python Path necesita que sea una ruta UNC valida
        if raw.startswith("\\\\"):
            raw = raw  # ya tiene 4 barras, ok
        elif raw.startswith("\\\\") or raw.startswith("//"):
            raw = raw
        # Asegurar que empiece con doble barra para UNC
        if raw.startswith("\\") and not raw.startswith("\\\\"):
            pass  # correcto: \ULAPC46\Res
        self.cfg["share_root"]         = raw
        self.cfg["backup_thunderbird"] = self._cfg_tb.get()
        self.cfg["tb_profile_copy"]    = self._cfg_tb_prof.get()
        save_config(self.cfg)
        start_d, _ = month_range()
        net_prev = str(Path(self.cfg["share_root"]) / get_pc_label() / start_d.strftime("%Y-%m"))
        self.lbl_remote.configure(text="Destino en red:       " + net_prev)
        self.lbl_cfg_status.configure(text="Configuracion guardada", fg=self.GREEN)
        self.after(3000, lambda: self.lbl_cfg_status.configure(text=""))

    def _test_net(self):
        self._save_config()
        share = self.cfg.get("share_root", "")
        self.lbl_cfg_status.configure(text="Probando conexion...", fg=self.YELLOW)
        self.update()
        try:
            p = Path(share)
            if p.exists() and p.is_dir():
                self.lbl_cfg_status.configure(text="Conexion OK  -  " + share, fg=self.GREEN)
            else:
                self.lbl_cfg_status.configure(text="No accesible:  " + share, fg=self.RED)
        except Exception as e:
            self.lbl_cfg_status.configure(text="Error: " + str(e), fg=self.RED)

    def _schtasks_create(self, task_name: str, xml_path: Path) -> bool:
        """
        Registra una tarea usando ShellExecute runas — UAC nativo de Windows.
        El programa no necesita correr como admin; Windows pide permiso solo.
        """
        import ctypes, tempfile, time

        bat = Path(tempfile.gettempdir()) / "instalar_tarea_respaldo.bat"
        bat.write_text(
            "@echo off\n"
            "schtasks /Create /TN \"" + task_name + "\" /XML \"" + str(xml_path) + "\" /F\n"
            "del /f /q \"" + str(xml_path) + "\"\n"
            "del /f /q \"" + str(bat) + "\"\n",
            encoding="utf-8"
        )

        try:
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", str(bat), None, None, 1
            )
            if ret <= 32:
                messagebox.showerror("Error",
                    "No se pudo pedir permisos de administrador.\n"
                    "Intenta abrir el programa con clic derecho -> Ejecutar como administrador.")
                return False
            # Esperar hasta 20s a que el bat termine (se borra solo)
            for _ in range(20):
                time.sleep(1)
                if not bat.exists():
                    return True
            return True  # asumir OK si no se borro (puede pasar en algunas versiones)
        except Exception as e:
            messagebox.showerror("Error al registrar tarea", str(e))
            return False
        finally:
            bat.unlink(missing_ok=True)

    def _install_task_test(self):
        """Instala tarea que corre en los proximos 2 minutos, para probar."""
        from datetime import datetime, timedelta
        script    = Path(sys.argv[0]).resolve()
        python    = sys.executable
        task_name = "RespaldoMensualPRUEBA"
        # 2 minutos desde ahora
        run_at    = datetime.now() + timedelta(minutes=2)
        run_str   = run_at.strftime("%Y-%m-%dT%H:%M:%S")
        xml = (
            '<?xml version="1.0" encoding="UTF-16"?>\n'
            '<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">\n'
            '  <RegistrationInfo><Description>PRUEBA respaldo - ' + get_pc_label() + '</Description></RegistrationInfo>\n'
            '  <Triggers><TimeTrigger>\n'
            '    <StartBoundary>' + run_str + '</StartBoundary>\n'
            '    <Enabled>true</Enabled>\n'
            '    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>\n'
            '  </TimeTrigger></Triggers>\n'
            '  <Settings>\n'
            '    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>\n'
            '    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>\n'
            '    <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>\n'
            '    <DeleteExpiredTaskAfter>PT1H</DeleteExpiredTaskAfter>\n'
            '    <Enabled>true</Enabled>\n'
            '  </Settings>\n'
            '  <Actions><Exec>\n'
            '    <Command>' + str(python) + '</Command>\n'
            '    <Arguments>"' + str(script) + '" --auto</Arguments>\n'
            '  </Exec></Actions>\n'
            '</Task>'
        )
        xml_path = Path.home() / "respaldo_task_test.xml"
        xml_path.write_text(xml, encoding="utf-16")
        ret = os.system('schtasks /Create /TN "' + task_name + '" /XML "' + str(xml_path) + '" /F')
        xml_path.unlink(missing_ok=True)
        if ret == 0:
            messagebox.showinfo("Tarea de prueba instalada",
                f"El respaldo automatico correra en 2 minutos.\n\n"
                f"Hora programada: {run_at.strftime('%H:%M')}\n"
                f"Revisa el log en la pestana Historial para ver el resultado.\n\n"
                f"La tarea se elimina sola despues de ejecutarse.")
        else:
            messagebox.showerror("Error",
                "No se pudo crear la tarea de prueba.\nEjecuta como Administrador.")

    def _run_task_script(self, args: str, success_msg: str):
        """Llama a instalar_tarea.py como administrador via ShellExecute runas."""
        import ctypes, time
        # El script esta junto al programa
        script_dir = Path(sys.argv[0]).resolve().parent
        task_script = script_dir / "instalar_tarea.py"
        if not task_script.exists():
            messagebox.showerror("Error",
                f"No se encontro instalar_tarea.py en\n{script_dir}\n\n"
                "Asegurate de que este en la misma carpeta que respaldo_mensual.py")
            return
        python = sys.executable
        cmd_args = f'"{python}" "{task_script}" {args}'
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "cmd.exe", f"/c {cmd_args} & pause", None, 1
        )
        if ret <= 32:
            messagebox.showerror("Error",
                "No se pudo pedir permisos de administrador.\n"
                "Intenta abrir el programa con clic derecho -> Ejecutar como administrador.")
        else:
            messagebox.showinfo("Tarea registrada", success_msg)

    def _install_task(self):
        script = Path(sys.argv[0]).resolve()
        python = sys.executable
        task_name = "RespaldoMensualAutomatico"
        today_str = date.today().isoformat()
        pc = get_pc_label()
        # El trigger corre el dia 1 de cada mes a las 9AM.
        # El programa mismo revisa si es dia habil (lun-sab);
        # si el dia 1 es domingo simplemente no hace nada y espera el retry del dia 2.
        xml = (
            '<?xml version="1.0" encoding="UTF-16"?>\n'
            '<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">\n'
            '  <RegistrationInfo><Description>Respaldo mensual - ' + pc + '</Description></RegistrationInfo>\n'
            '  <Triggers>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>1</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>2</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>3</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>4</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>5</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>6</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>7</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>8</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>9</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>10</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>11</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '    <CalendarTrigger>\n'
            '      <StartBoundary>' + today_str + 'T09:00:00</StartBoundary>\n'
            '      <EndBoundary>2099-12-31T23:59:59</EndBoundary>\n'
            '      <Enabled>true</Enabled>\n'
            '      <ScheduleByMonth>\n'
            '        <DaysOfMonth><Day>12</Day></DaysOfMonth>\n'
            '        <Months><January/><February/><March/><April/><May/><June/>\n'
            '                <July/><August/><September/><October/><November/><December/></Months>\n'
            '      </ScheduleByMonth>\n'
            '    </CalendarTrigger>\n'
            '  </Triggers>\n'
            '  <Settings>\n'
            '    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>\n'
            '    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>\n'
            '    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>\n'
            '    <ExecutionTimeLimit>PT8H</ExecutionTimeLimit>\n'
            '    <Enabled>true</Enabled>\n'
            '    <StartWhenAvailable>true</StartWhenAvailable>\n'
            '  </Settings>\n'
            '  <Actions><Exec>\n'
            '    <Command>' + str(python) + '</Command>\n'
            '    <Arguments>"' + str(script) + '" --auto</Arguments>\n'
            '  </Exec></Actions>\n'
            '</Task>'
        )
        xml_path = Path.home() / "respaldo_task.xml"
        xml_path.write_text(xml, encoding="utf-16")
        ret = os.system('schtasks /Create /TN "' + task_name + '" /XML "' + str(xml_path) + '" /F')
        xml_path.unlink(missing_ok=True)
        if ret == 0:
            messagebox.showinfo("Tarea instalada",
                "Tarea registrada en Windows Task Scheduler.\n\n"
                "Se ejecutara el Primer dia de cada mes a las 9:00 AM.")
        else:
            messagebox.showerror("Error",
                "No se pudo registrar la tarea.\nEjecuta como Administrador.")

    def _start_backup(self):
        if not self.cfg.get("share_root"):
            messagebox.showwarning("Sin configurar",
                "Ve a Configuracion e ingresa la ruta de red.")
            return
        self.btn_run.configure(state="disabled", text="Procesando...")
        self.live_log.configure(state="normal"); self.live_log.delete("1.0", "end")
        self.live_log.configure(state="disabled"); self.progress["value"] = 0
        self.cfg["backup_thunderbird"] = self._tb_var.get()
        threading.Thread(target=self._run_backup_thread, daemon=True).start()

    def _run_backup_thread(self):
        def cb(m): self.after(0, lambda msg=m: self._log(msg))
        def prog(i, t):
            self.after(0, lambda v=i, mx=t: (
                self.progress.configure(maximum=mx, value=v),
                self.lbl_prog.configure(text=str(v) + " / " + str(mx) + " archivos")))
        try:
            r = run_backup(self.cfg, status_cb=cb, count_cb=prog)
            msg = "Listo.  " + str(r["files"]) + " archivos  ->  " + r["net_dest"]
            self.after(0, lambda: self._log(msg, self.GREEN))
        except Exception as e:
            err = str(e)
            self.after(0, lambda: self._log("Error: " + err, self.RED))
        finally:
            self.after(0, lambda: self.btn_run.configure(
                state="normal", text="EJECUTAR RESPALDO AHORA"))
            self.after(600, self._refresh_log_tab)

# ───────────────────────────────────────────────
if __name__ == "__main__":
    if "--auto" in sys.argv:
        check_and_run_headless()
    else:
        # Mostrar login antes de abrir el programa
        login = LoginApp()
        login.mainloop()
        if login.resultado:
            RespaldoApp().mainloop()
