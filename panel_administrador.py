#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Panel del administrador — Respaldo Mensual
-------------------------------------------------------------------
Pensado para instalarse UNA sola vez en PC100 (la máquina donde llegan
todos los respaldos). Es una página web local: se abre con el navegador
en http://localhost:5000 desde la propia PC100, o desde cualquier otra
PC de la oficina usando la IP de PC100, por ejemplo http://192.168.1.100:5000

No necesita instalarse nada en las 120 laptops para esto: ellas ya
escriben su estado dentro de la carpeta compartida (la de "_estado"),
y este panel solo la lee.

Primer uso:
    pip install flask --break-system-packages   (o sin --break-system-packages en Windows)
    python panel_administrador.py
La primera vez te pedirá la carpeta local de los respaldos (la ruta
real en el disco de PC100 detrás de la carpeta compartida, ej. D:\\Respaldos)
y te deja fijar la contraseña del panel.
"""
import os, sys, json, hashlib, secrets
from pathlib import Path
from datetime import datetime, date

from flask import Flask, request, redirect, url_for, session, jsonify, render_template_string
from markupsafe import Markup

APP_DIR_NAME = "PanelRespaldo"
_PROGRAMDATA = os.environ.get("ProgramData")
if _PROGRAMDATA and os.name == "nt":
    CFG_DIR = Path(_PROGRAMDATA) / APP_DIR_NAME
else:
    CFG_DIR = Path.home() / f".{APP_DIR_NAME.lower()}"
CFG_FILE = CFG_DIR / "config.json"

DEFAULTS = {
    "destino_root": "",                      # ruta LOCAL en PC100, ej. D:\Respaldos
    "password_hash": hashlib.sha256("admin".encode()).hexdigest(),
    "secret_key": secrets.token_hex(16),
    # Lista de máquinas registradas — fuente de verdad local, en el disco de PC100.
    # Nunca se pierde aunque la carpeta de red no esté disponible.
    "maquinas": [],
}


def load_cfg():
    cfg = dict(DEFAULTS)
    try:
        cfg.update(json.loads(CFG_FILE.read_text(encoding="utf-8")))
    except Exception:
        pass
    return cfg


def save_cfg(cfg):
    CFG_DIR.mkdir(parents=True, exist_ok=True)
    CFG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


CFG = load_cfg()
app = Flask(__name__)
app.secret_key = CFG["secret_key"]


def estado_dir():
    return Path(CFG["destino_root"]) / "_estado"


def solicitudes_dir():
    return Path(CFG["destino_root"]) / "_solicitudes"


def prev_month_str():
    t = date.today()
    y, m = (t.year - 1, 12) if t.month == 1 else (t.year, t.month - 1)
    return f"{y}-{m:02d}"


def leer_estado(tag):
    """
    Junta el estado de cada máquina para el mes 'tag' (AAAA-MM).

    Fuente de verdad para la lista de máquinas: CFG["maquinas"] (disco local de PC100).
    Fuente de verdad para el estado de cada una: archivos JSON en la carpeta de red.

    Si la carpeta de red no está disponible, las máquinas siguen apareciendo
    en rojo (sin respaldo) en vez de desaparecer, y se muestra un aviso.
    """
    # 1. Empezar siempre con la lista registrada localmente
    data = {nm: None for nm in get_maquinas()}

    # 2. Leer los archivos de estado de la carpeta de red (fallo silencioso)
    red_ok = False
    d = estado_dir()
    try:
        if d.is_dir():
            red_ok = True
            for fp in d.glob("*.json"):
                nm = fp.stem
                try:
                    estado_mes = json.loads(fp.read_text(encoding="utf-8")).get(tag)
                    # Si la máquina mandó su estado pero no estaba en la lista, agregarla
                    if nm not in data:
                        data[nm] = estado_mes
                        agregar_maquina(nm)   # registrarla para la próxima vez
                    else:
                        data[nm] = estado_mes
                except Exception:
                    pass
            # Leer también _esperadas.txt de la red y agregar cualquier nombre nuevo
            esp = d / "_esperadas.txt"
            if esp.exists():
                for ln in esp.read_text(encoding="utf-8").splitlines():
                    nm = ln.strip()
                    if nm and nm not in data:
                        data[nm] = None
                        agregar_maquina(nm)
    except Exception:
        pass

    out = []
    for nm in sorted(data):
        r = data[nm]
        if r is None:
            out.append({"maquina": nm, "estado": "sin_respaldo", "cuando": None,
                        "archivos": None, "correos": None, "errores": [],
                        "red_ok": red_ok})
        else:
            ok = r.get("status") == "ok"
            out.append({"maquina": nm, "estado": "ok" if ok else "error",
                        "cuando": r.get("when"), "archivos": r.get("files"),
                        "correos": r.get("mails"), "errores": r.get("errors", []),
                        "red_ok": red_ok})
    return out, red_ok


def get_maquinas():
    """
    Fuente de verdad: la lista guardada en el config LOCAL de PC100.
    Nunca depende de la red.
    """
    return [m.strip() for m in CFG.get("maquinas", []) if m.strip()]


def set_maquinas(lista):
    """
    Guarda la lista en el config local Y en _esperadas.txt de la carpeta compartida
    (para que el agente también la vea si la necesita en el futuro).
    """
    CFG["maquinas"] = [m.strip() for m in lista if m.strip()]
    save_cfg(CFG)
    # Sincronizar con el archivo de red (si está disponible); fallo silencioso
    try:
        fp = estado_dir() / "_esperadas.txt"
        fp.parent.mkdir(parents=True, exist_ok=True)
        fp.write_text("\n".join(CFG["maquinas"]) + "\n", encoding="utf-8")
    except Exception:
        pass


def esperadas_texto():
    return "\n".join(get_maquinas())


def guardar_esperadas(texto):
    lista = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    set_maquinas(lista)


def agregar_maquina(nombre):
    """Agrega una máquina a la lista si no estaba ya. Llamado desde la API."""
    nombre = nombre.strip()
    if not nombre:
        return
    lista = get_maquinas()
    if nombre not in lista:
        lista.append(nombre)
        set_maquinas(lista)


# ------------------------------------------------------------------ plantilla visual (misma paleta que el agente)
BASE_CSS = """
:root{
  --bg:#F2F6FC; --card:#FFFFFF; --primary:#2563EB; --primary-dk:#1D4ED8;
  --green:#16A34A; --green-bg:#DCFCE7; --orange:#D97706; --orange-bg:#FEF3C7;
  --red:#DC2626; --red-bg:#FEE2E2; --text:#1F2937; --muted:#6B7280; --border:#D7E2F2;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font-family:"Segoe UI",Arial,sans-serif}
header{display:flex;align-items:center;justify-content:space-between;padding:18px 28px;background:var(--card);
       border-bottom:1px solid var(--border)}
header h1{font-size:20px;margin:0;color:var(--primary-dk)}
header a.logout{color:var(--muted);text-decoration:none;font-size:13px}
.wrap{max-width:1100px;margin:0 auto;padding:22px 20px}
.toolbar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:16px}
.toolbar input, .toolbar select{padding:8px 10px;border:1px solid var(--border);border-radius:8px;font-size:14px}
.btn{display:inline-block;padding:9px 16px;border-radius:8px;border:none;cursor:pointer;font-size:14px;font-weight:600;
     text-decoration:none;color:var(--primary-dk);background:#E5EDFB}
.btn:hover{background:#D3E1F8}
.btn-primary{background:var(--primary);color:#fff}
.btn-primary:hover{background:var(--primary-dk)}
.summary{font-size:14px;color:var(--muted);margin-bottom:14px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px}
.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:16px;
      border-left:6px solid var(--muted)}
.card.ok{border-left-color:var(--green)}
.card.error{border-left-color:var(--orange)}
.card.sin_respaldo{border-left-color:var(--red)}
.card h3{margin:0 0 4px 0;font-size:16px}
.badge{display:inline-block;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:700;margin-bottom:8px}
.badge.ok{background:var(--green-bg);color:var(--green)}
.badge.error{background:var(--orange-bg);color:var(--orange)}
.badge.sin_respaldo{background:var(--red-bg);color:var(--red)}
.meta{font-size:13px;color:var(--muted);margin:2px 0}
.errores{font-size:13px;color:var(--red);margin-top:8px;white-space:pre-wrap}
.card .acciones{margin-top:12px;display:flex;gap:8px}
.login-box{max-width:360px;margin:90px auto;background:var(--card);padding:28px;border-radius:12px;
           border:1px solid var(--border)}
.login-box h2{margin-top:0;color:var(--primary-dk)}
.login-box input{width:100%;padding:10px;margin:8px 0;border:1px solid var(--border);border-radius:8px}
.login-box .error{color:var(--red);font-size:13px}
textarea{width:100%;min-height:220px;padding:10px;border:1px solid var(--border);border-radius:8px;font-family:monospace}
.toast{position:fixed;bottom:20px;right:20px;background:var(--text);color:#fff;padding:10px 16px;border-radius:8px;
       font-size:14px;opacity:0;transition:opacity .3s}
.toast.show{opacity:1}
"""

LOGIN_HTML = """
<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Panel del administrador — Respaldo Mensual</title><style>{{css}}</style></head>
<body>
  <div class="login-box">
    <h2>🔒 Panel del administrador</h2>
    <p style="color:var(--muted);font-size:14px">Acceso restringido al encargado del respaldo.</p>
    {% if error %}<p class="error">{{ error }}</p>{% endif %}
    <form method="post">
      <input type="password" name="password" placeholder="Contraseña" autofocus>
      <button class="btn btn-primary" style="width:100%" type="submit">Entrar</button>
    </form>
  </div>
</body></html>
"""

SETUP_HTML = """
<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Configuración inicial</title><style>{{css}}</style></head>
<body>
  <div class="login-box" style="max-width:460px">
    <h2>⚙️ Configuración inicial</h2>
    <p style="color:var(--muted);font-size:14px">Esto solo se pide una vez, en esta PC100.</p>
    {% if error %}<p class="error">{{ error }}</p>{% endif %}
    <form method="post">
      <label style="font-size:13px;color:var(--muted)">Carpeta local de los respaldos (ej. D:\\Respaldos):</label>
      <input type="text" name="destino_root" value="{{ destino_root }}" placeholder="D:\\Respaldos">
      <label style="font-size:13px;color:var(--muted)">Contraseña del panel:</label>
      <input type="password" name="password" placeholder="Contraseña nueva">
      <button class="btn btn-primary" style="width:100%" type="submit">Guardar y continuar</button>
    </form>
  </div>
</body></html>
"""

DASH_HTML = """
<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Panel del administrador — Respaldo Mensual</title><style>{{css}}
.aviso-red{background:#FEF3C7;border:1px solid #F59E0B;color:#92400E;padding:10px 16px;
           border-radius:8px;margin-bottom:14px;font-size:14px;display:none}
.btn-quitar{background:none;border:none;color:var(--muted);cursor:pointer;
            font-size:18px;float:right;line-height:1;padding:0 2px}
.btn-quitar:hover{color:var(--red)}
.agregar-box{display:flex;gap:8px;margin-bottom:14px}
.agregar-box input{flex:1;padding:8px 10px;border:1px solid var(--border);border-radius:8px;font-size:14px}
</style></head>
<body>
<header>
  <h1>\U0001f5c2\ufe0f Respaldo Mensual \u2014 Panel del administrador</h1>
  <div>
    <a class="logout" href="{{ url_for('config_view') }}" style="margin-right:16px">\u2699\ufe0f Configuraci\u00f3n</a>
    <a class="logout" href="{{ url_for('logout') }}">Cerrar sesi\u00f3n</a>
  </div>
</header>
<div class="wrap">
  <div class="toolbar">
    <label>Mes: <input type="month" id="mes" value="{{ mes }}"></label>
    <input type="text" id="buscar" placeholder="\U0001f50d Buscar m\u00e1quina\u2026">
    <button class="btn" onclick="cargar()">&#x21bb; Actualizar</button>
    <a class="btn" href="{{ url_for('esperadas_view') }}">&#x1f4cb; Editar lista completa</a>
  </div>
  <div class="agregar-box">
    <input type="text" id="nueva-maquina" placeholder="Agregar m\u00e1quina, ej. LAPC 46">
    <button class="btn btn-primary" onclick="agregarMaquina()">+ Agregar</button>
  </div>
  <div class="aviso-red" id="aviso-red">
    &#x26a0; La carpeta de respaldos no est\u00e1 disponible en este momento.
    Las m\u00e1quinas registradas siguen apareciendo pero su estado puede estar desactualizado.
  </div>
  <div class="summary" id="resumen">Cargando\u2026</div>
  <div class="grid" id="grid"></div>
</div>
<div class="toast" id="toast"></div>
<script>
function toast(msg, color){
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.style.background = color || '#1F2937';
  t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); }, 3000);
}
function agregarMaquina(){
  var nombre = document.getElementById('nueva-maquina').value.trim();
  if(!nombre){ toast('Escribe el nombre de la m\u00e1quina primero.', '#DC2626'); return; }
  fetch('/api/maquinas', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({nombre: nombre})
  }).then(function(r){ return r.json(); }).then(function(d){
    if(d.ok){
      document.getElementById('nueva-maquina').value = '';
      toast('M\u00e1quina "' + nombre + '" agregada.', '#16A34A');
      cargar();
    } else {
      toast('No se pudo agregar.', '#DC2626');
    }
  });
}
function quitarMaquina(nombre){
  if(!confirm('\u00bfQuitar "' + nombre + '" de la lista? No se borran sus archivos de respaldo.')){ return; }
  fetch('/api/maquinas', {
    method:'DELETE', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({nombre: nombre})
  }).then(function(r){ return r.json(); }).then(function(d){
    if(d.ok){ toast('M\u00e1quina eliminada de la lista.', '#D97706'); cargar(); }
  });
}
function respaldarAhora(nombre){
  var mes = document.getElementById('mes').value;
  fetch('/api/respaldar/' + encodeURIComponent(nombre), {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({mes: mes})
  }).then(function(r){ return r.json(); }).then(function(d){
    toast(d.ok ? 'Solicitud enviada. Se ejecutar\u00e1 en los pr\u00f3ximos minutos.' : 'Error: ' + (d.error || ''), d.ok ? '#16A34A' : '#DC2626');
  });
}
function badgeTexto(e){
  if(e==='ok')    return '\u2714 Respaldada';
  if(e==='error') return '\u26a0 Con errores';
  return '\u2716 Sin respaldo';
}
function cargar(){
  var mes = document.getElementById('mes').value;
  var q   = document.getElementById('buscar').value.toLowerCase();
  fetch('/api/estado?mes=' + mes)
    .then(function(r){ return r.json(); })
    .then(function(resp){
      var lista  = resp.maquinas || [];
      var red_ok = resp.red_ok !== false;
      document.getElementById('aviso-red').style.display = red_ok ? 'none' : 'block';
      var grid = document.getElementById('grid');
      grid.innerHTML = '';
      var ok = 0, total = lista.length;
      lista.forEach(function(m){
        if(m.estado==='ok') ok++;
        if(q && m.maquina.toLowerCase().indexOf(q)===-1) return;
        var div = document.createElement('div');
        div.className = 'card ' + m.estado;
        var errHtml = '';
        if(m.errores && m.errores.length){
          errHtml = '<div class="errores">' + m.errores.slice(0,5).join('<br>') + '</div>';
        }
        div.innerHTML =
          '<button class="btn-quitar" title="Quitar de la lista">\u2715</button>' +
          '<span class="badge ' + m.estado + '">' + badgeTexto(m.estado) + '</span>' +
          '<h3>' + m.maquina + '</h3>' +
          '<div class="meta">\u00daltima vez: ' + (m.cuando||'\u2014') + '</div>' +
          '<div class="meta">Archivos: ' + (m.archivos==null?'\u2014':m.archivos) + ' &nbsp; Correos: ' + (m.correos==null?'\u2014':m.correos) + '</div>' +
          errHtml +
          '<div class="acciones">' +
          '<button class="btn btn-primary">&#x25b6; Respaldar ahora</button>' +
          '<a class="btn" href="/correos/' + encodeURIComponent(m.maquina) + '/' + mes + '">&#x1f4e7; Ver correos</a>' +
          '</div>';
        var nombre = m.maquina;
        div.querySelector('.btn-quitar').onclick   = function(){ quitarMaquina(nombre); };
        div.querySelector('.btn-primary').onclick  = function(){ respaldarAhora(nombre); };
        grid.appendChild(div);
      });
      document.getElementById('resumen').textContent = ok + ' de ' + total + ' m\u00e1quinas respaldadas en ' + mes + '.';
    })
    .catch(function(err){
      document.getElementById('aviso-red').style.display = 'block';
      document.getElementById('resumen').textContent = 'No se pudo conectar. Recarga la página.';
    });
}
document.getElementById('buscar').addEventListener('input', cargar);
document.getElementById('mes').addEventListener('change', cargar);
document.getElementById('nueva-maquina').addEventListener('keydown', function(e){ if(e.key==='Enter') agregarMaquina(); });
cargar();
setInterval(cargar, 30000);
</script>
</body></html>
"""

ESPERADAS_HTML = """
<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Máquinas esperadas</title><style>{{css}}</style></head>
<body>
<header><h1>📋 Máquinas esperadas</h1><a class="logout" href="{{ url_for('dashboard') }}">← Volver</a></header>
<div class="wrap">
  <p style="color:var(--muted)">Una por línea. Las que estén aquí y no reporten respaldo aparecerán en rojo.</p>
  <form method="post">
    <textarea name="lista">{{ lista }}</textarea><br><br>
    <button class="btn btn-primary" type="submit">Guardar</button>
  </form>
</div>
</body></html>
"""


def logueado():
    return session.get("ok") is True



@app.route("/api/diagnostico")
def api_diagnostico():
    d = estado_dir()
    return jsonify({
        "panel_ok": True,
        "destino_root": CFG.get("destino_root", "(no configurado)"),
        "carpeta_estado_existe": d.is_dir(),
        "maquinas_registradas": len(get_maquinas()),
        "maquinas": get_maquinas(),
        "sesion_activa": logueado(),
    })

@app.route("/login", methods=["GET", "POST"])
def login():
    if not CFG.get("destino_root"):
        return redirect(url_for("setup"))
    error = None
    if request.method == "POST":
        pw = request.form.get("password", "")
        if hashlib.sha256(pw.encode()).hexdigest() == CFG["password_hash"]:
            session["ok"] = True
            return redirect(url_for("dashboard"))
        error = "Contraseña incorrecta."
    return render_template_string(LOGIN_HTML, css=Markup(BASE_CSS), error=error)


@app.route("/setup", methods=["GET", "POST"])
def setup():
    error = None
    if request.method == "POST":
        destino = request.form.get("destino_root", "").strip()
        pw = request.form.get("password", "").strip()
        if not destino:
            error = "Indica la carpeta local de los respaldos."
        elif not pw:
            error = "Define una contraseña."
        else:
            CFG["destino_root"] = destino
            CFG["password_hash"] = hashlib.sha256(pw.encode()).hexdigest()
            save_cfg(CFG)
            return redirect(url_for("login"))
    return render_template_string(SETUP_HTML, css=Markup(BASE_CSS), error=error, destino_root=CFG.get("destino_root", ""))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def dashboard():
    if not CFG.get("destino_root"):
        return redirect(url_for("setup"))
    if not logueado():
        return redirect(url_for("login"))
    return render_template_string(DASH_HTML, css=Markup(BASE_CSS), mes=prev_month_str())


@app.route("/api/estado")
def api_estado():
    if not logueado():
        return jsonify({"error": "no autorizado", "maquinas": [], "red_ok": False}), 401
    tag = request.args.get("mes", prev_month_str())
    maquinas, red_ok = leer_estado(tag)
    return jsonify({"maquinas": maquinas, "red_ok": red_ok})


@app.route("/api/respaldar/<maquina>", methods=["POST"])
def api_respaldar(maquina):
    if not logueado():
        return jsonify({"ok": False}), 401
    mes = (request.get_json(silent=True) or {}).get("mes", prev_month_str())
    try:
        rf = solicitudes_dir() / f"{maquina.strip()}.json"
        rf.parent.mkdir(parents=True, exist_ok=True)
        rf.write_text(json.dumps({"accion": "respaldar_ahora", "mes": mes,
                                   "pedido_el": datetime.now().isoformat(timespec="seconds")}),
                      encoding="utf-8")
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/maquinas", methods=["GET", "POST", "DELETE"])
def api_maquinas():
    """API para agregar o quitar máquinas individuales desde el dashboard."""
    if not logueado():
        return jsonify({"ok": False}), 401
    if request.method == "GET":
        return jsonify({"maquinas": get_maquinas()})
    data = request.get_json(silent=True) or {}
    nombre = data.get("nombre", "").strip()
    if not nombre:
        return jsonify({"ok": False, "error": "nombre vacío"}), 400
    if request.method == "POST":
        agregar_maquina(nombre)
    elif request.method == "DELETE":
        lista = [m for m in get_maquinas() if m != nombre]
        set_maquinas(lista)
    return jsonify({"ok": True, "maquinas": get_maquinas()})


@app.route("/esperadas", methods=["GET", "POST"])
def esperadas_view():
    if not logueado():
        return redirect(url_for("login"))
    if request.method == "POST":
        guardar_esperadas(request.form.get("lista", ""))
        return redirect(url_for("dashboard"))
    return render_template_string(ESPERADAS_HTML, css=Markup(BASE_CSS), lista=esperadas_texto())


@app.route("/config", methods=["GET", "POST"])
def config_view():
    if not logueado():
        return redirect(url_for("login"))
    error = None
    if request.method == "POST":
        destino = request.form.get("destino_root", "").strip()
        pw = request.form.get("password", "").strip()
        if destino:
            CFG["destino_root"] = destino
        if pw:
            CFG["password_hash"] = hashlib.sha256(pw.encode()).hexdigest()
        save_cfg(CFG)
        return redirect(url_for("dashboard"))
    return render_template_string(SETUP_HTML, css=Markup(BASE_CSS), error=error, destino_root=CFG.get("destino_root", ""))


# ──────────────────────────────────────────────── visor de correos ──
import email as _email_mod
import email.header as _email_hdr
import html as _html_mod
import io


def _correos_root(maquina, mes):
    return Path(CFG["destino_root"]) / maquina.strip() / mes.strip() / "Correos"


def _parse_header(value):
    """Decodifica encabezados MIME (Base64/QP) a texto legible."""
    if not value:
        return ""
    try:
        parts = _email_hdr.decode_header(value)
        return str(_email_hdr.make_header(parts))
    except Exception:
        return str(value)


def _list_emls(maquina, mes):
    """
    Devuelve lista de dicts con los metadatos de cada .eml
    (carpeta, nombre, De, Para, Asunto, Fecha) sin leer el cuerpo.
    """
    root = _correos_root(maquina, mes)
    result = []
    if not root.is_dir():
        return result
    for carpeta_user in sorted(root.iterdir()):
        if not carpeta_user.is_dir():
            continue
        for categoria in ("Recibidos", "Enviados"):
            cat_dir = carpeta_user / categoria
            if not cat_dir.is_dir():
                continue
            for eml in sorted(cat_dir.glob("*.eml"), reverse=True):
                try:
                    raw = eml.read_bytes()
                    msg = _email_mod.message_from_bytes(raw)
                    result.append({
                        "carpeta": categoria,
                        "cuenta": carpeta_user.name,
                        "archivo": eml.name,
                        "de": _parse_header(msg.get("From", "")),
                        "para": _parse_header(msg.get("To", "")),
                        "asunto": _parse_header(msg.get("Subject", "(sin asunto)")),
                        "fecha": _parse_header(msg.get("Date", "")),
                        "size": eml.stat().st_size,
                    })
                except Exception:
                    pass
    return result


def _render_eml_body(msg):
    """
    Extrae el cuerpo del correo como HTML seguro para mostrarlo en el navegador.
    Prefiere la parte HTML; cae a texto plano si no hay HTML.
    """
    def get_charset(part):
        cs = part.get_content_charset()
        return cs if cs else "utf-8"

    html_body = plain_body = None
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/html" and html_body is None:
                try:
                    html_body = part.get_payload(decode=True).decode(get_charset(part), errors="replace")
                except Exception:
                    pass
            elif ct == "text/plain" and plain_body is None:
                try:
                    plain_body = part.get_payload(decode=True).decode(get_charset(part), errors="replace")
                except Exception:
                    pass
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                text = payload.decode(get_charset(msg), errors="replace")
                if msg.get_content_type() == "text/html":
                    html_body = text
                else:
                    plain_body = text
        except Exception:
            pass

    if html_body:
        # Eliminar <script> y <style> externos para seguridad básica
        import re as _re
        html_body = _re.sub(r'<script[\s\S]*?</script>', '', html_body, flags=_re.IGNORECASE)
        html_body = _re.sub(r'<link[^>]+>', '', html_body, flags=_re.IGNORECASE)
        return html_body, "html"
    if plain_body:
        return "<pre style='white-space:pre-wrap;font-family:inherit'>" + _html_mod.escape(plain_body) + "</pre>", "plain"
    return "<p style='color:#888'>Sin contenido visible.</p>", "empty"


def _list_attachments(msg):
    """Lista de adjuntos con nombre y tamaño."""
    atts = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                name = _parse_header(part.get_filename() or "adjunto")
                payload = part.get_payload(decode=True) or b""
                atts.append({"nombre": name, "size": len(payload)})
    return atts


@app.route("/correos/<maquina>/<mes>")
def correos_lista(maquina, mes):
    if not logueado():
        return redirect(url_for("login"))
    emls = _list_emls(maquina, mes)
    filas = ""
    for e in emls:
        url = url_for("correos_ver", maquina=maquina, mes=mes,
                      cuenta=e["cuenta"], carpeta=e["carpeta"], archivo=e["archivo"])
        badge = ("#16A34A" if e["carpeta"] == "Recibidos" else "#2563EB")
        filas += (
            f'<tr onclick="location.href=\'{url}\'" style="cursor:pointer">'
            f'<td><span style="background:{badge};color:#fff;padding:2px 8px;border-radius:99px;'
            f'font-size:11px">{_html_mod.escape(e["carpeta"])}</span></td>'
            f'<td>{_html_mod.escape(e["asunto"])}</td>'
            f'<td style="color:#555;font-size:13px">{_html_mod.escape(e["de"][:60])}</td>'
            f'<td style="color:#555;font-size:13px;white-space:nowrap">{_html_mod.escape(e["fecha"][:20])}</td>'
            f'</tr>'
        )
    volver = url_for("dashboard")
    html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Correos — {_html_mod.escape(maquina)} {_html_mod.escape(mes)}</title>
<style>{BASE_CSS}
table{{width:100%;border-collapse:collapse}}
th{{text-align:left;padding:8px 12px;background:#EFF6FF;color:#1D4ED8;font-size:13px}}
td{{padding:9px 12px;border-bottom:1px solid #E5EDFB;font-size:14px;max-width:300px;
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
tr:hover td{{background:#F0F7FF}}
.num{{color:#6B7280;font-size:13px}}
</style></head><body>
<header>
  <h1 style="margin:0">📧 Correos — {_html_mod.escape(maquina)} / {_html_mod.escape(mes)}</h1>
  <a class="logout" href="{volver}">← Volver al panel</a>
</header>
<div class="wrap">
  <p class="num">{len(emls)} correo(s) respaldado(s)</p>
  <table>
    <thead><tr><th>Carpeta</th><th>Asunto</th><th>De</th><th>Fecha</th></tr></thead>
    <tbody>{filas if filas else '<tr><td colspan="4" style="color:#888;padding:20px">No hay correos respaldados para este mes.</td></tr>'}</tbody>
  </table>
</div></body></html>"""
    return html


@app.route("/correos/<maquina>/<mes>/<cuenta>/<carpeta>/<archivo>")
def correos_ver(maquina, mes, cuenta, carpeta, archivo):
    if not logueado():
        return redirect(url_for("login"))
    # Validar que el archivo esté dentro de la ruta esperada (sin path traversal)
    root = _correos_root(maquina, mes)
    eml_path = (root / cuenta / carpeta / archivo).resolve()
    if not str(eml_path).startswith(str(root.resolve())):
        return "Ruta no permitida", 403
    if not eml_path.exists() or eml_path.suffix.lower() != ".eml":
        return "Correo no encontrado", 404
    try:
        raw = eml_path.read_bytes()
        msg = _email_mod.message_from_bytes(raw)
    except Exception as e:
        return f"Error al leer el correo: {e}", 500

    asunto = _parse_header(msg.get("Subject", "(sin asunto)"))
    de = _parse_header(msg.get("From", ""))
    para = _parse_header(msg.get("To", ""))
    cc = _parse_header(msg.get("Cc", ""))
    fecha = _parse_header(msg.get("Date", ""))
    body_html, body_type = _render_eml_body(msg)
    adjuntos = _list_attachments(msg)
    lista_url = url_for("correos_lista", maquina=maquina, mes=mes)

    adj_html = ""
    if adjuntos:
        items = "".join(f'<li>📎 {_html_mod.escape(a["nombre"])} ({a["size"]//1024} KB)</li>'
                        for a in adjuntos)
        adj_html = f'<div style="background:#FEF3C7;padding:10px 16px;border-radius:8px;margin:12px 0"><b>Adjuntos:</b><ul style="margin:4px 0 0 16px">{items}</ul></div>'

    html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>{_html_mod.escape(asunto)}</title>
<style>{BASE_CSS}
.cabecera{{background:#fff;border:1px solid #D7E2F2;border-radius:10px;
           padding:16px 20px;margin-bottom:16px}}
.cabecera .etiq{{color:#6B7280;font-size:12px;width:60px;display:inline-block}}
.cabecera .val{{font-size:14px}}
.cuerpo{{background:#fff;border:1px solid #D7E2F2;border-radius:10px;
         padding:20px;min-height:200px;overflow-x:auto}}
</style></head><body>
<header>
  <h1 style="margin:0;font-size:17px">📧 {_html_mod.escape(asunto)}</h1>
  <a class="logout" href="{lista_url}">← Volver a la lista</a>
</header>
<div class="wrap">
  <div class="cabecera">
    <div><span class="etiq">De:</span> <span class="val">{_html_mod.escape(de)}</span></div>
    <div><span class="etiq">Para:</span> <span class="val">{_html_mod.escape(para)}</span></div>
    {'<div><span class="etiq">CC:</span> <span class="val">' + _html_mod.escape(cc) + '</span></div>' if cc else ''}
    <div><span class="etiq">Fecha:</span> <span class="val">{_html_mod.escape(fecha)}</span></div>
    <div><span class="etiq">Carpeta:</span> <span class="val">{_html_mod.escape(carpeta)}</span></div>
  </div>
  {adj_html}
  <div class="cuerpo">{body_html}</div>
</div></body></html>"""
    return html


@app.route("/api/progreso")
def api_progreso():
    """
    Devuelve el progreso en tiempo real de todas las máquinas que están
    respaldando ahora mismo. Actualiza solo las que tienen archivo _progreso.json
    con menos de 5 minutos de antigüedad.
    """
    if not logueado():
        return jsonify({}), 401
    d = estado_dir()
    result = {}
    STALE_SECS = 5 * 60   # si no actualizó en 5 min, el proceso se interrumpió
    try:
        if not d.is_dir():
            return jsonify({})
        for fp in d.glob("*_progreso.json"):
            machine = fp.stem.replace("_progreso", "")
            try:
                data = json.loads(fp.read_text(encoding="utf-8"))
                # Verificar antigüedad
                upd = datetime.strptime(data["actualizado"], "%Y-%m-%d %H:%M:%S")
                age = (datetime.now() - upd).total_seconds()
                if age > STALE_SECS:
                    data["etapa"] = "interrumpido"
                    data["detalle"] = f"Sin actualización hace {int(age//60)} min"
                data["segundos_desde_update"] = int(age)
                result[machine] = data
            except Exception:
                pass
    except Exception:
        pass
    return jsonify(result)


@app.route("/api/correos/<maquina>/<mes>")
def api_correos(maquina, mes):
    """API JSON: lista de correos de una máquina y mes para usar desde el dashboard."""
    if not logueado():
        return jsonify({"error": "no autorizado"}), 401
    emls = _list_emls(maquina, mes)
    return jsonify({"total": len(emls), "correos": emls})


if __name__ == "__main__":
    print("Panel del administrador — abre http://localhost:5000 en tu navegador")
    print(f"(o http://<IP de esta PC>:5000 desde otra computadora de la oficina)")
    app.run(host="0.0.0.0", port=5000, threaded=True)