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
import os, sys, json, hashlib, secrets, shutil, uuid, time
from werkzeug.security import generate_password_hash, check_password_hash
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
    "destino_root": "",
    "history_backup_root": "",
    "tls_cert": "",
    "tls_key": "",
    "admin_user": "admin",
    "readonly_user": "consulta",
    "password_hash": hashlib.sha256("admin".encode()).hexdigest(),
    "readonly_hash": "",   # contraseña del usuario de solo lectura (vacío = deshabilitado)
    "secret_key": secrets.token_hex(16),
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
_TLS_ACTIVE = bool(CFG.get("tls_cert") and CFG.get("tls_key") and Path(CFG["tls_cert"]).is_file() and Path(CFG["tls_key"]).is_file())
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Strict", SESSION_COOKIE_SECURE=_TLS_ACTIVE)


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
            ok = r.get("status") in ("ok", "verified")
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
      <input type="text" name="username" placeholder="Usuario" autofocus>
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
.prog-bar{height:4px;background:#E5EDFB;border-radius:99px;margin:6px 0}
.prog-bar-inner{height:4px;border-radius:99px;background:var(--primary);animation:pulse 1.5s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.reagendado{background:#FEF3C7;border-radius:6px;padding:4px 8px;font-size:12px;
            color:#92400E;margin-top:6px}
.readonly-banner{background:#EFF6FF;border:1px solid #BFDBFE;color:#1D4ED8;
                 padding:8px 16px;border-radius:8px;margin-bottom:12px;font-size:13px}
.btn:disabled,.btn-primary:disabled{opacity:.45;cursor:not-allowed;pointer-events:none}
</style></head>
<body>
<header>
  <h1>&#x1f5c2;&#xfe0f; Respaldo Mensual &mdash; Panel del administrador</h1>
  <div>
    {% if es_admin %}
    <a class="logout" href="{{ url_for('config_view') }}" style="margin-right:16px">&#x2699;&#xfe0f; Configuraci&oacute;n</a>
    {% endif %}
    <a class="logout" href="{{ url_for('logout') }}">Cerrar sesi&oacute;n</a>
  </div>
</header>
<div class="wrap">
  {% if not es_admin %}
  <div class="readonly-banner">&#x1f441; Modo solo lectura &mdash; puedes ver el estado pero no iniciar respaldos.</div>
  {% endif %}
  <div class="toolbar">
    <label>Mes: <input type="month" id="mes" value="{{ mes }}"></label>
    <input type="text" id="buscar" placeholder="&#x1f50d; Buscar m&aacute;quina&hellip;">
    <button class="btn" onclick="cargar()">&#x21bb; Actualizar</button>
    {% if es_admin %}
    <a class="btn" href="{{ url_for('esperadas_view') }}">&#x1f4cb; Editar lista completa</a>
    <button class="btn" onclick="verReporte()">&#x1f4c4; Reporte mensual</button>
    {% endif %}
  </div>
  {% if es_admin %}
  <div class="agregar-box">
    <input type="text" id="nueva-maquina" placeholder="Agregar m&aacute;quina, ej. LAPC 46">
    <button class="btn btn-primary" onclick="agregarMaquina()">+ Agregar</button>
  </div>
  {% endif %}
  <div class="aviso-red" id="aviso-red">
    &#x26a0; La carpeta de respaldos no est&aacute; disponible ahora.
    Las m&aacute;quinas registradas siguen apareciendo con su &uacute;ltimo estado conocido.
  </div>
  <div class="summary" id="resumen">Cargando&hellip;</div>
  <div class="grid" id="grid"></div>
</div>
<div class="toast" id="toast"></div>
<script>
var ES_ADMIN = {{ 'true' if es_admin else 'false' }};

function toast(msg, color){
  var t = document.getElementById('toast');
  t.textContent = msg; t.style.background = color || '#1F2937';
  t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); }, 3500);
}

function agregarMaquina(){
  var nombre = document.getElementById('nueva-maquina').value.trim();
  if(!nombre){ toast('Escribe el nombre de la m\u00e1quina primero.','#DC2626'); return; }
  fetch('/api/maquinas',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({nombre:nombre})})
  .then(function(r){ return r.json(); })
  .then(function(d){
    if(d.ok){ document.getElementById('nueva-maquina').value='';
              toast('M\u00e1quina "'+nombre+'" agregada.','#16A34A'); cargar(); }
    else toast('Error: '+(d.error||'no se pudo agregar.'),'#DC2626');
  });
}

function quitarMaquina(nombre){
  if(!confirm('\u00bfQuitar "'+nombre+'" de la lista? No se borran sus archivos.')){ return; }
  fetch('/api/maquinas',{method:'DELETE',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({nombre:nombre})})
  .then(function(r){ return r.json(); })
  .then(function(d){ if(d.ok){ toast('Quitada.','#D97706'); cargar(); }});
}

function respaldarAhora(nombre){
  var mes = document.getElementById('mes').value;
  fetch('/api/respaldar/'+encodeURIComponent(nombre),{method:'POST',
    headers:{'Content-Type':'application/json'},body:JSON.stringify({mes:mes})})
  .then(function(r){ return r.json(); })
  .then(function(d){
    toast(d.ok ? 'Solicitud enviada. Se ejecutar\u00e1 en los pr\u00f3ximos minutos.' : 'Error: '+(d.error||''),
          d.ok ? '#16A34A' : '#DC2626');
  });
}

function verReporte(){
  var mes = document.getElementById('mes').value;
  window.open('/api/reporte/'+mes,'_blank');
}

function etapaTexto(e){
  if(e==='archivos')     return '&#x1F4C2; Copiando archivos\u2026';
  if(e==='correos')      return '&#x1F4E7; Respaldando correos\u2026';
  if(e==='enviando')     return '&#x1F4E4; Enviando a PC100\u2026';
  if(e==='listo')        return '\u2714 Terminado';
  if(e==='error')        return '\u26a0 Termin\u00f3 con errores';
  if(e==='interrumpido') return '\u23f8 Interrumpido';
  return e;
}

function badgeTexto(e){
  if(e==='ok')    return '\u2714 Respaldada';
  if(e==='error') return '\u26a0 Con errores';
  return '\u2716 Sin respaldo';
}

function cargar(){
  var mes = document.getElementById('mes').value;
  var q   = document.getElementById('buscar').value.toLowerCase();
  fetch('/api/estado?mes='+mes)
  .then(function(r){ return r.json(); })
  .then(function(resp){
    var lista  = resp.maquinas || [];
    var red_ok = resp.red_ok !== false;
    document.getElementById('aviso-red').style.display = red_ok ? 'none' : 'block';
    var grid = document.getElementById('grid');
    grid.innerHTML = '';
    var ok=0, enProceso=0, total=lista.length;
    lista.forEach(function(m){
      if(m.estado==='ok') ok++;
      if(q && m.maquina.toLowerCase().indexOf(q)===-1) return;
      var div = document.createElement('div');
      div.className = 'card '+m.estado;

      // Progreso en tiempo real
      var progHtml = '';
      var enCurso = false;
      if(m.progreso){
        enCurso = true; enProceso++;
        var etapa = m.progreso.etapa || '';
        progHtml =
          '<div class="prog-bar"><div class="prog-bar-inner"></div></div>'+
          '<div style="font-size:12px;color:#2563EB">'+etapaTexto(etapa)+
          ' &nbsp;'+escHtml(m.progreso.detalle||'')+
          ' &nbsp;<span style="color:#6B7280">arch='+m.progreso.archivos+
          ' corr='+m.progreso.correos+'</span></div>';
      }

      // Reagendado
      var reagHtml = '';
      if(m.reagendado){
        var mot = m.reagendado.motivo==='reagendado_por_error' ? 'Reagendado por error' : 'Reagendado (sin respaldo)';
        var errs = (m.reagendado.errores||[]).slice(0,2).join('; ');
        reagHtml = '<div class="reagendado">&#x1f504; '+escHtml(mot)+(errs?' &mdash; '+escHtml(errs):'')+'</div>';
      }

      var errHtml = '';
      if(m.errores && m.errores.length){
        errHtml = '<div class="errores">'+m.errores.slice(0,3).map(escHtml).join('<br>')+'</div>';
      }

      var archivosTxt = (m.archivos==null) ? '-' : m.archivos;
      var correosTxt  = (m.correos==null)  ? '-' : m.correos;

      // Botón respaldar: deshabilitado si ya respaldada o en curso
      var yaBien = m.estado==='ok';
      var btnTxt = yaBien ? '\u2714 Ya respaldada' : (enCurso ? '\u23f3 En proceso\u2026' : '&#x25b6; Respaldar ahora');
      var btnDis = (!ES_ADMIN || yaBien || enCurso) ? 'disabled' : '';
      var corrLink = '/correos/'+encodeURIComponent(m.maquina)+'/'+mes;

      div.innerHTML =
        (ES_ADMIN ? '<button class="btn-quitar" title="Quitar">\u00d7</button>' : '')+
        '<span class="badge '+m.estado+'">'+badgeTexto(m.estado)+'</span>'+
        '<h3>'+escHtml(m.maquina)+'</h3>'+
        '<div class="meta">\u00daltima vez: '+(m.cuando||'\u2014')+'</div>'+
        '<div class="meta">Archivos: '+archivosTxt+' &nbsp; Correos: '+correosTxt+'</div>'+
        errHtml + reagHtml + progHtml +
        '<div class="acciones">'+
        '<button class="btn btn-primary" '+btnDis+'>'+btnTxt+'</button>'+
        '<a class="btn" href="'+corrLink+'">&#x1f4e7; Ver correos</a>'+
        '</div>';

      var nombre = m.maquina;
      if(ES_ADMIN){
        var bq = div.querySelector('.btn-quitar');
        if(bq) bq.onclick = function(){ quitarMaquina(nombre); };
      }
      var bp = div.querySelector('.btn-primary');
      if(bp && !bp.disabled) bp.onclick = function(){ respaldarAhora(nombre); };

      grid.appendChild(div);
    });

    var txt = ok+' de '+total+' m\u00e1quinas respaldadas en '+mes+'.';
    if(enProceso) txt += '  &#x23f3; '+enProceso+' respaldando ahora.';
    document.getElementById('resumen').innerHTML = txt;

    // Si hay máquinas en proceso, refrescar más rápido
    if(enProceso){
      clearTimeout(window._timer);
      window._timer = setTimeout(cargar, 8000);
    }
  })
  .catch(function(){
    document.getElementById('aviso-red').style.display='block';
    document.getElementById('resumen').textContent='No se pudo conectar. Recarga la p\u00e1gina.';
  });
}

function escHtml(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

document.getElementById('buscar').addEventListener('input', cargar);
document.getElementById('mes').addEventListener('change', cargar);
{% if es_admin %}
document.getElementById('nueva-maquina').addEventListener('keydown',
  function(e){ if(e.key==='Enter') agregarMaquina(); });
{% endif %}
cargar();
window._timer = setInterval(cargar, 30000);
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


def solo_lectura():
    """El usuario de solo lectura puede ver el estado pero no respaldar ni modificar."""
    return session.get("readonly") is True


def puede_ver():
    return logueado() or solo_lectura()



@app.route("/api/diagnostico")
def api_diagnostico():
    if not puede_ver():
        return jsonify({"error": "no autorizado"}), 401
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
        username = request.form.get("username", "").strip()
        pw = request.form.get("password", "")
        ph = hashlib.sha256(pw.encode()).hexdigest()
        stored_admin = CFG.get("password_hash", "")
        admin_ok = check_password_hash(stored_admin, pw) if stored_admin.startswith(("scrypt:", "pbkdf2:")) else ph == stored_admin
        if username == CFG.get("admin_user", "admin") and admin_ok:
            if not stored_admin.startswith(("scrypt:", "pbkdf2:")):
                CFG["password_hash"] = generate_password_hash(pw); save_cfg(CFG)
            session.clear(); session["ok"] = True; session["username"] = username
            _audit_event("admin_login")
            return redirect(url_for("dashboard"))
        ro = CFG.get("readonly_hash", "")
        reader_ok = (check_password_hash(ro, pw) if ro.startswith(("scrypt:", "pbkdf2:")) else ph == ro) if ro else False
        if username == CFG.get("readonly_user", "consulta") and reader_ok:
            if not ro.startswith(("scrypt:", "pbkdf2:")):
                CFG["readonly_hash"] = generate_password_hash(pw); save_cfg(CFG)
            session.clear(); session["readonly"] = True; session["username"] = username
            _audit_event("readonly_login")
            return redirect(url_for("dashboard"))
        error = "Usuario o contraseña incorrectos."
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
            CFG["password_hash"] = generate_password_hash(pw)
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
    if not puede_ver():
        return redirect(url_for("login"))
    es_admin = logueado()
    return render_template_string(DASH_HTML, css=Markup(BASE_CSS),
                                   mes=prev_month_str(), es_admin=es_admin)


@app.route("/api/estado")
def api_estado():
    if not puede_ver():
        return jsonify({"error": "no autorizado", "maquinas": [], "red_ok": False}), 401
    tag = request.args.get("mes", prev_month_str())
    maquinas, red_ok = leer_estado(tag)
    # Enriquecer con progreso en tiempo real si está disponible
    progreso = _leer_progreso()
    for m in maquinas:
        nm = m["maquina"]
        if nm in progreso:
            m["progreso"] = progreso[nm]
    # Enriquecer con info de reagendado
    reagendados = _leer_reagendados()
    for m in maquinas:
        nm = m["maquina"]
        if nm in reagendados:
            m["reagendado"] = reagendados[nm]
    return jsonify({"maquinas": maquinas, "red_ok": red_ok,
                    "es_admin": logueado()})


@app.route("/api/respaldar/<maquina>", methods=["POST"])
def api_respaldar(maquina):
    if not logueado():   # solo lectura NO puede pedir respaldos
        return jsonify({"ok": False, "error": "Solo el administrador puede pedir respaldos"}), 403
    mes = str((request.get_json(silent=True) or {}).get("mes", prev_month_str())).strip()
    if not maquina.strip() or any(x in maquina for x in ("/", "\\", "..")):
        return jsonify({"ok": False, "error": "Máquina inválida."}), 400
    try:
        if len(mes) != 7 or mes[4] != "-" or not mes[:4].isdigit() or not mes[5:].isdigit() or not 1 <= int(mes[5:]) <= 12:
            return jsonify({"ok": False, "error": "Período inválido. Usa AAAA-MM."}), 400
        if (int(mes[:4]), int(mes[5:])) >= (date.today().year, date.today().month):
            return jsonify({"ok": False, "error": "El respaldo manual debe ser de un mes anterior."}), 400
        rf = solicitudes_dir() / f"{maquina.strip()}.json"
        rf.parent.mkdir(parents=True, exist_ok=True)
        rf.write_text(json.dumps({"accion": "respaldar_ahora", "mes": mes,
                                   "requested_by": session.get("username", "admin"),
                                   "pedido_el": datetime.now().isoformat(timespec="seconds")}),
                      encoding="utf-8")
        _audit_event("manual_backup_requested", {"machine": maquina, "period": mes})
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/maquinas", methods=["GET", "POST", "DELETE"])
def api_maquinas():
    """API para agregar o quitar máquinas individuales desde el dashboard."""
    if not logueado():
        return jsonify({"ok": False}), 403
    if request.method == "GET":
        return jsonify({"maquinas": get_maquinas()})
    data = request.get_json(silent=True) or {}
    nombre = data.get("nombre", "").strip()
    if not nombre:
        return jsonify({"ok": False, "error": "nombre vacío"}), 400
    if request.method == "POST":
        # Verificar duplicados (insensible a mayúsculas)
        existentes = [m.lower() for m in get_maquinas()]
        if nombre.lower() in existentes:
            return jsonify({"ok": False, "error": f"Ya existe una máquina llamada '{nombre}'"}), 409
        agregar_maquina(nombre)
    elif request.method == "DELETE":
        lista = [m for m in get_maquinas() if m.lower() != nombre.lower()]
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
        history_backup_root = request.form.get("history_backup_root", "").strip()
        tls_cert = request.form.get("tls_cert", "").strip()
        tls_key = request.form.get("tls_key", "").strip()
        pw = request.form.get("password", "").strip()
        pw_ro = request.form.get("password_ro", "").strip()
        admin_user = request.form.get("admin_user", "").strip()
        if admin_user:
            CFG["admin_user"] = admin_user
        if destino:
            CFG["destino_root"] = destino
        if history_backup_root:
            CFG["history_backup_root"] = history_backup_root
        if tls_cert:
            CFG["tls_cert"] = tls_cert
        if tls_key:
            CFG["tls_key"] = tls_key
        readonly_user = request.form.get("readonly_user", "").strip()
        if readonly_user:
            CFG["readonly_user"] = readonly_user
        if pw:
            CFG["password_hash"] = generate_password_hash(pw)
        if pw_ro:
            CFG["readonly_hash"] = generate_password_hash(pw_ro)
        elif request.form.get("disable_ro"):
            CFG["readonly_hash"] = ""
        save_cfg(CFG)
        _audit_event("settings_changed", {"admin_user_changed": bool(admin_user), "credentials_changed": bool(pw or pw_ro)})
        return redirect(url_for("dashboard"))
    # Plantilla extendida con campo para usuario de solo lectura
    cfg_html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Configuración</title><style>{BASE_CSS}</style></head><body>
<header><h1 style="margin:0">⚙️ Configuración del panel</h1>
<a class="logout" href="{url_for('dashboard')}">← Volver</a></header>
<div class="wrap">
  <div style="background:#fff;border:1px solid #D7E2F2;border-radius:12px;padding:24px;max-width:500px">
    {'<p style="color:var(--red)">' + error + '</p>' if error else ''}
    <form method="post">
      <label style="font-size:13px;color:var(--muted)">Carpeta local de los respaldos:</label>
      <input type="text" name="destino_root" value="{CFG.get('destino_root','')}"
             style="width:100%;padding:9px;border:1px solid #D7E2F2;border-radius:8px;margin:6px 0 14px">
      <label style="font-size:13px;color:var(--muted)">Destino independiente para copia diaria del historial:</label>
      <input type="text" name="history_backup_root" value="{CFG.get('history_backup_root','')}"
             style="width:100%;padding:9px;border:1px solid #D7E2F2;border-radius:8px;margin:6px 0 14px">

       <label style="font-size:13px;color:var(--muted)">Certificado TLS (.crt/.pem):</label>
      <input type="text" name="tls_cert" value="{CFG.get('tls_cert','')}" style="width:100%;padding:9px;margin:6px 0 14px">
      <label style="font-size:13px;color:var(--muted)">Clave TLS (.key):</label>
      <input type="text" name="tls_key" value="{CFG.get('tls_key','')}" style="width:100%;padding:9px;margin:6px 0 14px">
      <label style="font-size:13px;color:var(--muted)">Usuario administrador:</label>
       <input type="text" name="admin_user" value="{CFG.get('admin_user', 'admin')}"
              style="width:100%;padding:9px;border:1px solid #D7E2F2;border-radius:8px;margin:6px 0 14px"> 

      <label style="font-size:13px;color:var(--muted)">Nueva contraseña del administrador (vacío = no cambiar):</label>
      <input type="password" name="password"
             style="width:100%;padding:9px;border:1px solid #D7E2F2;border-radius:8px;margin:6px 0 14px">
      <hr style="border:none;border-top:1px solid #E5EDFB;margin:8px 0 14px">
      <label style="font-size:13px;color:var(--muted)">
        Contraseña usuario de <b>solo lectura</b> (vacío = no cambiar):<br>
        <span style="font-weight:normal">Este usuario solo puede ver el estado, no respaldar ni modificar nada.</span>
      </label>
       <label style="font-size:13px;color:var(--muted)">Usuario de solo lectura:</label>
       <input type="text" name="readonly_user" value="{CFG.get('readonly_user', 'consulta')}"
              style="width:100%;padding:9px;border:1px solid #D7E2F2;border-radius:8px;margin:6px 0 14px">
      <input type="password" name="password_ro"
             style="width:100%;padding:9px;border:1px solid #D7E2F2;border-radius:8px;margin:6px 0 8px">
      {'<label style="font-size:12px;color:var(--muted)"><input type="checkbox" name="disable_ro"> Deshabilitar usuario de solo lectura</label><br><br>' if CFG.get("readonly_hash") else ''}
      <button class="btn btn-primary" type="submit" style="width:100%">Guardar</button>
    </form>
  </div>
</div></body></html>"""
    return cfg_html


@app.route("/api/reporte/<mes>")
def api_reporte(mes):
    """Devuelve el reporte mensual en texto plano si existe."""
    if not puede_ver():
        return "No autorizado", 403
    rp = estado_dir() / f"reporte_{mes}.txt"
    if not rp.exists():
        return f"No hay reporte para {mes} todavía.", 404
    return rp.read_text(encoding="utf-8"), 200, {"Content-Type": "text/plain; charset=utf-8"}


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
    for eml in sorted(root.rglob("*.eml"), key=lambda p: p.stat().st_mtime, reverse=True):
            try:
                raw = eml.read_bytes()
                msg = _email_mod.message_from_bytes(raw)
                result.append({
                    "carpeta": ("Enviados" if "enviados" in [p.casefold() for p in eml.relative_to(root).parts] else
                                "Recibidos" if "recibidos" in [p.casefold() for p in eml.relative_to(root).parts] else "Personalizada"),
                    "cuenta": next((p for p in eml.relative_to(root).parts if p.startswith("Correo ")), "Cuenta"),
                    "ruta": eml.relative_to(root).as_posix(),
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
    if not puede_ver():
        return redirect(url_for("login"))
    emls = _list_emls(maquina, mes)
    filas = ""
    for e in emls:
        url = url_for("correos_ver", maquina=maquina, mes=mes, relative_eml=e["ruta"])
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


@app.route("/correos/<maquina>/<mes>/<path:relative_eml>")
def correos_ver(maquina, mes, relative_eml):
    if not puede_ver():
        return redirect(url_for("login"))
    # Validar que el archivo esté dentro de la ruta esperada (sin path traversal)
    root = _correos_root(maquina, mes)
    eml_path = (root / relative_eml).resolve()
    try:
        eml_path.relative_to(root.resolve())
    except ValueError:
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


STALE_SECS = 5 * 60


def _leer_progreso():
    """Lee todos los archivos *_progreso.json del directorio de estado."""
    result = {}
    d = estado_dir()
    try:
        if not d.is_dir():
            return result
        for fp in d.glob("*_progreso.json"):
            machine = fp.stem.replace("_progreso", "")
            try:
                data = json.loads(fp.read_text(encoding="utf-8"))
                upd = datetime.strptime(data["actualizado"], "%Y-%m-%d %H:%M:%S")
                age = (datetime.now() - upd).total_seconds()
                if age > STALE_SECS:
                    data["etapa"] = "interrumpido"
                    data["detalle"] = f"Sin señal hace {int(age//60)} min — puede haberse apagado"
                data["segundos"] = int(age)
                result[machine] = data
            except Exception:
                pass
    except Exception:
        pass
    return result


def _leer_reagendados():
    """Lee solicitudes de reintento pendientes (escritas por el agente al generar reporte)."""
    result = {}
    sol = Path(CFG.get("destino_root", "")) / "_solicitudes"
    try:
        if not sol.is_dir():
            return result
        for fp in sol.glob("*.json"):
            machine = fp.stem
            try:
                data = json.loads(fp.read_text(encoding="utf-8"))
                motivo = data.get("motivo", "")
                if motivo in ("reagendado_por_error", "reagendado_sin_respaldo"):
                    result[machine] = {
                        "motivo": motivo,
                        "errores": data.get("errores_previos", []),
                        "pedido_el": data.get("pedido_el", ""),
                    }
            except Exception:
                pass
    except Exception:
        pass
    return result


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
    if not puede_ver():
        return jsonify({"error": "no autorizado"}), 401
    emls = _list_emls(maquina, mes)
    return jsonify({"total": len(emls), "correos": emls})




def _audit_event(action, details=None):
    CFG_DIR.mkdir(parents=True, exist_ok=True)
    record = {"at": datetime.now().astimezone().isoformat(), "username": session.get("username", "system"),
              "role": "admin" if logueado() else ("readonly" if solo_lectura() else "system"),
              "action": action, "details": details or {}}
    with (CFG_DIR / "audit.jsonl").open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, ensure_ascii=False) + "\n")


def _sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def backup_history():
    source = Path(CFG.get("destino_root", "")) / "_historial"
    target_root = Path(CFG.get("history_backup_root", ""))
    if not CFG.get("history_backup_root"):
        raise RuntimeError("Configura un destino independiente para el historial.")
    if not source.is_dir():
        return {"copied": 0, "verified": 0, "status": "empty"}
    if source.resolve() == target_root.resolve() or source.resolve() in target_root.resolve().parents or target_root.resolve() in source.resolve().parents:
        raise RuntimeError("El destino del historial no puede estar dentro de la carpeta fuente.")
    if source.resolve().drive.casefold() == target_root.resolve().drive.casefold():
        raise RuntimeError("El historial debe respaldarse en un volumen físicamente independiente.")
    target = target_root / date.today().isoformat()
    copied = verified = 0
    for src in source.rglob("*"):
        if not src.is_file():
            continue
        dst = target / src.relative_to(source)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        copied += 1
        if src.stat().st_size != dst.stat().st_size or _sha256_file(src) != _sha256_file(dst):
            raise IOError(f"La copia no pasó la verificación SHA-256: {src}")
        verified += 1
    return {"copied": copied, "verified": verified, "status": "verified"}




def backup_history_with_retries():
    delays = (300, 900, 1800)
    log_path = CFG_DIR / "history_backup_log.jsonl"
    for attempt in range(4):
        try:
            result = backup_history()
            record = {"at": datetime.now().astimezone().isoformat(), "attempt": attempt + 1, **result}
            CFG_DIR.mkdir(parents=True, exist_ok=True)
            with log_path.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(record, ensure_ascii=False) + "\n")
            return result
        except Exception as exc:
            record = {"at": datetime.now().astimezone().isoformat(), "attempt": attempt + 1, "status": "error", "error": str(exc)}
            CFG_DIR.mkdir(parents=True, exist_ok=True)
            with log_path.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(record, ensure_ascii=False) + "\n")
            if attempt == 3:
                raise
            time.sleep(delays[attempt])

@app.route("/api/historial")
def api_historial():
    if not puede_ver():
        return jsonify({"error": "no autorizado"}), 401
    machine = request.args.get("maquina", "").strip()
    period = request.args.get("mes", "").strip()
    execution_id = request.args.get("execution_id", "").strip()
    user = request.args.get("usuario", "").strip()
    if any(x in machine for x in ("/", "\\", "..")):
        return jsonify({"error": "máquina inválida"}), 400
    if execution_id:
        try:
            execution_id = str(uuid.UUID(execution_id))
        except ValueError:
            return jsonify({"error": "execution_id inválido"}), 400
    root = Path(CFG.get("destino_root", "")) / "_historial"
    files = [root / machine / f"{execution_id}.json"] if machine and execution_id else (list(root.rglob("*.json")) if root.exists() else [])
    records = []
    for item in files:
        try:
            value = json.loads(item.read_text(encoding="utf-8"))
            if period and value.get("period") != period:
                continue
            if machine and value.get("machine", "").casefold() != machine.casefold():
                continue
            if user and value.get("user", "").casefold() != user.casefold():
                continue
            if execution_id and value.get("execution_id") != execution_id:
                continue
            records.append(value)
        except (OSError, ValueError):
            continue
    records.sort(key=lambda x: x.get("started_at", ""), reverse=True)
    return jsonify({"total": len(records), "historial": records})


@app.route("/api/traslados", methods=["GET", "POST"])
def api_traslados():
    folder = Path(CFG.get("destino_root", "")) / "_traslados"
    if request.method == "GET":
        if not puede_ver():
            return jsonify({"error": "no autorizado"}), 401
        events = []
        for f in folder.glob("*.json") if folder.exists() else []:
            try:
                events.append(json.loads(f.read_text(encoding="utf-8")))
            except (OSError, ValueError):
                continue
        return jsonify({"traslados": events})
    if not logueado():
        return jsonify({"error": "solo administrador"}), 403
    data = request.get_json(silent=True) or {}
    machine = str(data.get("maquina", "")).strip()
    period = str(data.get("mes", "")).strip()
    execution_id = str(data.get("execution_id", "")).strip()
    disk = str(data.get("disco", "")).strip()
    if not machine or any(x in machine for x in ("/", "\\", "..")) or not disk or len(period) != 7 or period[4] != "-" or not period[:4].isdigit() or not period[5:].isdigit() or not 1 <= int(period[5:]) <= 12:
        return jsonify({"error": "Máquina, período AAAA-MM y disco externo son obligatorios."}), 400
    try:
        execution_id = str(uuid.UUID(execution_id))
    except ValueError:
        return jsonify({"error": "execution_id inválido."}), 400
    history = Path(CFG.get("destino_root", "")) / "_historial" / machine / f"{execution_id}.json"
    try:
        execution = json.loads(history.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return jsonify({"error": "No existe ese execution_id en el historial."}), 404
    if execution.get("period") != period or execution.get("verification_status") != "verified":
        return jsonify({"error": "El respaldo debe tener verificación de transferencia."}), 409
    event = {"event_id": str(uuid.uuid4()), "execution_id": execution_id, "machine": machine, "period": period,
             "registered_at": datetime.now().astimezone().isoformat(), "administrator": session.get("username"),
             "disk": disk, "notes": str(data.get("notes", ""))[:1000], "physical_verification": False}
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"{event['event_id']}.json").write_text(json.dumps(event, ensure_ascii=False, indent=2), encoding="utf-8")
    _audit_event("external_transfer_registered", event)
    return jsonify({"ok": True, "event": event}), 201


@app.route("/api/respaldo/<maquina>/<periodo>", methods=["DELETE"])
def api_eliminar_respaldo(maquina, periodo):
    if not logueado():
        return jsonify({"error": "solo administrador"}), 403
    data = request.get_json(silent=True) or {}
    if data.get("confirmacion") != "ELIMINAR":
        return jsonify({"error": "Escribe ELIMINAR para confirmar."}), 400
    if len(periodo) != 7 or periodo[4] != "-" or not periodo[:4].isdigit() or not periodo[5:].isdigit() or not 1 <= int(periodo[5:]) <= 12:
        return jsonify({"error": "Período inválido."}), 400
    if not maquina or any(x in maquina for x in ("/", "\\", "..")):
        return jsonify({"error": "Máquina inválida."}), 400
    try:
        event_id = str(uuid.UUID(str(data.get("event_id", ""))))
    except ValueError:
        return jsonify({"error": "event_id inválido."}), 400
    event_file = Path(CFG.get("destino_root", "")) / "_traslados" / f"{event_id}.json"
    try:
        event = json.loads(event_file.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return jsonify({"error": "Registra primero el traslado externo."}), 409
    if event.get("machine") != maquina or event.get("period") != periodo:
        return jsonify({"error": "El traslado no corresponde a ese respaldo."}), 409
    root = Path(CFG["destino_root"]).resolve()
    target = (root / maquina / periodo).resolve()
    try:
        target.relative_to(root)
    except ValueError:
        return jsonify({"error": "Ruta fuera del destino de respaldos."}), 403
    if not target.is_dir():
        return jsonify({"error": "No se encontró el respaldo."}), 404
    shutil.rmtree(target)
    _audit_event("backup_manually_deleted", {"machine": maquina, "period": periodo, "event_id": event_id})
    return jsonify({"ok": True})



@app.route("/historial")
def historial_view():
    if not puede_ver():
        return redirect(url_for("login"))
    return """<!doctype html><meta charset="utf-8"><title>Historial de ejecuciones</title>
    <h1>Historial de respaldos</h1><form id="f"><input name="maquina" placeholder="Máquina">
    <input name="usuario" placeholder="Usuario Windows"><input name="mes" placeholder="AAAA-MM"><input name="execution_id" placeholder="execution_id">
    <button>Buscar</button></form><pre id="r"></pre>
    <script>document.querySelector('#f').onsubmit=async e=>{e.preventDefault();let q=new URLSearchParams(new FormData(e.target));
    let r=await fetch('/api/historial?'+q);document.querySelector('#r').textContent=JSON.stringify(await r.json(),null,2)}</script>"""


@app.route("/traslados")
def traslados_view():
    if not logueado():
        return redirect(url_for("login"))
    return """<!doctype html><meta charset="utf-8"><title>Traslados</title>
    <h1>Registrar traslado a disco externo</h1>
    <form id="f"><input name="maquina" placeholder="Máquina" required><input name="mes" placeholder="AAAA-MM" required>
    <input name="execution_id" placeholder="execution_id" required><input name="disco" placeholder="Disco externo" required>
    <input name="notes" placeholder="Notas"><button>Registrar</button></form><pre id="r"></pre><h2>Registros existentes</h2><ul id="events"></ul>
    <script>document.querySelector('#f').onsubmit=async e=>{e.preventDefault();let d=Object.fromEntries(new FormData(e.target));
    let r=await fetch('/api/traslados',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});
    document.querySelector('#r').textContent=await r.text();refresh()};async function refresh(){let d=await(await fetch('/api/traslados')).json();let ul=document.querySelector('#events');ul.textContent='';
    for(let x of d.traslados){let li=document.createElement('li');li.textContent=x.period+' · '+x.machine+' · '+x.disk+' · '+x.event_id+' ';
    let b=document.createElement('button');b.textContent='Eliminar respaldo local';b.onclick=async()=>{if(prompt('Escribe ELIMINAR para confirmar')!=='ELIMINAR')return;
    let z=await fetch('/api/respaldo/'+encodeURIComponent(x.machine)+'/'+encodeURIComponent(x.period),{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({confirmacion:'ELIMINAR',event_id:x.event_id})});alert(await z.text());refresh()};
    li.appendChild(b);ul.appendChild(li)}}refresh()</script>"""
if __name__ == "__main__":
    if "--backup-history" in sys.argv:
        try:
            print(json.dumps(backup_history_with_retries(), ensure_ascii=False))
        except Exception as exc:
            CFG_DIR.mkdir(parents=True, exist_ok=True)
            with (CFG_DIR / "history_backup_errors.log").open("a", encoding="utf-8") as fh:
                fh.write(f"{datetime.now().isoformat()} {exc}\n")
            raise
    else:
        cert = CFG.get("tls_cert", "").strip()
        key = CFG.get("tls_key", "").strip()
        if cert and key and Path(cert).is_file() and Path(key).is_file():
            print("Panel protegido con HTTPS en https://localhost:5000")
            app.run(host="0.0.0.0", port=5000, threaded=True, ssl_context=(cert, key))
        else:
            print("Configura tls_cert y tls_key desde el panel local; mientras tanto solo escucha en 127.0.0.1.")
            app.run(host="127.0.0.1", port=5000, threaded=True)