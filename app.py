from flask import (Flask, render_template, request, redirect,
                   url_for, flash, send_from_directory, abort, jsonify)
import sqlite3, os, uuid, re
from datetime import datetime
from werkzeug.utils import secure_filename
import requests as http_requests

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'trimar-pe-secret-2024')

UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', '/data/uploads')
DB_PATH       = os.environ.get('DB_PATH', 'pe.db')
ALERTA_HORAS  = int(os.environ.get('ALERTA_HORAS', 24))
ALLOWED_EXT   = {'pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx', 'xls', 'xlsx'}

RESEND_API_KEY = os.environ.get('RESEND_API_KEY', '')
MAIL_FROM      = os.environ.get('MAIL_FROM', 'Trimar PE <onboarding@resend.dev>')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ── ESTADOS (6 pasos, sin provisorio) ─────────────────────────────────────────
ESTADOS = [
    ('SOLICITUD_RECIBIDA',    'Solicitud recibida',    'Bunge envió la solicitud'),
    ('ENVIADO_DESPACHANTE',   'Enviado a despachante', 'Derivado al despachante'),
    ('OFICIALIZADO_RECIBIDO', 'Oficializado recibido', 'Despachante envió el PE oficial'),
    ('ENVIADO_BUNGE',         'Enviado a Bunge',       'PE oficial enviado a Bunge'),
    ('CUMPLIDO_PENDIENTE',    'Cumplido pendiente',    'Buque terminó de cargar'),
    ('CUMPLIDO_ENVIADO',      'Cumplido enviado',      'PE cumplido enviado a Bunge'),
]
ESTADO_KEYS = [e[0] for e in ESTADOS]

ACCIONES = {
    'SOLICITUD_RECIBIDA':    ('Enviar al despachante',   'ENVIADO_DESPACHANTE'),
    'ENVIADO_DESPACHANTE':   ('Registrar oficializado',  'OFICIALIZADO_RECIBIDO'),
    'OFICIALIZADO_RECIBIDO': ('Enviar a Bunge',          'ENVIADO_BUNGE'),
    'ENVIADO_BUNGE':         ('Registrar fin de carga',  'CUMPLIDO_PENDIENTE'),
    'CUMPLIDO_PENDIENTE':    ('Enviar cumplido a Bunge', 'CUMPLIDO_ENVIADO'),
    'CUMPLIDO_ENVIADO':      (None, None),
}

ESPERA_TRIMAR      = {'SOLICITUD_RECIBIDA','OFICIALIZADO_RECIBIDO','CUMPLIDO_PENDIENTE'}
ESPERA_DESPACHANTE = {'ENVIADO_DESPACHANTE'}
ESPERA_BUNGE       = {'ENVIADO_BUNGE'}

# ── HELPERS ────────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

def save_upload(file):
    if not file or file.filename == '':
        return None, None
    if not allowed_file(file.filename):
        return None, 'Tipo de archivo no permitido'
    ext = file.filename.rsplit('.', 1)[1].lower()
    unique_name = f"{uuid.uuid4().hex}.{ext}"
    file.save(os.path.join(UPLOAD_FOLDER, unique_name))
    return unique_name, None

def estado_label(key):
    if key == 'ELIMINADO': return 'Eliminado'
    for k, label, _ in ESTADOS:
        if k == key: return label
    return key

def estado_desc(key):
    for k, _, desc in ESTADOS:
        if k == key: return desc
    return ''

def estado_index(key):
    try: return ESTADO_KEYS.index(key)
    except ValueError: return 0

def horas_en_estado(fecha_str):
    try:
        fecha = datetime.fromisoformat(fecha_str)
        return (datetime.now() - fecha).total_seconds() / 3600
    except: return 0

def format_fecha(fecha_str):
    if not fecha_str: return '—'
    try:
        dt = datetime.fromisoformat(fecha_str)
        return dt.strftime('%d/%m/%Y %H:%M')
    except: return fecha_str

def format_fecha_corta(fecha_str):
    if not fecha_str: return '—'
    try:
        dt = datetime.fromisoformat(fecha_str)
        return dt.strftime('%d/%m/%y')
    except: return fecha_str

def format_ton(valor):
    """Formatea toneladas con 3 decimales."""
    if valor is None: return '—'
    try: return f'{float(valor):,.3f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except: return str(valor)

def validar_numero_pe(numero):
    """Valida formato: (2 dígitos año)(EC01)(3 dígitos aduana)(6 dígitos nro)(1 letra)
    Ej: 26040EC01000961H"""
    if not numero: return True  # vacío es ok, no es obligatorio
    patron = r'^\d{2}\d{3}EC\d{2}\d{6}[A-Z]$'
    # Flexible: acepta también con espacios que se normalizan
    numero_clean = numero.replace(' ', '')
    return bool(re.match(r'^\d{5}EC\d{2}\d{6}[A-Z]$', numero_clean))

def registrar_historial(conn, permiso_id, estado_anterior, estado_nuevo, usuario, notas=None, archivo=None):
    conn.execute('''
        INSERT INTO historial (permiso_id, estado_anterior, estado_nuevo, usuario, notas, archivo, fecha)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (permiso_id, estado_anterior, estado_nuevo, usuario, notas, archivo, datetime.now().isoformat()))

# ── EXTRACCIÓN PDF SOLICITUD BUNGE ─────────────────────────────────────────────
def extraer_datos_pdf(filepath):
    datos = {}
    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            texto = '\n'.join(page.extract_text() or '' for page in pdf.pages)

        m = re.search(r'Buque:.*?\n([A-Z][A-Z ]+?)(?:\s+Producto:|$)', texto, re.MULTILINE)
        if m: datos['buque'] = m.group(1).strip()

        m = re.search(r'Parcel:\s+(\S+)', texto)
        if m: datos['parcel'] = m.group(1).strip()

        m = re.search(r'Producto:\s+\d+.*?\n\d+\s+([A-Z][A-Z ]+?)(?:\s+CAE:)', texto, re.MULTILINE)
        if m: datos['producto'] = m.group(1).strip()

        m = re.search(r'([\d\.]+,\d+)TO\b', texto)
        if not m: m = re.search(r'([\d\.]+,\d+)\s+TO\b', texto)
        if m:
            ton_norm = m.group(1).replace('.','').replace(',','.')
            try: float(ton_norm); datos['toneladas'] = ton_norm
            except: pass

        m = re.search(r'Fecha Solicitud:\s+(\d{2}\.\d{2}\.\d{4})', texto)
        if m: datos['fecha_solicitud'] = m.group(1)

    except Exception as e:
        datos['error'] = str(e)
    return datos

# ── EXTRACCIÓN PE OFICIALIZADO ─────────────────────────────────────────────────
def extraer_datos_pe_oficializado(filepath):
    datos = {}
    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            texto = '\n'.join(page.extract_text() or '' for page in pdf.pages)

        m = re.search(r'(\d{2})\s+(\d{3})\s+([A-Z]{2}\d{2})\s+(\d{6})\s+([A-Z])\s+\d+ de \d+', texto)
        if m: datos['numero_pe'] = f"{m.group(1)}{m.group(2)}{m.group(3)}{m.group(4)}{m.group(5)}"

        m = re.search(r'Nombre del Transporte\s*\n\S+\s+\S+\s+\S+\s+([A-Z][A-Z ]+)', texto)
        if m: datos['buque'] = m.group(1).strip()

        m = re.search(r'TONELADA\s+([\d\.]+,\d+)', texto)
        if m:
            ton = m.group(1).replace('.','').replace(',','.')
            try: float(ton); datos['toneladas'] = ton
            except: pass

        m = re.search(r'(\d{2}/\d{2}/\d{4})\s+\*{4}', texto)
        if m: datos['vto_embarque'] = m.group(1)

        m = re.search(r'OFICIALIZADO\s+(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})', texto)
        if m: datos['fecha_oficializacion'] = m.group(1)

        m = re.search(r'Pais dest\.:\s*([A-Z]+)', texto)
        if m: datos['destino'] = m.group(1)

    except Exception as e:
        datos['error'] = str(e)
    return datos

# ── EMAIL (Resend) ─────────────────────────────────────────────────────────────
def enviar_email(destinatario, asunto, cuerpo_html, archivo_path=None, archivo_nombre=None):
    if not RESEND_API_KEY:
        return False, 'RESEND_API_KEY no configurado'
    try:
        payload = {'from': MAIL_FROM, 'to': [destinatario], 'subject': asunto, 'html': cuerpo_html}
        if archivo_path and os.path.exists(archivo_path):
            import base64
            with open(archivo_path, 'rb') as f:
                contenido = base64.b64encode(f.read()).decode('utf-8')
            payload['attachments'] = [{'filename': archivo_nombre or os.path.basename(archivo_path), 'content': contenido}]
        resp = http_requests.post('https://api.resend.com/emails',
            headers={'Authorization': f'Bearer {RESEND_API_KEY}', 'Content-Type': 'application/json'},
            json=payload, timeout=8)
        if resp.status_code in (200, 201): return True, None
        return False, f'Resend error {resp.status_code}: {resp.text}'
    except Exception as e:
        return False, str(e)

def mail_despachante(pe, despachante, archivo_path=None):
    if not despachante or not despachante['email']:
        return False, 'El despachante no tiene email registrado'
    asunto = f"Solicitud PE — {pe['buque_nombre']} | PARCEL {pe['viaje'] or ''} | {pe['producto'] or ''}"
    ton_str = format_ton(pe['toneladas_solicitadas']) + ' tn' if pe['toneladas_solicitadas'] else '—'
    cuerpo = f"""
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:#0f1923;padding:1.5rem 2rem;border-radius:8px 8px 0 0">
        <h2 style="color:#fff;margin:0;font-size:1.1rem">AGENCIA TRIMAR S.A.</h2>
        <p style="color:#5aaee0;margin:.25rem 0 0;font-size:.85rem">Puerto Quequén · Permisos de Embarque</p>
      </div>
      <div style="background:#f2f5f8;padding:2rem;border-radius:0 0 8px 8px">
        <p>Estimado/a <strong>{despachante['nombre']}</strong>,</p>
        <p>Les solicitamos la confección del siguiente Permiso de Embarque:</p>
        <table style="width:100%;border-collapse:collapse;margin:1.5rem 0;background:#fff;border-radius:6px;overflow:hidden">
          <tr style="background:#005b9a;color:#fff">
            <td style="padding:.6rem 1rem;font-size:.8rem;font-weight:600">CAMPO</td>
            <td style="padding:.6rem 1rem;font-size:.8rem;font-weight:600">DATO</td>
          </tr>
          <tr><td style="padding:.6rem 1rem;border-bottom:1px solid #dce4eb;font-size:.85rem;color:#7a8fa0">Buque</td>
              <td style="padding:.6rem 1rem;border-bottom:1px solid #dce4eb;font-size:.85rem;font-weight:600">{pe['buque_nombre']}</td></tr>
          <tr><td style="padding:.6rem 1rem;border-bottom:1px solid #dce4eb;font-size:.85rem;color:#7a8fa0">N° Proyecto (PARCEL)</td>
              <td style="padding:.6rem 1rem;border-bottom:1px solid #dce4eb;font-size:.85rem;font-weight:600">{pe['viaje'] or '—'}</td></tr>
          <tr><td style="padding:.6rem 1rem;border-bottom:1px solid #dce4eb;font-size:.85rem;color:#7a8fa0">Producto</td>
              <td style="padding:.6rem 1rem;border-bottom:1px solid #dce4eb;font-size:.85rem;font-weight:600">{pe['producto'] or '—'}</td></tr>
          <tr><td style="padding:.6rem 1rem;font-size:.85rem;color:#7a8fa0">Toneladas</td>
              <td style="padding:.6rem 1rem;font-size:.85rem;font-weight:600">{ton_str}</td></tr>
        </table>
        <p style="font-size:.85rem;color:#3a4a58">Se adjunta la solicitud original de Bunge Argentina.</p>
        <p style="font-size:.85rem;color:#3a4a58">Aguardamos el envío del PE oficializado a la brevedad.</p>
        <p style="margin-top:1.5rem;font-size:.85rem;color:#3a4a58">Saludos,<br><strong>Agencia Trimar S.A.</strong><br>trimar@trimar.com.ar</p>
      </div>
    </div>"""
    return enviar_email(despachante['email'], asunto, cuerpo,
                        archivo_path=archivo_path,
                        archivo_nombre=f"Solicitud_PE_{pe['buque_nombre'].replace(' ','_')}.pdf")

# ── JINJA GLOBALS ──────────────────────────────────────────────────────────────
app.jinja_env.globals.update(
    estado_label=estado_label, estado_desc=estado_desc, estado_index=estado_index,
    horas_en_estado=horas_en_estado, format_fecha=format_fecha,
    format_fecha_corta=format_fecha_corta, format_ton=format_ton,
    ESTADOS=ESTADOS, ESTADO_KEYS=ESTADO_KEYS, ACCIONES=ACCIONES,
    ESPERA_TRIMAR=ESPERA_TRIMAR, ESPERA_DESPACHANTE=ESPERA_DESPACHANTE,
    ESPERA_BUNGE=ESPERA_BUNGE, ALERTA_HORAS=ALERTA_HORAS,
    enumerate=enumerate, len=len, abs=abs,
)

# ── INIT DB ────────────────────────────────────────────────────────────────────
def init_db():
    with get_db() as conn:
        conn.executescript('''
            CREATE TABLE IF NOT EXISTS buques (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL, viaje TEXT,
                creado_en TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS despachantes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL, email TEXT, activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS permisos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                buque_id INTEGER NOT NULL, despachante_id INTEGER,
                numero_pe TEXT, producto TEXT, numero_oc TEXT,
                toneladas_solicitadas REAL, toneladas_cargadas REAL,
                estado TEXT NOT NULL DEFAULT 'SOLICITUD_RECIBIDA',
                fecha_solicitud TEXT, fecha_ultimo_cambio TEXT,
                fecha_cumplido TEXT, fecha_fin_carga TEXT,
                vto_embarque TEXT, fecha_oficializacion TEXT,
                notas TEXT,
                archivo_solicitud TEXT, archivo_oficializado TEXT, archivo_cumplido TEXT,
                facturado INTEGER DEFAULT 0, numero_factura TEXT, fecha_factura TEXT,
                eliminado INTEGER DEFAULT 0,
                creado_en TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (buque_id) REFERENCES buques(id),
                FOREIGN KEY (despachante_id) REFERENCES despachantes(id)
            );
            CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                permiso_id INTEGER NOT NULL, estado_anterior TEXT,
                estado_nuevo TEXT NOT NULL, usuario TEXT, notas TEXT,
                archivo TEXT, fecha TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (permiso_id) REFERENCES permisos(id)
            );
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL, activo INTEGER DEFAULT 1
            );
        ''')
        # Migraciones para BDs existentes
        cols = [r[1] for r in conn.execute("PRAGMA table_info(permisos)").fetchall()]
        migraciones = [
            ('facturado','INTEGER DEFAULT 0'), ('numero_factura','TEXT'),
            ('fecha_factura','TEXT'), ('fecha_fin_carga','TEXT'),
            ('vto_embarque','TEXT'), ('fecha_oficializacion','TEXT'),
            ('eliminado','INTEGER DEFAULT 0'), ('numero_oc','TEXT'),
        ]
        for col, tipo in migraciones:
            if col not in cols:
                conn.execute(f"ALTER TABLE permisos ADD COLUMN {col} {tipo}")

        cur = conn.execute('SELECT COUNT(*) FROM usuarios')
        if cur.fetchone()[0] == 0:
            for nombre in ['Facundo','Operador 2','Operador 3']:
                conn.execute("INSERT INTO usuarios (nombre) VALUES (?)", (nombre,))

# ── API ENDPOINTS ──────────────────────────────────────────────────────────────
@app.route('/api/extraer-pdf', methods=['POST'])
def api_extraer_pdf():
    if 'archivo' not in request.files:
        return jsonify({'error': 'Sin archivo'}), 400
    f = request.files['archivo']
    tmp_path = os.path.join(UPLOAD_FOLDER, f'tmp_{uuid.uuid4().hex}.pdf')
    f.save(tmp_path)
    datos = extraer_datos_pdf(tmp_path)
    os.remove(tmp_path)
    return jsonify(datos)

@app.route('/api/extraer-pe-oficializado', methods=['POST'])
def api_extraer_pe_oficializado():
    if 'archivo' not in request.files:
        return jsonify({'error': 'Sin archivo'}), 400
    f = request.files['archivo']
    tmp_path = os.path.join(UPLOAD_FOLDER, f'tmp_{uuid.uuid4().hex}.pdf')
    f.save(tmp_path)
    datos = extraer_datos_pe_oficializado(tmp_path)
    os.remove(tmp_path)
    return jsonify(datos)

@app.route('/api/usuarios')
def api_usuarios():
    db = get_db()
    lista = [dict(u) for u in db.execute('SELECT id, nombre FROM usuarios WHERE activo=1 ORDER BY nombre').fetchall()]
    return jsonify(lista)

# ── RUTAS PRINCIPALES ──────────────────────────────────────────────────────────
@app.route('/')
def index():
    db = get_db()
    filtro   = request.args.get('estado', '')
    busqueda = request.args.get('q', '').strip()
    desp_f   = request.args.get('despachante', '')
    fecha_d  = request.args.get('fecha_desde', '')
    fecha_h  = request.args.get('fecha_hasta', '')

    query = '''
        SELECT p.*, b.nombre as buque_nombre, b.viaje,
               d.nombre as despachante_nombre
        FROM permisos p
        JOIN buques b ON p.buque_id = b.id
        LEFT JOIN despachantes d ON p.despachante_id = d.id
        WHERE 1=1
    '''
    params = []
    if filtro == 'ELIMINADO':
        query += ' AND p.eliminado = 1'
    elif filtro:
        query += ' AND p.estado = ? AND p.eliminado = 0'; params.append(filtro)
    else:
        query += ' AND p.eliminado = 0'

    if busqueda:
        query += ' AND (b.nombre LIKE ? OR p.numero_pe LIKE ? OR p.producto LIKE ? OR b.viaje LIKE ? OR p.numero_oc LIKE ?)'
        params += [f'%{busqueda}%']*5
    if desp_f:
        query += ' AND p.despachante_id = ?'; params.append(desp_f)
    if fecha_d:
        query += ' AND DATE(p.creado_en) >= ?'; params.append(fecha_d)
    if fecha_h:
        query += ' AND DATE(p.creado_en) <= ?'; params.append(fecha_h)
    query += ' ORDER BY p.id DESC'
    permisos = db.execute(query, params).fetchall()

    contadores = {}
    for key, _, _ in ESTADOS:
        contadores[key] = db.execute('SELECT COUNT(*) FROM permisos WHERE estado=? AND eliminado=0',(key,)).fetchone()[0]
    contadores['ELIMINADO'] = db.execute('SELECT COUNT(*) FROM permisos WHERE eliminado=1').fetchone()[0]

    total   = db.execute('SELECT COUNT(*) FROM permisos WHERE eliminado=0').fetchone()[0]
    activos = db.execute("SELECT COUNT(*) FROM permisos WHERE estado != 'CUMPLIDO_ENVIADO' AND eliminado=0").fetchone()[0]
    alertas = 0
    for p in db.execute("SELECT fecha_ultimo_cambio,creado_en,estado FROM permisos WHERE estado != 'CUMPLIDO_ENVIADO' AND eliminado=0").fetchall():
        ref = p['fecha_ultimo_cambio'] or p['creado_en'] or ''
        if ref and horas_en_estado(ref) > ALERTA_HORAS:
            alertas += 1

    pend_facturar   = db.execute("SELECT COUNT(*) FROM permisos WHERE facturado=0 AND estado IN ('CUMPLIDO_PENDIENTE','CUMPLIDO_ENVIADO') AND eliminado=0").fetchone()[0]
    por_trimar      = sum(contadores.get(e,0) for e in ESPERA_TRIMAR)
    por_despachante = sum(contadores.get(e,0) for e in ESPERA_DESPACHANTE)
    por_bunge       = sum(contadores.get(e,0) for e in ESPERA_BUNGE)

    despachantes = db.execute('SELECT * FROM despachantes WHERE activo=1 ORDER BY nombre').fetchall()
    usuarios     = db.execute('SELECT * FROM usuarios WHERE activo=1').fetchall()

    return render_template('index.html',
        permisos=permisos, contadores=contadores,
        filtro=filtro, busqueda=busqueda, desp_f=desp_f,
        fecha_d=fecha_d, fecha_h=fecha_h,
        total=total, activos=activos, alertas=alertas,
        pend_facturar=pend_facturar,
        por_trimar=por_trimar, por_despachante=por_despachante, por_bunge=por_bunge,
        despachantes=despachantes, usuarios=usuarios,
    )

@app.route('/nuevo', methods=['GET','POST'])
def nuevo_pe():
    db = get_db()
    if request.method == 'POST':
        buque_nombre = request.form['buque_nombre'].strip()
        parcel       = request.form.get('viaje','').strip()
        producto     = request.form.get('producto','').strip()
        toneladas    = request.form.get('toneladas_solicitadas','') or None
        numero_oc    = request.form.get('numero_oc','').strip()
        notas        = request.form.get('notas','').strip()
        usuario      = request.form.get('usuario','Sistema')
        ahora        = datetime.now().isoformat()

        # Validar duplicado PARCEL
        if parcel:
            dup = db.execute('''
                SELECT p.id FROM permisos p JOIN buques b ON p.buque_id=b.id
                WHERE b.viaje=? AND p.eliminado=0
            ''', (parcel,)).fetchone()
            if dup:
                flash(f'Ya existe un PE con el N° de Proyecto "{parcel}" (PE #{dup["id"]}). Verificá antes de continuar.', 'danger')
                return redirect(request.url)

        archivo_sol = None
        if 'archivo_solicitud' in request.files:
            f = request.files['archivo_solicitud']
            nombre, err = save_upload(f)
            if err:
                flash(err, 'danger')
                return redirect(request.url)
            archivo_sol = nombre

        row = db.execute('SELECT id FROM buques WHERE nombre=? AND viaje=?',(buque_nombre,parcel)).fetchone()
        buque_id = row['id'] if row else db.execute(
            'INSERT INTO buques (nombre,viaje) VALUES (?,?)',(buque_nombre,parcel)
        ).lastrowid

        cur = db.execute('''
            INSERT INTO permisos
              (buque_id,producto,toneladas_solicitadas,numero_oc,notas,estado,
               fecha_solicitud,fecha_ultimo_cambio,archivo_solicitud)
            VALUES (?,?,?,?,?,'SOLICITUD_RECIBIDA',?,?,?)
        ''',(buque_id,producto,toneladas,numero_oc,notas,ahora,ahora,archivo_sol))
        permiso_id = cur.lastrowid

        registrar_historial(db, permiso_id, None, 'SOLICITUD_RECIBIDA', usuario,
                            'PE creado' + (' — solicitud adjunta' if archivo_sol else ''),
                            archivo_sol)
        db.commit()
        flash('Permiso de Embarque creado.', 'success')
        return redirect(url_for('detalle_pe', pe_id=permiso_id))

    usuarios = db.execute('SELECT * FROM usuarios WHERE activo=1').fetchall()
    return render_template('nuevo.html', usuarios=usuarios)

@app.route('/pe/<int:pe_id>')
def detalle_pe(pe_id):
    db  = get_db()
    pe  = db.execute('''
        SELECT p.*, b.nombre as buque_nombre, b.viaje,
               d.nombre as despachante_nombre, d.email as despachante_email
        FROM permisos p
        JOIN buques b ON p.buque_id=b.id
        LEFT JOIN despachantes d ON p.despachante_id=d.id
        WHERE p.id=?
    ''', (pe_id,)).fetchone()
    if not pe:
        flash('PE no encontrado.','danger')
        return redirect(url_for('index'))
    historial    = db.execute('SELECT * FROM historial WHERE permiso_id=? ORDER BY fecha DESC',(pe_id,)).fetchall()
    despachantes = db.execute('SELECT * FROM despachantes WHERE activo=1').fetchall()
    usuarios     = db.execute('SELECT * FROM usuarios WHERE activo=1').fetchall()
    return render_template('detalle.html', pe=pe, historial=historial,
                           despachantes=despachantes, usuarios=usuarios)

@app.route('/pe/<int:pe_id>/avanzar', methods=['POST'])
def avanzar_estado(pe_id):
    db = get_db()
    pe = db.execute('''
        SELECT p.*, b.nombre as buque_nombre, b.viaje,
               d.nombre as despachante_nombre, d.email as despachante_email
        FROM permisos p JOIN buques b ON p.buque_id=b.id
        LEFT JOIN despachantes d ON p.despachante_id=d.id
        WHERE p.id=?
    ''',(pe_id,)).fetchone()
    if not pe: abort(404)

    if pe['eliminado']:
        flash('No se puede operar sobre un PE eliminado.','danger')
        return redirect(url_for('detalle_pe', pe_id=pe_id))

    estado_actual = pe['estado']
    accion = ACCIONES.get(estado_actual)
    if not accion or not accion[1]:
        flash('Este PE ya está en el estado final.','warning')
        return redirect(url_for('detalle_pe', pe_id=pe_id))

    nuevo_estado     = accion[1]
    usuario          = request.form.get('usuario','Sistema')
    notas            = request.form.get('notas','').strip()
    ahora            = datetime.now().isoformat()
    updates          = {'estado': nuevo_estado, 'fecha_ultimo_cambio': ahora}
    archivo_guardado = None

    if nuevo_estado == 'ENVIADO_DESPACHANTE':
        desp_id = request.form.get('despachante_id')
        if desp_id:
            updates['despachante_id'] = desp_id
            desp = db.execute('SELECT * FROM despachantes WHERE id=?',(desp_id,)).fetchone()
            if desp and desp['email']:
                archivo_path = os.path.join(UPLOAD_FOLDER, pe['archivo_solicitud']) if pe['archivo_solicitud'] else None
                pe_dict = dict(pe); pe_dict['despachante_nombre'] = desp['nombre']
                ok, err = mail_despachante(pe_dict, desp, archivo_path)
                if ok: flash('✉ Mail enviado al despachante.','info')
                else:  flash(f'Advertencia: no se pudo enviar el mail ({err}).','warning')

    if nuevo_estado == 'OFICIALIZADO_RECIBIDO':
        numero_pe            = request.form.get('numero_pe','').strip()
        vto_embarque         = request.form.get('vto_embarque','').strip()
        fecha_oficializacion = request.form.get('fecha_oficializacion','').strip()
        if numero_pe:
            if not validar_numero_pe(numero_pe):
                flash('Formato de N° PE inválido. Debe ser: (2 dígitos año)(3 dígitos aduana)(EC01 o EC02)(6 dígitos)(letra). Ej: 26040EC01000961H','warning')
            updates['numero_pe'] = numero_pe
        if vto_embarque:         updates['vto_embarque'] = vto_embarque
        if fecha_oficializacion: updates['fecha_oficializacion'] = fecha_oficializacion
        if 'archivo' in request.files:
            nombre, err = save_upload(request.files['archivo'])
            if err: flash(err,'danger'); return redirect(url_for('detalle_pe', pe_id=pe_id))
            if nombre: updates['archivo_oficializado'] = nombre; archivo_guardado = nombre

    if nuevo_estado == 'CUMPLIDO_PENDIENTE':
        fecha_fin = request.form.get('fecha_fin_carga','').strip()
        ton_carg  = request.form.get('toneladas_cargadas','').strip()
        if not fecha_fin or not ton_carg:
            flash('La fecha/hora de finalización y las toneladas cargadas son obligatorias.','danger')
            return redirect(url_for('detalle_pe', pe_id=pe_id))
        updates['fecha_fin_carga']    = fecha_fin
        updates['toneladas_cargadas'] = ton_carg

    if nuevo_estado == 'CUMPLIDO_ENVIADO':
        updates['fecha_cumplido'] = ahora
        if 'archivo' in request.files:
            nombre, err = save_upload(request.files['archivo'])
            if err: flash(err,'danger'); return redirect(url_for('detalle_pe', pe_id=pe_id))
            if nombre: updates['archivo_cumplido'] = nombre; archivo_guardado = nombre

    set_clause = ', '.join(f'{k}=?' for k in updates)
    db.execute(f'UPDATE permisos SET {set_clause} WHERE id=?', list(updates.values())+[pe_id])
    registrar_historial(db, pe_id, estado_actual, nuevo_estado, usuario, notas or None, archivo_guardado)
    db.commit()
    flash(f'Estado actualizado: {estado_label(nuevo_estado)}','success')
    return redirect(url_for('detalle_pe', pe_id=pe_id))

@app.route('/pe/<int:pe_id>/editar', methods=['POST'])
def editar_pe(pe_id):
    db = get_db()
    db.execute('''UPDATE permisos SET numero_pe=?,producto=?,toneladas_solicitadas=?,numero_oc=?,notas=? WHERE id=?''',
        (request.form.get('numero_pe','').strip(),
         request.form.get('producto','').strip(),
         request.form.get('toneladas_solicitadas','') or None,
         request.form.get('numero_oc','').strip(),
         request.form.get('notas','').strip(), pe_id))
    db.commit()
    flash('Datos actualizados.','success')
    return redirect(url_for('detalle_pe', pe_id=pe_id))

@app.route('/pe/<int:pe_id>/facturar', methods=['POST'])
def facturar_pe(pe_id):
    db = get_db()
    pe = db.execute('SELECT * FROM permisos WHERE id=?',(pe_id,)).fetchone()
    if not pe: abort(404)
    numero_factura = request.form.get('numero_factura','').strip()
    fecha_factura  = request.form.get('fecha_factura','').strip()
    usuario        = request.form.get('usuario','Sistema')
    if not numero_factura or not fecha_factura:
        flash('El número y la fecha de factura son obligatorios.','danger')
        return redirect(url_for('detalle_pe', pe_id=pe_id))
    db.execute('UPDATE permisos SET facturado=1, numero_factura=?, fecha_factura=? WHERE id=?',
               (numero_factura, fecha_factura, pe_id))
    registrar_historial(db, pe_id, pe['estado'], pe['estado'], usuario,
                        f'Facturado — N° {numero_factura} del {fecha_factura}')
    db.commit()
    flash(f'Factura N° {numero_factura} registrada.','success')
    return redirect(url_for('detalle_pe', pe_id=pe_id))

@app.route('/pe/<int:pe_id>/eliminar', methods=['POST'])
def eliminar_pe(pe_id):
    db = get_db()
    confirmacion = request.form.get('confirmacion','').strip().lower()
    if confirmacion != 'borrar':
        flash('Confirmación incorrecta. Escribí "borrar" para eliminar.','danger')
        return redirect(url_for('detalle_pe', pe_id=pe_id))
    pe = db.execute('SELECT * FROM permisos WHERE id=?',(pe_id,)).fetchone()
    if not pe: abort(404)

    # Eliminar archivos
    for campo in ('archivo_solicitud','archivo_oficializado','archivo_cumplido'):
        if pe[campo]:
            try: os.remove(os.path.join(UPLOAD_FOLDER, pe[campo]))
            except: pass

    # Marcar como eliminado en lugar de borrar
    db.execute("UPDATE permisos SET eliminado=1, estado='ELIMINADO' WHERE id=?", (pe_id,))
    registrar_historial(db, pe_id, pe['estado'], 'ELIMINADO', 
                        request.form.get('usuario','Sistema'), 'PE eliminado')
    db.commit()
    flash(f'PE #{pe_id} marcado como eliminado.','success')
    return redirect(url_for('index'))

@app.route('/uploads/<path:filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

# ── DESPACHANTES ───────────────────────────────────────────────────────────────
@app.route('/despachantes')
def despachantes():
    db = get_db()
    lista = db.execute('SELECT * FROM despachantes ORDER BY nombre').fetchall()
    return render_template('despachantes.html', despachantes=lista)

@app.route('/despachantes/nuevo', methods=['POST'])
def nuevo_despachante():
    db = get_db()
    nombre = request.form['nombre'].strip()
    email  = request.form.get('email','').strip()
    db.execute('INSERT INTO despachantes (nombre,email) VALUES (?,?)',(nombre,email))
    db.commit()
    flash(f'Despachante {nombre} agregado.','success')
    return redirect(url_for('despachantes'))

@app.route('/despachantes/<int:d_id>/toggle', methods=['POST'])
def toggle_despachante(d_id):
    db = get_db()
    db.execute('UPDATE despachantes SET activo=1-activo WHERE id=?',(d_id,))
    db.commit()
    return redirect(url_for('despachantes'))

# ── USUARIOS ───────────────────────────────────────────────────────────────────
@app.route('/usuarios')
def usuarios():
    db = get_db()
    lista = db.execute('SELECT * FROM usuarios ORDER BY nombre').fetchall()
    return render_template('usuarios.html', usuarios=lista)

@app.route('/usuarios/nuevo', methods=['POST'])
def nuevo_usuario():
    db = get_db()
    nombre = request.form['nombre'].strip()
    db.execute('INSERT INTO usuarios (nombre) VALUES (?)',(nombre,))
    db.commit()
    flash(f'Usuario {nombre} agregado.','success')
    return redirect(url_for('usuarios'))

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
