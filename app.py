import os, json, re, base64, sqlite3
from flask import Flask, request, jsonify, send_file, render_template, g
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import anthropic
import tempfile

app = Flask(__name__)

# En Railway, los archivos se guardan en /data si existe, si no en el dir actual
DATA_DIR = "/data" if os.path.exists("/data") else os.path.dirname(__file__)
DB_PATH = os.path.join(DATA_DIR, "permisos.db")

HEADERS = [
    "Nº PE", "Fecha Oficialización", "Buque", "Bandera", "País Destino",
    "Mercadería", "Posición Arancelaria", "Cond. Venta",
    "Peso Bruto (kg)", "Toneladas", "FOB Total (USD)", "Flete (USD)",
    "Precio Unit. (USD/Tn)", "Precio Oficial (USD/Tn)",
    "DJVE", "Fecha Cierre Vta.", "Vto. Embarque",
    "Ant. Gan. (USD)", "Cotización"
]

FIELDS = [
    "nro_pe", "fecha_oficializacion", "buque", "bandera", "pais_destino",
    "mercaderia", "posicion_arancelaria", "cond_venta",
    "peso_bruto_kg", "toneladas", "fob_total_usd", "flete_usd",
    "precio_unit_usd_tn", "precio_oficial_usd_tn",
    "djve", "fecha_cierre_vta", "vto_embarque",
    "ant_gan_usd", "cotizacion"
]

# ---------- DB ----------

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DB_PATH)
        db.row_factory = sqlite3.Row
    return db

@app.teardown_appcontext
def close_db(exc):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def init_db():
    with app.app_context():
        db = get_db()
        db.execute(f"""
            CREATE TABLE IF NOT EXISTS permisos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {', '.join(f'{f} TEXT' for f in FIELDS)},
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        db.commit()
        # Seed with existing data if DB is empty
        count = db.execute("SELECT COUNT(*) FROM permisos").fetchone()[0]
        if count == 0:
            seed_data = [
                ("26040EC01000700V","13/05/2026","EVER FEAT","LIBERIA","MALASIA","Aceite de Girasol (a granel)","1512.11.10.919G","FCA",264000,264,371712.00,None,1408.00,1288,"26001DJVE001338V","21/04/2026","27/06/2026",1700.16,1384.00),
                ("26040EC01000719X","15/05/2026","MERCOSUL ITAJAI","BRASIL","MALASIA","Aceite de Girasol (a granel)","1512.11.10.919G","CFR",198000,198,255024.00,4356.00,1288.00,1288,"26001DJVE001392V","24/04/2026","29/06/2026",1275.12,1391.00),
                ("26040EC01000720A","15/05/2026","MERCOSUL SANTOS","BRASIL","MALASIA","Aceite de Girasol (a granel)","1512.11.10.919G","CFR",198000,198,255024.00,4356.00,1288.00,1288,"26001DJVE001392V","24/04/2026","29/06/2026",1275.12,1391.00),
                ("26040EC01000744G","20/05/2026","MSC CHLOE","PORTUGAL","JAPÓN","Aceite de Girasol Alto Oleico (a granel)","1512.11.10.911P","CFR",315000,315,393368.85,11721.15,1248.79,1286,"26001DJVE001694D","18/05/2026","04/07/2026",2025.45,1398.00),
                ("26040EC01000745H","20/05/2026","MSC CHLOE","PORTUGAL","JAPÓN","Aceite de Girasol Alto Oleico (a granel)","1512.11.10.911P","CFR",315000,315,405090.00,11721.15,1286.00,1286,"26001DJVE001694D","18/05/2026","04/07/2026",2025.45,1398.00),
                ("26040EC01000764X","21/05/2026","CMA CGM RODOLPHE","SINGAPUR","MALASIA","Aceite de Girasol (a granel)","1512.11.10.919G","FCA",264000,264,371712.00,None,1408.00,1288,"26001DJVE001338V","21/04/2026","05/07/2026",1700.16,1397.00),
            ]
            placeholders = ','.join(['?'] * len(FIELDS))
            db.executemany(f"INSERT INTO permisos ({','.join(FIELDS)}) VALUES ({placeholders})", seed_data)
            db.commit()

# ---------- Excel builder ----------

def _border():
    t = Side(style="thin")
    return Border(left=t, right=t, top=t, bottom=t)

def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Permisos de Exportación"

    # Header
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", start_color="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _border()
    ws.row_dimensions[1].height = 35

    fills = [PatternFill("solid", start_color="DCE6F1"), PatternFill("solid", start_color="FFFFFF")]
    for r_idx, row in enumerate(rows, 2):
        fill = fills[r_idx % 2]
        for c_idx, field in enumerate(FIELDS, 1):
            val = row[field]
            try: val = float(val) if val and '.' in str(val) else (int(val) if val else None)
            except: pass
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = _border()
            cell.fill = fill

    # Totals
    total_row = len(rows) + 2
    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.fill = PatternFill("solid", start_color="BDD7EE")
        cell.border = _border()
        cell.font = Font(name="Arial", bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=total_row, column=1).value = "TOTALES"
    for col in [9, 10, 11, 12, 18]:
        cl = get_column_letter(col)
        ws.cell(row=total_row, column=col).value = f"=SUM({cl}2:{cl}{total_row-1})"

    widths = [20,18,20,12,14,35,20,14,16,12,16,14,18,20,22,18,16,16,12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    return tmp.name

# ---------- Claude extraction ----------

def extract_pe_with_claude(pdf_bytes):
    client = anthropic.Anthropic()
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
    prompt = """Extraé los siguientes campos de este Permiso de Exportación (PE) argentino y devolvé SOLO un JSON válido, sin markdown ni texto adicional:

{
  "nro_pe": "ej: 26040EC01000764X",
  "fecha_oficializacion": "DD/MM/AAAA",
  "buque": "nombre del buque",
  "bandera": "país de la bandera",
  "pais_destino": "país destino",
  "mercaderia": "descripción corta de la mercadería",
  "posicion_arancelaria": "ej: 1512.11.10.919G",
  "cond_venta": "FCA o CFR o FOB etc",
  "peso_bruto_kg": número entero,
  "toneladas": número entero,
  "fob_total_usd": número decimal,
  "flete_usd": número decimal o null si no hay,
  "precio_unit_usd_tn": número decimal,
  "precio_oficial_usd_tn": número entero,
  "djve": "ej: 26001DJVE001338V",
  "fecha_cierre_vta": "DD/MM/AAAA",
  "vto_embarque": "DD/MM/AAAA",
  "ant_gan_usd": número decimal,
  "cotizacion": número decimal
}

Devolvé SOLO el JSON, nada más."""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
            {"type": "text", "text": prompt}
        ]}]
    )
    text = response.content[0].text.strip()
    text = re.sub(r'^```json\s*', '', text)
    text = re.sub(r'\s*```$', '', text)
    return json.loads(text)

# ---------- Routes ----------

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "pdf" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    pdf_bytes = request.files["pdf"].read()
    try:
        data = extract_pe_with_claude(pdf_bytes)
    except Exception as e:
        return jsonify({"error": f"Error extrayendo datos: {str(e)}"}), 500
    db = get_db()
    exists = db.execute("SELECT 1 FROM permisos WHERE nro_pe=?", (data.get("nro_pe",""),)).fetchone()
    data["already_exists"] = exists is not None
    return jsonify(data)

@app.route("/confirm", methods=["POST"])
def confirm():
    pe = request.json
    nro = pe.get("nro_pe", "")
    db = get_db()
    if db.execute("SELECT 1 FROM permisos WHERE nro_pe=?", (nro,)).fetchone():
        return jsonify({"error": f"El PE {nro} ya existe"}), 409
    pe.pop("already_exists", None)
    placeholders = ','.join(['?'] * len(FIELDS))
    vals = [str(pe.get(f, '')) if pe.get(f) is not None else None for f in FIELDS]
    db.execute(f"INSERT INTO permisos ({','.join(FIELDS)}) VALUES ({placeholders})", vals)
    db.commit()
    return jsonify({"ok": True, "nro_pe": nro})

@app.route("/download")
def download():
    db = get_db()
    rows = db.execute(f"SELECT {','.join(FIELDS)} FROM permisos ORDER BY fecha_oficializacion, nro_pe").fetchall()
    path = build_excel(rows)
    return send_file(path, as_attachment=True, download_name="permisos_exportacion.xlsx")

@app.route("/list")
def list_pes():
    db = get_db()
    rows = db.execute("SELECT nro_pe, buque, pais_destino, fecha_oficializacion FROM permisos ORDER BY fecha_oficializacion DESC, nro_pe DESC").fetchall()
    return jsonify([dict(r) for r in rows])

if __name__ == "__main__":
    init_db()
    print("✅ App corriendo en http://localhost:5000")
    app.run(debug=False, port=5000)
