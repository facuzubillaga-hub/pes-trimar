import pandas as pd
import io
from datetime import datetime

# Mapeo flexible de nombres de columna a campo normalizado
COL_MAP = {
    'buque': ['buque', 'vessel'],
    'booking': ['booking'],
    'producto': ['producto', 'porduct', 'product'],
    'consignee': ['consignee'],
    'cantidad': ['cantidad', 'containers'],
    'toneladas': ['tn'],
    'pod': ['pod'],
    'etd': ['etd'],
    'fecha_carga': ['fecha de carga', 'consolidated date'],
    'contrato': ['contrato'],
    'envase': ['envase', 'packing'],
    'linea': ['maritime line'],
    'bandera': ['flag'],
}

def _find_col(columns, candidates):
    """Find first matching column name (case-insensitive, stripped)."""
    cols_clean = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        if cand.lower() in cols_clean:
            return cols_clean[cand.lower()]
    return None

def _safe_str(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    return s if s and s != 'nan' else None

def _safe_date(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        try:
            # Excel serial date
            d = pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(val))
            return d.strftime('%d/%m/%Y')
        except:
            return None
    if hasattr(val, 'strftime'):
        return val.strftime('%d/%m/%Y')
    s = str(val).strip()
    return s if s and s != 'nan' else None

def parse_plan_cargas(file_bytes):
    """Parse all sheets from Plan de Cargas Excel. Returns list of row dicts."""
    xl = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)
    rows = []

    for sheet_name, df in xl.items():
        # Find header row: first row with >5 non-empty cells
        header_row = None
        for i, row in df.iterrows():
            vals = [v for v in row if str(v) not in ['nan', '', 'NaT'] and v == v]
            if len(vals) > 5:
                header_row = i
                break
        if header_row is None:
            continue

        data = df.iloc[header_row+1:].copy()
        data.columns = df.iloc[header_row]
        data = data.dropna(how='all')

        # Map columns
        col = {field: _find_col(data.columns, candidates)
               for field, candidates in COL_MAP.items()}

        for _, row in data.iterrows():
            booking = _safe_str(row[col['booking']]) if col['booking'] else None
            buque = _safe_str(row[col['buque']]) if col['buque'] else None

            # Skip rows without booking or buque
            if not booking and not buque:
                continue
            # Skip header-like repeated rows
            if booking and 'booking' in booking.lower():
                continue

            # Clean buque: sometimes has voyage number after /
            if buque and '/' in buque:
                buque = buque.split('/')[0].strip()

            rows.append({
                'semana': sheet_name,
                'buque': buque,
                'booking': booking,
                'producto': _safe_str(row[col['producto']]) if col['producto'] else None,
                'consignee': _safe_str(row[col['consignee']]) if col['consignee'] else None,
                'cantidad_contenedores': _safe_str(row[col['cantidad']]) if col['cantidad'] else None,
                'toneladas': _safe_str(row[col['toneladas']]) if col['toneladas'] else None,
                'pod': _safe_str(row[col['pod']]) if col['pod'] else None,
                'etd': _safe_date(row[col['etd']]) if col['etd'] else None,
                'fecha_carga': _safe_date(row[col['fecha_carga']]) if col['fecha_carga'] else None,
                'contrato': _safe_str(row[col['contrato']]) if col['contrato'] else None,
                'envase': _safe_str(row[col['envase']]) if col['envase'] else None,
                'linea': _safe_str(row[col['linea']]) if col['linea'] else None,
            })

    return rows
