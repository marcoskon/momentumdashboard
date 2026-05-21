"""
parse_data.py  —  FCI Momentum Pymes
Parser que lee GESTION MOMENTUM + BALANCE para generar:
  - Datos patrimoniales, cartera, VCP
  - Atribución diaria de VCP desde cuentas contables del Balance
"""

import json, re, subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd

ROOT     = Path(__file__).parent.parent
RAW_DIR  = ROOT / "raw"
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

DIAS_ES = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']

# Cuentas contables para atribución del VCP
# (código, tipo) — el nombre se lee automáticamente del archivo Balance
CUENTAS_ATRIB = [
    ('4000000000000000001', 'Precio' ),  # Resultado Títulos Públicos en $
    ('4000000000000000009', 'Precio' ),  # Diferencia de Cotización US3
    ('4000000000000000010', 'Devengo'),  # Resultado Cheque Pago diferido US3
    ('4000000000000000013', 'MTM'    ),  # Resultado de Operación Futuro
    ('4001001000000000004', 'Costo'  ),  # Costo Fondos Comunes de Inversión
    ('4001001000000000013', 'Costo'  ),  # Gastos por negociación de CPD
    ('4001001000000000015', 'Costo'  ),  # Gastos Sociedad Depositaria Clase A
    ('4001001000000000016', 'Costo'  ),  # Gastos Sociedad Depositaria Clase B
    ('4001001000000000017', 'Costo'  ),  # Gastos Sociedad Depositaria Clase C
    ('4001001000000000021', 'Costo'  ),  # Gastos por Negociación CPD DM
    ('4001002000000000021', 'Costo'  ),  # Resultado por CPD incobrables
    ('4001002001000000000', 'Costo'  ),  # Honorarios de Administración
    ('4002001000000000006', 'Devengo'),  # Ventas Fondos Comunes de Inversión
    ('4002001000000000017', 'Devengo'),  # Resultado Cheques de Pago Diferido
    ('4002001000000000022', 'Devengo'),  # Intereses Pase
    ('4002001002000000002', 'Devengo'),  # Intereses Caución
    ('4002001004000000013', 'Precio' ),  # Resultado por Tenencia Títulos Provinciales
]
CODIGO_TOTAL = '4000000000'

# ─── UTILIDADES ──────────────────────────────────────────────────────────────

def date_from_filename(fname):
    m = re.search(r'(\d{8})', fname)
    if m:
        r = m.group(1)
        return f"{r[:4]}-{r[4:6]}-{r[6:]}"
    m = re.search(r'(\d{2}-\d{2}-\d{4})', fname)
    if m:
        d, mo, y = m.group(1).split('-')
        return f"{y}-{mo}-{d}"
    return None

def safe(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return float(str(v).replace(',', '.').replace('%', '').strip())
    except Exception:
        return None

def col_a(df):
    return df.iloc[:, 0].astype(str).str.strip()

def find_row(df, *labels):
    ca = col_a(df).str.upper()
    for label in labels:
        mask = ca == label.upper()
        if mask.any():
            return mask.idxmax()
    return None

def find_row_contains(df, text):
    ca = col_a(df).str.upper()
    mask = ca.str.contains(text.upper(), na=False)
    return mask.idxmax() if mask.any() else None

def find_value(df, *labels):
    idx = find_row(df, *labels)
    return safe(df.iloc[idx, 1]) if idx is not None else None

def section_subtotal(df, section_label, value_col_idx):
    start = find_row_contains(df, section_label)
    if start is None:
        return None
    for i in range(start + 2, min(start + 500, len(df))):
        a = str(df.iloc[i, 0]).strip()
        v = safe(df.iloc[i, value_col_idx])
        if a in ('', 'nan', 'None') and v is not None:
            return v
    return None

def section_details(df, section_label, name_col, value_col, extra_cols=None):
    start = find_row_contains(df, section_label)
    if start is None:
        return []
    rows = []
    for i in range(start + 2, min(start + 500, len(df))):
        name = str(df.iloc[i, name_col]).strip()
        val  = safe(df.iloc[i, value_col])
        if name in ('', 'nan', 'None'):
            if val is not None:
                break
            continue
        if val is None:
            continue
        row = {'nombre': name, 'total': val, 'total_prev': None}
        if extra_cols:
            for k, idx in extra_cols.items():
                row[k] = safe(df.iloc[i, idx])
        rows.append(row)
    return rows

# ─── BALANCE: LEER CUENTAS CONTABLES ─────────────────────────────────────────

def to_xlsx(path: Path) -> Path:
    """Convierte .xls a .xlsx usando LibreOffice si es necesario."""
    if path.suffix.lower() == '.xlsx':
        return path
    out = Path('/tmp') / (path.stem + '.xlsx')
    if not out.exists():
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'xlsx',
             str(path), '--outdir', '/tmp/'],
            capture_output=True, timeout=60
        )
    return out

def read_balance_codes(path: Path) -> dict:
    """Lee el balance y devuelve {codigo: {'valor': float, 'nombre': str}}."""
    try:
        xlsx = to_xlsx(path)
        df = pd.read_excel(xlsx, header=None)
        data = {}
        for _, row in df.iterrows():
            code = str(row[1]).strip().replace(' ', '')
            nombre = str(row[0]).strip().lstrip('. ').strip()
            try:
                val = float(row[2])
                if not pd.isna(val):
                    data[code] = {'valor': val, 'nombre': nombre}
            except Exception:
                pass
        return data
    except Exception as e:
        print(f"    ⚠️  Error leyendo balance: {e}")
        return {}

def compute_atribucion(bal_t0: Path, bal_t1: Path, pn_prev: float) -> dict:
    """
    Calcula la atribución diaria del VCP comparando dos archivos de balance.
    bps = −(valor_T0 − valor_T-1) / PN_prev × 10.000
    El signo negativo porque las cuentas de ingreso tienen saldo negativo.
    """
    if not pn_prev:
        return {'delta_total': 0, 'items': []}

    b0 = read_balance_codes(bal_t0)
    b1 = read_balance_codes(bal_t1)

    if not b0 or not b1:
        return {'delta_total': 0, 'items': []}

    def get_val(bdict, code):
        entry = bdict.get(code)
        return entry['valor'] if entry else None

    def bps(code):
        v0 = get_val(b0, code)
        v1 = get_val(b1, code)
        if v0 is None or v1 is None:
            return None
        return -(v0 - v1) / pn_prev * 10000

    # Total
    delta_total = bps(CODIGO_TOTAL)
    if delta_total is None:
        return {'delta_total': 0, 'items': []}

    # Detalle por cuenta
    items = []
    suma_items = 0.0
    for code, tipo in CUENTAS_ATRIB:
        b = bps(code)
        if b is None:
            continue
        v0 = get_val(b0, code) or 0
        v1 = get_val(b1, code) or 0
        delta_abs = -(v0 - v1)
        # Nombre desde el balance (el que tenga el code en cualquiera de los dos archivos)
        nombre = (b0.get(code) or b1.get(code) or {}).get('nombre', f'Cuenta {code}')
        items.append({
            'n':   nombre,
            't':   tipo,
            'bps': round(b, 4),
            'dp':  round(delta_abs / 1e6, 4),
            'det': f'Código {code}',
            '_code': code,
        })
        suma_items += b

    # Netear Gastos Depositaria A + B + C → uno solo
    DEP_CODES = {'4001001000000000015', '4001001000000000016', '4001001000000000017'}
    dep_items = [i for i in items if i.get('_code') in DEP_CODES]
    if len(dep_items) >= 1:
        net_bps_dep = sum(i['bps'] for i in dep_items)
        net_dp_dep  = sum(i['dp']  for i in dep_items)
        items = [i for i in items if i.get('_code') not in DEP_CODES]
        items.append({
            'n':   'Gastos Soc. Depositaria',
            't':   'Costo',
            'bps': round(net_bps_dep, 4),
            'dp':  round(net_dp_dep,  4),
            'det': 'Neto: Gastos Depositaria Clase A + B + C',
            '_code': '__dep__',
        })
        suma_items = suma_items - sum(i['bps'] for i in dep_items) + net_bps_dep

    # Netear Costo FCI + Ventas FCI → Money Market (FCI)
    FCI_CODES = {'4001001000000000004', '4002001000000000006'}
    fci_items = [i for i in items if i.get('_code') in FCI_CODES]
    if len(fci_items) >= 1:
        net_bps = sum(i['bps'] for i in fci_items)
        net_dp  = sum(i['dp']  for i in fci_items)
        items = [i for i in items if i.get('_code') not in FCI_CODES]
        if abs(net_bps) > 0.0001:
            items.append({
                'n':   'Money Market (FCI)',
                't':   'Devengo',
                'bps': round(net_bps, 4),
                'dp':  round(net_dp,  4),
                'det': 'Neto: Ventas FCI − Costo FCI',
            })
            suma_items = suma_items - sum(i['bps'] for i in fci_items) + net_bps
    # Limpiar campo interno
    for i in items:
        i.pop('_code', None)

    # Residual (diferencia entre total y suma de items)
    residual = round(delta_total - suma_items, 4)
    if abs(residual) > 0.01:
        items.append({
            'n':   'Residual / ajustes contables',
            't':   'Otros',
            'bps': residual,
            'dp':  None,
            'det': 'Diferencia entre RESULTADOS total y cuentas desagregadas',
        })

    return {
        'delta_total': round(delta_total, 4),
        'pn_prev':     round(pn_prev),
        'items':       items,
    }

# ─── PARSER GESTIÓN ──────────────────────────────────────────────────────────

def parse_gestion(path):
    xl = pd.ExcelFile(path)
    print(f"  Hojas: {xl.sheet_names}")

    df = pd.read_excel(path, sheet_name='GESTION MOMENTUM', header=None)

    # Fecha
    fecha_actual = None
    for i in range(10):
        m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', str(df.iloc[i, 0]))
        if m:
            p = m.group(1).split('/')
            fecha_actual = f"{int(p[0]):02d}/{int(p[1]):02d}/{p[2]}"
            break

    # Balance
    activo_total    = find_value(df, 'ACTIVO')
    pasivo_total    = find_value(df, 'PASIVO')
    patrimonio_neto = find_value(df, 'Valor Patrimonial del Fondo')
    creditos        = find_value(df, 'CREDITOS')
    liq_corriente   = find_value(df, 'Liquidez Corriente (3)', 'Liquidez Corriente')
    liq_efectiva    = find_value(df, 'Liquidez Efectiva (2)', 'Liquidez Efectiva')
    provisiones     = find_value(df, 'PROVISIONES')
    prov_dep        = find_value(df, 'PROV. HONORARIOS SOC. DEPOSITARIA')
    prov_ger        = find_value(df, 'PROV. HONORARIOS SOCIEDAD GERENTE')
    prov_gast       = find_value(df, 'PROV. GASTOS SOCIEDAD DEPOSITARIA')
    rescates        = find_value(df, 'RESCATES')
    prevision       = find_value(df, 'PREVISIÓN INCOBRABLES', 'PREVISION INCOBRABLES')
    deudas_pesos    = find_value(df, 'DEUDAS EN PESOS')

    lc_ratio = round(abs(activo_total) / abs(pasivo_total), 1) if activo_total and pasivo_total else None

    # Duration
    duration = None
    dur_idx = find_row_contains(df, 'Duration del Fondo')
    if dur_idx is not None:
        m = re.search(r'(\d+[\.,]?\d*)', str(df.iloc[dur_idx, 0]))
        if m:
            duration = float(m.group(1).replace(',', '.'))

    # Inversiones por sección
    tp_total    = section_subtotal(df, 'TITULO PUBLICO',           4)
    on_total    = section_subtotal(df, 'OBLIGACIONES NEGOCIABLES', 4)
    ff_total    = section_subtotal(df, 'FIDEICOMISOS FINANCIEROS', 4)
    mm_total    = section_subtotal(df, 'MONEY MARKET',             4)
    prov_total  = section_subtotal(df, 'TITULO PROVINCIAL',        4)
    pases_total = section_subtotal(df, 'Aperturas, Pases y Cauciones', 5)
    cpd_usd_total = section_subtotal(df, 'Cheques de Pago Diferido en Dolar', 5)

    tp_det   = section_details(df, 'TITULO PUBLICO',           0, 4, {'cierre': 3, 'cantidad': 2})
    on_det   = section_details(df, 'OBLIGACIONES NEGOCIABLES', 0, 4, {'cierre': 3, 'cantidad': 2})
    ff_det   = section_details(df, 'FIDEICOMISOS FINANCIEROS', 0, 4, {'cierre': 3, 'cantidad': 2})
    prov_det = section_details(df, 'TITULO PROVINCIAL',        0, 4, {'cierre': 3, 'cantidad': 2})
    mm_det   = section_details(df, 'MONEY MARKET',             0, 4, {'cierre': 3})
    fut_det  = section_details(df, 'FUTUROS en Pesos',         0, 4, {'cantidad': 2, 'cierre': 3})

    # CPD desde Reporte de Cheques
    df_chq = pd.read_excel(path, sheet_name='Reporte de Cheques', header=None)
    chq = df_chq.iloc[3:].reset_index(drop=True)
    chq['fecha_vto'] = pd.to_datetime(chq.iloc[:, 2], errors='coerce')
    chq['monto']     = chq.iloc[:, 9].apply(safe)
    today = datetime.now().date()
    validos = chq.dropna(subset=['fecha_vto', 'monto'])
    validos = validos[validos['fecha_vto'].dt.date >= today]
    por_fecha = validos.groupby('fecha_vto')['monto'].sum().sort_index()
    cpd_pesos_total = round(chq.dropna(subset=['monto'])['monto'].sum())

    cpd_venc = []
    dias_hab = 0
    for fdt, monto in por_fecha.items():
        if dias_hab >= 15:
            break
        d = fdt.date() if hasattr(fdt, 'date') else fdt
        if d.weekday() >= 5:
            continue
        cpd_venc.append({'fecha': d.strftime('%d/%m'), 'dia': DIAS_ES[d.weekday()], 'monto': round(monto)})
        dias_hab += 1

    # ── Diversificación de libradores ─────────────────────────────────────
    lib_map = {}
    for _, row in chq.iterrows():
        nombre  = str(row.iloc[0]).strip() if hasattr(row, 'iloc') else ''
        # La hoja tiene los datos en el DataFrame 'chq' ya filtrado
        # Necesitamos releer sin filtro de fecha para tener todos los cheques
        pass

    # Releer sin filtro de fecha para obtener todos los cheques activos
    df_lib = pd.read_excel(path, sheet_name='Reporte de Cheques', header=None)
    lib_rows = df_lib.iloc[3:].reset_index(drop=True)
    lib_map = {}
    for _, row in lib_rows.iterrows():
        nombre = str(row.iloc[0]).strip()
        if not nombre or nombre in ('nan','None','') or nombre.startswith('Cheque'):
            continue
        monto = safe(row.iloc[9])   # Col J: Importe Actual Moneda Fondo
        tasa  = safe(row.iloc[15])  # Col P: Tasa de Descuento
        pct   = safe(row.iloc[22])  # Col W: % sobre VP (decimal)
        if monto is None or monto <= 0:
            continue
        if nombre not in lib_map:
            lib_map[nombre] = {'monto': 0, 'tasa_w': 0, 'pct_vp': 0, 'n': 0}
        lib_map[nombre]['monto']  += monto
        lib_map[nombre]['tasa_w'] += (tasa or 0) * monto  # para promedio ponderado
        lib_map[nombre]['pct_vp'] += (pct or 0) * 100     # decimal → %
        lib_map[nombre]['n']      += 1

    libradores = sorted([
        {
            'nombre':   k,
            'monto':    round(v['monto']),
            'tasa_prom': round(v['tasa_w'] / v['monto'], 2) if v['monto'] else 0,
            'pct_vp':   round(v['pct_vp'], 4),
            'n_cheques': v['n'],
        }
        for k, v in lib_map.items()
    ], key=lambda x: -x['monto'])
    print(f"  Libradores: {len(libradores)} únicos")

    # VCP por clase
    df_cp  = pd.read_excel(path, sheet_name='Posicion cuotapartista', header=None)
    datos  = df_cp.iloc[3:].reset_index(drop=True)

    # Cuotapartistas externos (excluye MOMENTUM)
    cuotapartistas_ext = []
    for _, row in datos.iterrows():
        nombre = str(row.iloc[1]).strip()
        if not nombre or nombre in ('nan', 'None', ''):
            continue
        if 'MOMENTUM' in nombre.upper():
            continue
        clase_cp = str(row.iloc[2]).strip()
        if clase_cp not in ('A', 'B', 'C'):
            continue
        inv = safe(row.iloc[4])
        qty = safe(row.iloc[3])
        if inv is None:
            continue
        # Limpiar nombre
        nombre_limpio = nombre.replace(' (ACDI)', '').replace(' S.A.', '').strip()
        cuotapartistas_ext.append({
            'nombre': nombre_limpio,
            'clase':  clase_cp,
            'inversion': round(inv),
            'cuotapartes': round(qty) if qty else None,
        })

    vcp_data = {}
    for clase in ['A', 'B', 'C']:
        mask = datos.iloc[:, 2].astype(str).str.strip() == clase
        rows = datos[mask]
        if rows.empty:
            continue
        vcp_val    = safe(rows.iloc[0, 8])
        saldo_cp   = rows.iloc[:, 3].apply(safe).dropna().sum()   # cantidad cuotapartes
        inversion  = rows.iloc[:, 4].apply(safe).dropna().sum()   # monto invertido en $

        # Ingresos/egresos/saldo desde hoja principal (busca por nombre)
        ing_hoy  = find_value(df, f'INGRESOS EN EL DIA Clase {clase}')
        egr_hoy  = find_value(df, f'EGRESOS EN EL DIA Clase {clase}')
        saldo_gs = find_value(df, f'SALDO TOTAL DE CUOTAPARTES Clase {clase}')

        vcp_data[clase] = {
            'vcp':           round(vcp_val * 1000, 3) if vcp_val else None,
            'vcp_prev':      None,
            'vcp_abril':     None,
            'vcp_dic':       None,
            'saldo_cp':      round(saldo_gs) if saldo_gs else (round(saldo_cp) if saldo_cp else None),
            'inversion':     round(inversion) if inversion else None,   # monto en pesos
            'inversion_prev':None,
            'saldo_cp_prev': None,
            'ingresos_hoy':  round(ing_hoy) if ing_hoy else 0,
            'egresos_hoy':   round(egr_hoy) if egr_hoy else 0,
            'ingresos_prev': None,
            'egresos_prev':  None,
        }

    # Cruzar con día anterior
    prev = find_prev_json(fecha_actual)
    fecha_anterior = None
    activo_prev = pasivo_prev = pn_prev = cred_prev = None

    if prev:
        fecha_anterior = prev.get('meta', {}).get('fecha_actual')
        pp = prev.get('patrimonial', {})
        activo_prev = pp.get('activo_total')
        pasivo_prev = pp.get('pasivo_total')
        pn_prev     = pp.get('patrimonio_neto')
        cred_prev   = pp.get('creditos_totales')
        prev_ca     = prev.get('cartera', {})
        prev_cp     = prev.get('cuotapartes', {})

        def match_prev(det_list, prev_list):
            if not prev_list:
                return det_list
            prev_map = {r['nombre']: r['total'] for r in prev_list}
            for row in det_list:
                row['total_prev'] = prev_map.get(row['nombre'])
            return det_list

        tp_det   = match_prev(tp_det,   prev_ca.get('tp_det',   []))
        on_det   = match_prev(on_det,   prev_ca.get('on_det',   []))
        ff_det   = match_prev(ff_det,   prev_ca.get('ff_det',   []))
        prov_det = match_prev(prov_det, prev_ca.get('prov_det', []))

        for clase in ['A', 'B', 'C']:
            if clase in vcp_data and clase in prev_cp:
                vcp_data[clase]['vcp_prev']      = prev_cp[clase].get('vcp')
                vcp_data[clase]['saldo_cp_prev'] = prev_cp[clase].get('saldo_cp')
                vcp_data[clase]['vcp_abril']     = prev_cp[clase].get('vcp_abril')
                vcp_data[clase]['vcp_dic']       = prev_cp[clase].get('vcp_dic')
                vcp_data[clase]['ingresos_prev']  = prev_cp[clase].get('ingresos_hoy')
                vcp_data[clase]['egresos_prev']   = prev_cp[clase].get('egresos_hoy')
                vcp_data[clase]['inversion_prev'] = prev_cp[clase].get('inversion')
    else:
        prev_ca = {}

    return {
        'meta': {
            'fecha_actual':   fecha_actual,
            'fecha_anterior': fecha_anterior,
            'duration_dias':  duration,
            'participes':     len(vcp_data),
        },
        'patrimonial': {
            'activo_total':          activo_total,
            'activo_total_prev':     activo_prev,
            'pasivo_total':          pasivo_total,
            'pasivo_total_prev':     pasivo_prev,
            'patrimonio_neto':       patrimonio_neto,
            'patrimonio_neto_prev':  pn_prev,
            'creditos_totales':      creditos,
            'creditos_totales_prev': cred_prev,
            'liquidez_corriente':    lc_ratio,
            'liquidez_efectiva':     liq_efectiva,
        },
        'balance_detalle': {
            'provisiones': provisiones, 'prov_depositaria': prov_dep,
            'prov_gerente': prov_ger,   'prov_gastos': prov_gast,
            'rescates': rescates,       'prevision_incobrables': prevision,
            'deudas_pesos': deudas_pesos,
        },
        'cartera': {
            'cpd_pesos':              cpd_pesos_total,
            'cpd_pesos_prev':         prev_ca.get('cpd_pesos'),
            'cpd_usd':                cpd_usd_total,
            'cpd_usd_prev':           prev_ca.get('cpd_usd'),
            'pases_cauciones':        pases_total,
            'pases_cauciones_prev':   prev_ca.get('pases_cauciones'),
            'money_market':           mm_total,
            'money_market_prev':      prev_ca.get('money_market'),
            'titulos_publicos':       tp_total,
            'titulos_publicos_prev':  prev_ca.get('titulos_publicos'),
            'obligaciones_neg':       on_total,
            'obligaciones_neg_prev':  prev_ca.get('obligaciones_neg'),
            'fideicomisos':           ff_total,
            'fideicomisos_prev':      prev_ca.get('fideicomisos'),
            'titulo_provincial':      prov_total,
            'titulo_provincial_prev': prev_ca.get('titulo_provincial'),
            'tp_det': tp_det, 'on_det': on_det,
            'ff_det': ff_det, 'prov_det': prov_det,
            'money_market_det': mm_det,
        },
        'futuros': [
            {'contrato': f['nombre'], 'cantidad': f.get('cantidad'),
             'cierre': f.get('cierre'), 'total_hoy': f['total'], 'total_prev': None}
            for f in fut_det
        ],
        'cuotapartes':      vcp_data,
        'cuotapartistas_ext': cuotapartistas_ext,
        'libradores': libradores,
        'cpd_vencimientos': cpd_venc,
        'atribucion':       {'delta_total': 0, 'items': []},  # se llena después
    }


def find_prev_json(fecha_actual):
    jsons = sorted(DATA_DIR.glob('20??-??-??.json'), reverse=True)
    for j in jsons:
        try:
            data = json.loads(j.read_text(encoding='utf-8'))
            f = data.get('meta', {}).get('fecha_actual')
            if f and f != fecha_actual:
                return data
        except Exception:
            pass
    return None


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    gestiones = {}
    balances  = {}

    for f in sorted(RAW_DIR.glob('*')):
        if not f.is_file():
            continue
        nl = f.name.lower()
        ds = date_from_filename(f.stem)
        if not ds:
            continue
        if 'gestion' in nl or 'momentum' in nl:
            gestiones[ds] = f
        elif 'balance' in nl:
            balances[ds] = f

    if not gestiones:
        print('No se encontraron archivos GESTION en raw/')
        return

    print(f"Gestiones: {list(gestiones.keys())}")
    print(f"Balances:  {list(balances.keys())}")

    for ds, path in sorted(gestiones.items()):
        print(f'\n📅 {ds} — {path.name}')
        try:
            result = parse_gestion(path)

            # Atribución desde balance si hay T0 y T-1 disponibles
            fecha_actual = result['meta'].get('fecha_actual')   # "DD/MM/YYYY"
            if fecha_actual:
                # Convertir "DD/MM/YYYY" → "YYYY-MM-DD" para buscar archivo
                p = fecha_actual.split('/')
                fecha_iso = f"{p[2]}-{p[1]}-{p[0]}"
                # Buscar todos los balances ordenados
                bal_dates = sorted(balances.keys())
                # T0 = balance de la misma fecha que el informe
                # T-1 = balance del día hábil anterior disponible
                if fecha_iso in balances:
                    bal_t0 = balances[fecha_iso]
                    prev_bal = [d for d in bal_dates if d < fecha_iso]
                    if prev_bal:
                        bal_t1 = balances[prev_bal[-1]]
                        pn_prev = result['patrimonial'].get('patrimonio_neto_prev') or \
                                  result['patrimonial'].get('patrimonio_neto')
                        print(f"  Balance T0: {bal_t0.name}")
                        print(f"  Balance T-1: {bal_t1.name}")
                        result['atribucion'] = compute_atribucion(bal_t0, bal_t1, pn_prev)
                        dt = result['atribucion']['delta_total']
                        print(f"  Δ VCP: {dt:+.4f} bps")
                    else:
                        print(f"  ⚠️  Sin balance T-1 para {fecha_iso}")
                else:
                    print(f"  ⚠️  Sin balance T0 para {fecha_iso} (disponibles: {bal_dates})")

            out = DATA_DIR / f'{ds}.json'
            out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            p = result['patrimonial']
            print(f"  Activo: {p['activo_total']:,.0f} | PN: {p['patrimonio_neto']:,.0f}")
            print(f"  ✅ data/{ds}.json")
        except Exception as e:
            import traceback
            print(f'  ❌ {e}')
            traceback.print_exc()

    dates = sorted(p.stem for p in DATA_DIR.glob('20??-??-??.json'))
    (DATA_DIR / 'index.json').write_text(
        json.dumps({'dates': dates, 'latest': dates[-1] if dates else None},
                   ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'\n✅ index.json — {len(dates)} días: {dates}')

if __name__ == '__main__':
    main()
