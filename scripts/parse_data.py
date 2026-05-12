"""
parse_data.py  —  FCI Momentum Pymes
Parser robusto por nombre de sección, inmune a cambios de fila.
"""

import json, re
from pathlib import Path
from datetime import datetime

import pandas as pd

ROOT     = Path(__file__).parent.parent
RAW_DIR  = ROOT / "raw"
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

DIAS_ES = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']

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
    """Devuelve el índice de la primera fila donde col A coincide con algún label."""
    ca = col_a(df).str.upper()
    for label in labels:
        mask = ca == label.upper()
        if mask.any():
            return mask.idxmax()
    return None

def find_row_contains(df, text):
    """Devuelve el índice de la primera fila donde col A contiene el texto."""
    ca = col_a(df).str.upper()
    mask = ca.str.contains(text.upper(), na=False)
    return mask.idxmax() if mask.any() else None

def find_value(df, *labels):
    """Busca label en col A, devuelve col B (índice 1)."""
    idx = find_row(df, *labels)
    return safe(df.iloc[idx, 1]) if idx is not None else None

def section_subtotal(df, section_label, value_col_idx):
    """
    Encuentra el header de sección, luego busca el subtotal:
    la primera fila donde col A está vacía y value_col tiene número.
    """
    start = find_row_contains(df, section_label)
    if start is None:
        return None
    # Busca desde start+2 (salteando header de columnas)
    for i in range(start + 2, min(start + 500, len(df))):
        a = str(df.iloc[i, 0]).strip()
        v = safe(df.iloc[i, value_col_idx])
        if a in ('', 'nan', 'None') and v is not None:
            return v
    return None

def section_details(df, section_label, name_col, value_col, extra_cols=None):
    """
    Extrae filas de detalle de una sección:
    - name_col: índice de columna con nombre del instrumento
    - value_col: índice de columna con el valor total
    - extra_cols: dict {nombre: índice} para columnas adicionales
    Devuelve lista de dicts.
    """
    start = find_row_contains(df, section_label)
    if start is None:
        return []
    rows = []
    for i in range(start + 2, min(start + 500, len(df))):
        name = str(df.iloc[i, name_col]).strip()
        val  = safe(df.iloc[i, value_col])
        if name in ('', 'nan', 'None'):
            if val is not None:
                break  # subtotal — fin de sección
            continue
        if val is None:
            continue
        # Detecta que no es otra sección (evita confundir headers con datos)
        if name.isupper() and len(name) > 20 and val is None:
            break
        row = {'nombre': name, 'total': val}
        if extra_cols:
            for k, idx in extra_cols.items():
                row[k] = safe(df.iloc[i, idx])
        rows.append(row)
    return rows

# ─── PARSER ──────────────────────────────────────────────────────────────────

def parse_gestion(path):
    xl = pd.ExcelFile(path)
    print(f"  Hojas: {xl.sheet_names}")

    df = pd.read_excel(path, sheet_name='GESTION MOMENTUM', header=None)

    # ── FECHA DEL INFORME ────────────────────────────────────────────────────
    fecha_actual = None
    for i in range(10):
        m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', str(df.iloc[i, 0]))
        if m:
            p = m.group(1).split('/')
            fecha_actual = f"{int(p[0]):02d}/{int(p[1]):02d}/{p[2]}"
            break

    # ── BALANCE PRINCIPAL (col A = etiqueta, col B = valor) ─────────────────
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

    lc_ratio = None
    if activo_total and pasivo_total and pasivo_total != 0:
        lc_ratio = round(abs(activo_total) / abs(pasivo_total), 1)

    # ── DURATION ─────────────────────────────────────────────────────────────
    duration = None
    dur_idx = find_row_contains(df, 'Duration del Fondo')
    if dur_idx is not None:
        m = re.search(r'(\d+[\.,]?\d*)', str(df.iloc[dur_idx, 0]))
        if m:
            duration = float(m.group(1).replace(',', '.'))

    # ── INVERSIONES — subtotales por sección ─────────────────────────────────
    # Columna E (índice 4) = Total para la mayoría de secciones
    # Columna F (índice 5) = Total para Pases y CPD
    tp_total   = section_subtotal(df, 'TITULO PUBLICO',          4)
    on_total   = section_subtotal(df, 'OBLIGACIONES NEGOCIABLES', 4)
    ff_total   = section_subtotal(df, 'FIDEICOMISOS FINANCIEROS', 4)
    mm_total   = section_subtotal(df, 'MONEY MARKET',             4)
    prov_total = section_subtotal(df, 'TITULO PROVINCIAL',        4)
    pases_total = section_subtotal(df, 'Aperturas, Pases y Cauciones', 5)
    cpd_usd_total = section_subtotal(df, 'Cheques de Pago Diferido en Dolar', 5)

    # ── INVERSIONES — detalle de instrumentos ────────────────────────────────
    # Títulos Públicos: col A=nombre, col E=total, col D=cierre, col I=vto
    tp_det = section_details(df, 'TITULO PUBLICO', 0, 4,
                             {'cierre': 3, 'cantidad': 2, 'vencimiento': 8})

    # ONs: mismo esquema
    on_det = section_details(df, 'OBLIGACIONES NEGOCIABLES', 0, 4,
                             {'cierre': 3, 'cantidad': 2, 'vencimiento': 8})

    # Fideicomisos
    ff_det = section_details(df, 'FIDEICOMISOS FINANCIEROS', 0, 4,
                             {'cierre': 3, 'cantidad': 2, 'vencimiento': 8})

    # Futuros ROFEX: col A=nombre, col C=cantidad, col D=cierre, col E=total
    fut_det = section_details(df, 'FUTUROS en Pesos', 0, 4,
                              {'cantidad': 2, 'cierre': 3, 'vencimiento': 8})

    # Money Market
    mm_det = section_details(df, 'MONEY MARKET', 0, 4, {'cierre': 3})

    # Provincial
    prov_det = section_details(df, 'TITULO PROVINCIAL', 0, 4,
                               {'cierre': 3, 'cantidad': 2})

    # CPD Pesos (desde Reporte de Cheques)
    df_chq = pd.read_excel(path, sheet_name='Reporte de Cheques', header=None)
    chq = df_chq.iloc[3:].reset_index(drop=True)
    chq['fecha_vto'] = pd.to_datetime(chq.iloc[:, 2], errors='coerce')
    chq['monto']     = chq.iloc[:, 9].apply(safe)
    today = datetime.now().date()
    validos = chq.dropna(subset=['fecha_vto', 'monto'])
    validos = validos[validos['fecha_vto'].dt.date >= today]
    por_fecha = validos.groupby('fecha_vto')['monto'].sum().sort_index()
    cpd_pesos_total = round(chq.dropna(subset=['monto'])['monto'].sum())

    # CPD vencimientos próximos 15 días hábiles
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

    # ── VCP Y SALDOS POR CLASE ────────────────────────────────────────────────
    df_cp   = pd.read_excel(path, sheet_name='Posicion cuotapartista', header=None)
    datos   = df_cp.iloc[3:].reset_index(drop=True)
    vcp_data = {}
    for clase in ['A', 'B', 'C']:
        mask = datos.iloc[:, 2].astype(str).str.strip() == clase
        rows = datos[mask]
        if rows.empty:
            continue
        vcp_val  = safe(rows.iloc[0, 8])
        saldo_cp = rows.iloc[:, 3].apply(safe).dropna().sum()
        vcp_data[clase] = {
            'vcp':           round(vcp_val * 1000, 3) if vcp_val else None,
            'vcp_prev':      None,
            'vcp_abril':     None,
            'vcp_dic':       None,
            'saldo_cp':      round(saldo_cp) if saldo_cp else None,
            'saldo_cp_prev': None,
            'ingresos_hoy':  None,
            'egresos_hoy':   None,
            'ingresos_prev': None,
            'egresos_prev':  None,
        }

    # ── CRUZAR CON DÍA ANTERIOR ───────────────────────────────────────────────
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
        prev_cp = prev.get('cuotapartes', {})
        for clase in ['A', 'B', 'C']:
            if clase in vcp_data and clase in prev_cp:
                vcp_data[clase]['vcp_prev']      = prev_cp[clase].get('vcp')
                vcp_data[clase]['saldo_cp_prev'] = prev_cp[clase].get('saldo_cp')
                vcp_data[clase]['vcp_abril']     = prev_cp[clase].get('vcp_abril')
                vcp_data[clase]['vcp_dic']       = prev_cp[clase].get('vcp_dic')
    else:
        prev_ca = {}

    # ── JSON FINAL ────────────────────────────────────────────────────────────
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
            'provisiones':           provisiones,
            'prov_depositaria':      prov_dep,
            'prov_gerente':          prov_ger,
            'prov_gastos':           prov_gast,
            'rescates':              rescates,
            'prevision_incobrables': prevision,
            'deudas_pesos':          deudas_pesos,
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
            'money_market_det':       mm_det,
            'tp_det':                 tp_det,
            'on_det':                 on_det,
            'ff_det':                 ff_det,
            'prov_det':               prov_det,
        },
        'futuros': [
            {
                'contrato':   f['nombre'],
                'cantidad':   f.get('cantidad'),
                'cierre':     f.get('cierre'),
                'total_hoy':  f['total'],
                'total_prev': None,
            }
            for f in fut_det
        ],
        'cuotapartes':      vcp_data,
        'cpd_vencimientos': cpd_venc,
        'atribucion':       {'delta_total': 0, 'items': []},
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
    archivos = {}
    for f in sorted(RAW_DIR.glob('*')):
        if not f.is_file():
            continue
        nl = f.name.lower()
        if 'gestion' in nl or 'momentum' in nl:
            ds = date_from_filename(f.stem)
            if ds:
                archivos[ds] = f
            else:
                print(f'  ⚠️  Sin fecha en "{f.name}"')

    if not archivos:
        print('No se encontraron archivos en raw/')
        return

    for ds, path in sorted(archivos.items()):
        print(f'\n📅 {ds} — {path.name}')
        try:
            result = parse_gestion(path)
            out = DATA_DIR / f'{ds}.json'
            out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')

            # Resumen de lo extraído
            m = result['meta']
            p = result['patrimonial']
            c = result['cartera']
            print(f"  Fecha:    {m['fecha_actual']}  |  Anterior: {m['fecha_anterior']}")
            print(f"  Activo:   {p['activo_total']:,.0f}")
            print(f"  PN:       {p['patrimonio_neto']:,.0f}")
            print(f"  Duration: {m['duration_dias']} días")
            print(f"  CPD $:    {c['cpd_pesos']:,.0f}")
            print(f"  CPD USD:  {c['cpd_usd']}")
            print(f"  Pases:    {c['pases_cauciones']}")
            print(f"  MM:       {c['money_market']}")
            print(f"  TP:       {c['titulos_publicos']}")
            print(f"  ONs:      {c['obligaciones_neg']}")
            print(f"  FF:       {c['fideicomisos']}")
            print(f"  Futuros:  {len(result['futuros'])} contratos")
            print(f"  VCP A:    {result['cuotapartes'].get('A', {}).get('vcp')}")
            print(f"  CPD venc: {len(result['cpd_vencimientos'])} días")
            print(f"  ✅ data/{ds}.json")
        except Exception as e:
            import traceback
            print(f'  ❌ {e}')
            traceback.print_exc()

    dates = sorted(p.stem for p in DATA_DIR.glob('20??-??-??.json'))
    (DATA_DIR / 'index.json').write_text(
        json.dumps({'dates': dates, 'latest': dates[-1] if dates else None},
                   ensure_ascii=False, indent=2),
        encoding='utf-8'
    )
    print(f'\n✅ index.json — {len(dates)} días: {dates}')

if __name__ == '__main__':
    main()
