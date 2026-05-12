"""
parse_data.py  —  FCI Momentum Pymes
=====================================
Parser robusto que busca datos por NOMBRE DE ETIQUETA, no por posición de celda.
Inmune a cambios de orden de filas.

Hojas detectadas en GESTION_MOMENTUM_YYYYMMDD.xlsx:
  - "GESTION MOMENTUM"    → balance (col A = etiqueta, col B = valor)
  - "Reporte de Cheques"  → vencimientos CPD (col C = fecha, col J = monto fondo)
  - "Posicion cuotapartista" → VCP y saldos por clase A/B/C
"""

import json, re
from pathlib import Path
from datetime import datetime, date

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

def find_value(df, *labels):
    """Busca en col 0 por nombre (case-insensitive), devuelve col 1."""
    col0 = df.iloc[:, 0].astype(str).str.strip().str.upper()
    for label in labels:
        mask = col0 == label.upper().strip()
        if mask.any():
            return safe(df.loc[mask.idxmax(), df.columns[1]])
    return None

# ─── PARSER ──────────────────────────────────────────────────────────────────

def parse_gestion(path):
    xl = pd.ExcelFile(path)
    print(f"  Hojas: {xl.sheet_names}")

    # ── 1. GESTION MOMENTUM: balance principal ────────────────────────────
    df = pd.read_excel(path, sheet_name='GESTION MOMENTUM', header=None)

    # Fecha del informe (busca "Informe de Gestión al DD/MM/YYYY" en primeras 10 filas)
    fecha_actual = None
    for i in range(10):
        m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', str(df.iloc[i, 0]))
        if m:
            p = m.group(1).split('/')
            fecha_actual = f"{int(p[0]):02d}/{int(p[1]):02d}/{p[2]}"
            break

    activo_total     = find_value(df, 'ACTIVO')
    pasivo_total     = find_value(df, 'PASIVO')
    patrimonio_neto  = find_value(df, 'Valor Patrimonial del Fondo')
    creditos         = find_value(df, 'CREDITOS')
    inversiones      = find_value(df, 'INVERSIONES')
    titulos_pub      = find_value(df, 'TITULOS PUBLICOS')
    futuros_val      = find_value(df, 'FUTUROS')
    disponib         = find_value(df, 'DISPONIBILIDADES')
    provisiones      = find_value(df, 'PROVISIONES')
    prov_dep         = find_value(df, 'PROV. HONORARIOS SOC. DEPOSITARIA')
    prov_ger         = find_value(df, 'PROV. HONORARIOS SOCIEDAD GERENTE')
    prov_gast        = find_value(df, 'PROV. GASTOS SOCIEDAD DEPOSITARIA')
    rescates         = find_value(df, 'RESCATES')
    prevision        = find_value(df, 'PREVISIÓN INCOBRABLES', 'PREVISION INCOBRABLES')
    deudas_pesos     = find_value(df, 'DEUDAS EN PESOS')
    liq_corriente    = find_value(df, 'Liquidez Corriente (3)', 'Liquidez Corriente')
    liq_efectiva     = find_value(df, 'Liquidez Efectiva (2)', 'Liquidez Efectiva')

    lc_ratio = None
    if activo_total and pasivo_total and pasivo_total != 0:
        lc_ratio = round(abs(activo_total) / abs(pasivo_total), 1)

    # ── 2. POSICION CUOTAPARTISTA: VCP y saldos ───────────────────────────
    df_cp = pd.read_excel(path, sheet_name='Posicion cuotapartista', header=None)
    datos = df_cp.iloc[3:].reset_index(drop=True)

    vcp_data = {}
    for clase in ['A', 'B', 'C']:
        mask  = datos.iloc[:, 2].astype(str).str.strip() == clase
        rows  = datos[mask]
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

    # ── 3. REPORTE DE CHEQUES: vencimientos CPD agrupados por fecha ───────
    df_chq = pd.read_excel(path, sheet_name='Reporte de Cheques', header=None)
    chq    = df_chq.iloc[3:].reset_index(drop=True)

    chq['fecha_vto'] = pd.to_datetime(chq.iloc[:, 2], errors='coerce')
    chq['monto']     = chq.iloc[:, 9].apply(safe)   # Importe Actual Moneda Fondo

    today = datetime.now().date()
    validos = chq.dropna(subset=['fecha_vto', 'monto'])
    validos = validos[validos['fecha_vto'].dt.date >= today]

    por_fecha = (
        validos.groupby('fecha_vto')['monto']
        .sum()
        .sort_index()
    )

    cpd_total = round(chq.dropna(subset=['monto'])['monto'].sum())

    cpd_venc = []
    dias_hab = 0
    for fdt, monto in por_fecha.items():
        if dias_hab >= 15:
            break
        d = fdt.date() if hasattr(fdt, 'date') else fdt
        if d.weekday() >= 5:
            continue
        cpd_venc.append({
            'fecha': d.strftime('%d/%m'),
            'dia':   DIAS_ES[d.weekday()],
            'monto': round(monto),
        })
        dias_hab += 1

    # ── 4. CRUZAR CON DÍA ANTERIOR ───────────────────────────────────────
    prev = find_prev_json(fecha_actual)
    fecha_anterior = None
    activo_prev = pasivo_prev = pn_prev = cred_prev = None

    if prev:
        fecha_anterior = prev.get('meta', {}).get('fecha_actual')
        prev_pat = prev.get('patrimonial', {})
        activo_prev  = prev_pat.get('activo_total')
        pasivo_prev  = prev_pat.get('pasivo_total')
        pn_prev      = prev_pat.get('patrimonio_neto')
        cred_prev    = prev_pat.get('creditos_totales')
        prev_cp = prev.get('cuotapartes', {})
        for clase in ['A', 'B', 'C']:
            if clase in vcp_data and clase in prev_cp:
                vcp_data[clase]['vcp_prev']      = prev_cp[clase].get('vcp')
                vcp_data[clase]['saldo_cp_prev'] = prev_cp[clase].get('saldo_cp')
                vcp_data[clase]['vcp_abril']     = prev_cp[clase].get('vcp_abril')
                vcp_data[clase]['vcp_dic']       = prev_cp[clase].get('vcp_dic')

    # ── 5. JSON FINAL ─────────────────────────────────────────────────────
    return {
        'meta': {
            'fecha_actual':   fecha_actual,
            'fecha_anterior': fecha_anterior,
            'duration_dias':  None,
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
            'inversiones':           inversiones,
            'titulos_publicos':      titulos_pub,
            'futuros':               futuros_val,
            'disponibilidades':      disponib,
            'provisiones':           provisiones,
            'prov_depositaria':      prov_dep,
            'prov_gerente':          prov_ger,
            'prov_gastos':           prov_gast,
            'rescates':              rescates,
            'prevision_incobrables': prevision,
            'deudas_pesos':          deudas_pesos,
        },
        'cartera': {
            'cpd_pesos':             cpd_total,
            'cpd_pesos_prev':        None,
            'cpd_usd':               None,
            'cpd_usd_prev':          None,
            'pases_cauciones':       None,
            'pases_cauciones_prev':  None,
            'money_market':          None,
            'money_market_prev':     None,
        },
        'cuotapartes':      vcp_data,
        'cpd_vencimientos': cpd_venc,
        'futuros':          [],
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
            print(f'  ✅ data/{ds}.json')
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
