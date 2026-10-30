import os
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st
import requests

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(page_title="Dashboard Renta Fija - Break-Even", layout="wide")

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# Overrides manuales de VPV (idealmente pasarlos a la columna 'VPV' de la hoja Lecaps)
VPV_OVERRIDE = {'S29E7': 111.6660}

# =========================================================
# CALENDARIO (compartido)
# =========================================================
def es_habil(f, feriados):
    return f.weekday() < 5 and f not in feriados

def proximo_habil(f, feriados):
    actual = f
    while not es_habil(actual, feriados):
        actual += timedelta(days=1)
    return actual

def calcular_fecha_liq(f_op, plazo_t, feriados):
    actual = f_op
    sumados = 0
    while sumados < plazo_t:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            sumados += 1
    return proximo_habil(actual, feriados)

def restar_dias_habiles(f, dias_a_restar, feriados):
    actual = f
    restados = 0
    while restados < dias_a_restar:
        actual -= timedelta(days=1)
        if es_habil(actual, feriados):
            restados += 1
    return actual

def contar_dias_habiles(f_inicio, f_fin, feriados):
    """Cuenta días hábiles en el intervalo (f_inicio, f_fin]: excluye inicio, incluye fin."""
    actual = f_inicio
    dias = 0
    while actual < f_fin:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            dias += 1
    return dias

def _es_ultimo_dia_feb(f):
    return f.month == 2 and (f + timedelta(days=1)).month == 3

def dias360_excel(f_inicio, f_fin):
    """DAYS360 método US (NASD), igual al default de Excel."""
    d1, m1, y1 = f_inicio.day, f_inicio.month, f_inicio.year
    d2, m2, y2 = f_fin.day, f_fin.month, f_fin.year
    if _es_ultimo_dia_feb(f_inicio) and _es_ultimo_dia_feb(f_fin):
        d2 = 30
    if _es_ultimo_dia_feb(f_inicio):
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    if d1 == 31:
        d1 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)

def normalizar_precio(x):
    """Precios pueden venir como 1.05 o 105. Devuelve base 100. NaN-safe."""
    if x is None or pd.isna(x):
        return np.nan
    x = float(x)
    return x * 100 if x < 10 else x

def _k(ticker):
    """Clave normalizada de ticker (la API devuelve en mayúsculas)."""
    return str(ticker).strip().upper()

# =========================================================
# API 1816 (compartida)
# =========================================================
def _get_api_key():
    try:
        k = st.secrets["API_1816_KEY"]
        if k:
            return k
    except Exception:
        pass
    return os.environ.get("API_1816_KEY")

@st.cache_resource(ttl=3000)
def _token_1816(api_key):
    """Token reutilizado ~50 min: el plan permite máx. 50 tokens por día (compartidos con el Action)."""
    auth = requests.post("https://api.1816.com.ar/v1/auth/token",
                         json={"apiKey": api_key, "module": "mercado"}, timeout=10)
    auth.raise_for_status()
    return auth.json()["token"]

@st.cache_data(ttl=300)
def obtener_precios_1816(lista_tickers, api_key):
    """Lanza excepción si algo falla (así no se cachea el error)."""
    url = "https://api.1816.com.ar/v1/mercado/indicadores"
    params = {"tickers": [t.strip().lower() for t in lista_tickers], "campos": ["precioDirty"]}
    response = requests.get(url, headers={"Authorization": f"Bearer {_token_1816(api_key)}"},
                            params=params, timeout=10)
    if response.status_code in (401, 403):  # token vencido: se pide uno nuevo una sola vez
        _token_1816.clear()
        response = requests.get(url, headers={"Authorization": f"Bearer {_token_1816(api_key)}"},
                                params=params, timeout=10)
    response.raise_for_status()
    data = response.json()

    precios = {}
    if isinstance(data, dict) and "instrumentos" in data:
        for ticker, valores in data["instrumentos"].items():
            if isinstance(valores, dict) and valores.get("precioDirty") is not None:
                precios[ticker.upper()] = float(valores["precioDirty"])
    if not precios:
        raise ValueError("La API respondió sin precioDirty para ningún ticker.")
    return precios, datetime.now(TZ_AR)  # se cachea junto con la hora de la consulta

# =========================================================
# CARGA DE DATOS
# =========================================================
def ruta_inputs():
    """Ruta del Excel único con toda la info (acepta variantes de mayúsculas: Linux distingue)."""
    for nombre in ("Inputs.xlsx", "inputs.xlsx", "INPUTS.xlsx"):
        if os.path.exists(nombre):
            return nombre
    return None

ARCHIVO_CIERRES = "precios_1816.csv"  # lo escribe guardar_cierres_be.py (GitHub Actions)
FUENTE_HIST = "Excel + cierres 1816" if os.path.exists(ARCHIVO_CIERRES) else "Excel"

def firma_archivos():
    """Cambia cuando se modifica el Excel o llega un cierre nuevo: invalida la caché de datos."""
    return tuple(os.path.getmtime(p) if p and os.path.exists(p) else 0
                 for p in (ruta_inputs(), ARCHIVO_CIERRES))

def combinar_precios(df_excel):
    """Hoja Precios del Excel + cierres diarios de 1816. Donde el Excel tiene precio, manda el Excel;
    el CSV completa las fechas (y especies) que falten."""
    if not os.path.exists(ARCHIVO_CIERRES):
        return df_excel
    df_api = pd.read_csv(ARCHIVO_CIERRES, parse_dates=['Fecha'])
    if df_api.empty:
        return df_excel
    a = df_excel.copy()
    a['Fecha'] = pd.to_datetime(a['Fecha'], errors='coerce').dt.normalize()
    a = a.dropna(subset=['Fecha']).drop_duplicates('Fecha', keep='last').set_index('Fecha')
    df_api['Fecha'] = df_api['Fecha'].dt.normalize()
    b = df_api.drop_duplicates('Fecha', keep='last').set_index('Fecha')
    b.columns = [_k(c) for c in b.columns]
    out = a.combine_first(b)
    cols_excel = [c for c in a.columns]
    out = out[cols_excel + [c for c in out.columns if c not in cols_excel]]  # orden del Excel primero
    return out.reset_index()

@st.cache_data
def cargar_datos_cer(firma):
    """firma: ver firma_archivos() (solo sirve para invalidar la caché)."""
    file_path = ruta_inputs()
    if file_path is None:
        return None

    xls = pd.ExcelFile(file_path)
    df_lec = pd.read_excel(xls, sheet_name='Lecaps')
    df_cer = pd.read_excel(xls, sheet_name='LECER')
    df_precios = pd.read_excel(xls, sheet_name='Precios')
    df_indice = pd.read_excel(xls, sheet_name='CER')
    df_feriados = pd.read_excel(xls, sheet_name='Feriados')

    for d in (df_lec, df_cer, df_precios, df_indice):
        d.columns = d.columns.str.strip()

    # Hoja Tamar (opcional)
    df_tamar = pd.DataFrame()
    hojas_tamar = [h for h in xls.sheet_names if 'TAMAR' in h.upper()]
    if hojas_tamar:
        df_tamar = pd.read_excel(xls, sheet_name=hojas_tamar[0])
        df_tamar.columns = df_tamar.columns.str.strip()
        col_f = next((c for c in df_tamar.columns if 'FECHA' in c.upper()), df_tamar.columns[0])
        col_v = next((c for c in df_tamar.columns if c != col_f), df_tamar.columns[-1])
        df_tamar['Fecha_dt'] = pd.to_datetime(df_tamar[col_f]).dt.date
        df_tamar['Tamar_val'] = pd.to_numeric(df_tamar[col_v], errors='coerce')
        df_tamar['Tamar_pct'] = df_tamar['Tamar_val'].apply(lambda v: v * 100 if pd.notnull(v) and v < 1.0 else v)
        df_tamar = df_tamar.sort_values('Fecha_dt').reset_index(drop=True)

    if 'Tamar' in df_indice.columns and 'CER' not in df_indice.columns:
        df_indice.rename(columns={'Tamar': 'CER'}, inplace=True)

    feriados = set(pd.to_datetime(df_feriados["Feriados"]).dt.date)
    df_indice['Fecha'] = pd.to_datetime(df_indice['Fecha']).dt.date
    df_indice = df_indice.sort_values('Fecha').reset_index(drop=True)

    # Limpieza de la curva CER: descartar filas constantes repetidas al final
    diffs = df_indice['CER'].diff().abs()
    if not diffs.empty and (diffs > 0).any():
        last_changing_idx = df_indice[diffs > 0].index[-1]
        df_indice_clean = df_indice.loc[:last_changing_idx].copy()
    else:
        df_indice_clean = df_indice.copy()

    df_precios = combinar_precios(df_precios)
    df_precios['Fecha_dt'] = pd.to_datetime(df_precios['Fecha']).dt.date
    df_precios = df_precios.sort_values('Fecha_dt').reset_index(drop=True)

    df_lec['Vencimiento_dt'] = pd.to_datetime(df_lec['Vencimiento']).dt.date
    df_cer['Vencimiento_dt'] = pd.to_datetime(df_cer['Vencimiento']).dt.date

    pares = []
    for _, row_lec in df_lec.iterrows():
        vto = row_lec['Vencimiento_dt']
        matching_cer = df_cer[df_cer['Vencimiento_dt'] == vto]
        for _, row_cer in matching_cer.iterrows():
            pares.append({
                'label': f"{row_lec['Ticker']} vs {row_cer['Ticker']} (Vto: {vto.strftime('%d/%m/%Y')})",
                'ticker_lec': _k(row_lec['Ticker']),
                'ticker_cer': _k(row_cer['Ticker']),
                'vencimiento': vto
            })

    # Diagnóstico: especies que no encontraron contraparte con el mismo vencimiento
    vtos_cer = set(df_cer['Vencimiento_dt'])
    vtos_lec = set(df_lec['Vencimiento_dt'])
    sin_par = []
    for _, r in df_lec.iterrows():
        if r['Vencimiento_dt'] not in vtos_cer:
            sin_par.append({'Hoja': 'Lecaps', 'Ticker': _k(r['Ticker']),
                            'Vencimiento': r['Vencimiento_dt'].strftime('%d/%m/%Y'),
                            'Motivo': 'No hay ninguna LECER con el mismo vencimiento'})
    for _, r in df_cer.iterrows():
        if r['Vencimiento_dt'] not in vtos_lec:
            sin_par.append({'Hoja': 'LECER', 'Ticker': _k(r['Ticker']),
                            'Vencimiento': r['Vencimiento_dt'].strftime('%d/%m/%Y'),
                            'Motivo': 'No hay ninguna Lecap con el mismo vencimiento'})

    # Columnas con precios que no figuran en ninguna hoja de especies (típico: falta cargar la ficha)
    conocidos = {_k(t) for t in df_lec['Ticker']} | {_k(t) for t in df_cer['Ticker']}
    if 'Letam' in xls.sheet_names:
        conocidos |= {_k(t) for t in pd.read_excel(xls, sheet_name='Letam')['Ticker']}
    for col in df_precios.columns:
        if _k(col) in ('FECHA', 'FECHA_DT') or _k(col) in conocidos:
            continue
        sin_par.append({'Hoja': 'Precios', 'Ticker': _k(col), 'Vencimiento': '-',
                        'Motivo': 'Tiene columna de precios pero no figura en Lecaps, LECER ni Letam '
                                  '(¿falta cargar la especie en su hoja?)'})

    return df_lec, df_cer, df_precios, df_indice_clean, df_tamar, feriados, pares, sin_par

@st.cache_data
def cargar_datos_tamar(firma):
    file_path = ruta_inputs()
    if file_path is None:
        return None

    xls = pd.ExcelFile(file_path)
    df_lec = pd.read_excel(xls, sheet_name='Lecaps')
    df_let = pd.read_excel(xls, sheet_name='Letam')
    df_precios = pd.read_excel(xls, sheet_name='Precios')
    df_tamar = pd.read_excel(xls, sheet_name='Tamar')
    df_feriados = pd.read_excel(xls, sheet_name='Feriados')

    feriados = set(pd.to_datetime(df_feriados["Feriados"]).dt.date)

    df_tamar['Fecha'] = pd.to_datetime(df_tamar['Fecha']).dt.date
    df_tamar = df_tamar[df_tamar['Tamar'] > 0].sort_values('Fecha').reset_index(drop=True)

    df_precios.columns = df_precios.columns.str.strip()
    df_precios = combinar_precios(df_precios)
    df_precios['Fecha_dt'] = pd.to_datetime(df_precios['Fecha']).dt.date
    df_precios = df_precios.sort_values('Fecha_dt').reset_index(drop=True)

    return df_lec, df_let, df_precios, df_tamar, feriados

# =========================================================
# MOTOR CER
# =========================================================
def id_periodo_cer(f):
    """(año, mes) del día 16 con que arranca el período CER que contiene a f."""
    if f.day >= 16:
        return (f.year, f.month)
    return (f.year - 1, 12) if f.month == 1 else (f.year, f.month - 1)

def fin_periodo_cer(f):
    """Día 15 con que termina el período CER que contiene a f."""
    if f.day <= 15:
        return date(f.year, f.month, 15)
    return date(f.year + 1, 1, 15) if f.month == 12 else date(f.year, f.month + 1, 15)

def calcular_meses_efectivos_bcra(f_inicio, f_fin):
    """N efectivo según Com. "A" 3507 del BCRA (períodos 16 → 15)."""
    if f_inicio >= f_fin:
        return 0.0

    n_efectivo = 0.0
    actual = f_inicio

    while actual < f_fin:
        if actual.day >= 16:
            inicio_periodo = date(actual.year, actual.month, 16)
            fin_periodo = date(actual.year + 1, 1, 15) if actual.month == 12 else date(actual.year, actual.month + 1, 15)
        else:
            fin_periodo = date(actual.year, actual.month, 15)
            inicio_periodo = date(actual.year - 1, 12, 16) if actual.month == 1 else date(actual.year, actual.month - 1, 16)

        dias_totales_periodo = (fin_periodo - inicio_periodo).days + 1
        fin_tramo = min(f_fin, fin_periodo)
        dias_en_tramo = (fin_tramo - actual).days if actual == f_inicio else (fin_tramo - actual).days + 1

        if dias_en_tramo > 0:
            n_efectivo += dias_en_tramo / dias_totales_periodo

        actual = fin_periodo + timedelta(days=1)

    return n_efectivo

def obtener_limite_cer_historico(f_op):
    """Hasta qué fecha el mercado conocía el CER en f_op (IPC se publica ~ día 14-15)."""
    if f_op.day < 15:
        return date(f_op.year, f_op.month, 15)
    return date(f_op.year + 1, 1, 15) if f_op.month == 12 else date(f_op.year, f_op.month + 1, 15)

def cer_conocido(df_indice, f_corte):
    """
    Último CER publicado <= f_corte. Si los datos terminan antes de f_corte pero dentro del
    mismo período CER, completa extrapolando con la tasa diaria constante observada en ese período.
    Devuelve (fecha, cer) o (None, None).
    """
    d = df_indice[df_indice['Fecha'] <= f_corte]
    if d.empty:
        return None, None
    f_u, c_u = d['Fecha'].iloc[-1], float(d['CER'].iloc[-1])

    if f_u < f_corte and id_periodo_cer(f_u) == id_periodo_cer(f_corte):
        y, m = id_periodo_cer(f_corte)
        base = d[(d['Fecha'] >= date(y, m, 16)) & (d['Fecha'] < f_u)]
        if not base.empty:
            f_b, c_b = base['Fecha'].iloc[0], float(base['CER'].iloc[0])
            dias = (f_u - f_b).days
            if dias > 0 and c_b > 0:
                g = (c_u / c_b) ** (1.0 / dias)
                c_u = c_u * g ** (f_corte - f_u).days
                f_u = f_corte
    return f_u, c_u

def calcular_be_cer(ticker_lec, ticker_cer, p_lec, p_cer, f_op, df_lec, df_cer, df_indice, feriados,
                    plazo_t=1, modo_historico=True):
    """
    modo_historico=True : el CER conocido en f_op se limita por la regla de publicación del IPC.
    modo_historico=False: simulación "hoy": se usa todo el CER cargado en el Excel.
    """
    if pd.isna(p_lec) or pd.isna(p_cer) or p_lec <= 0 or p_cer <= 0:
        return None

    row_lec = df_lec[df_lec['Ticker'].str.strip().str.upper() == ticker_lec].iloc[0]
    row_cer = df_cer[df_cer['Ticker'].str.strip().str.upper() == ticker_cer].iloc[0]

    f_liq = calcular_fecha_liq(f_op, plazo_t, feriados)

    # --- LECAP ---
    f_vto_lec = pd.to_datetime(row_lec['Vencimiento']).date()
    f_em_lec = pd.to_datetime(row_lec['Emisión']).date()
    tem_lec = row_lec['Tasa']
    d360_lec = dias360_excel(f_em_lec, f_vto_lec)

    if ticker_lec in VPV_OVERRIDE:
        vpv_lec = VPV_OVERRIDE[ticker_lec]
    elif 'VPV' in row_lec.index and pd.notnull(row_lec['VPV']):
        vpv_lec = float(row_lec['VPV'])
    else:
        vpv_lec = 100.0 * ((1 + tem_lec) ** (d360_lec / 30.0))

    f_cobro_lec = proximo_habil(f_vto_lec, feriados)
    dias_cartera_lec = (f_cobro_lec - f_liq).days
    if dias_cartera_lec <= 0:
        return None

    rend_lec = (vpv_lec / p_lec) - 1.0
    tir_lec = ((1.0 + rend_lec) ** (365.0 / dias_cartera_lec)) - 1.0

    # --- LECER ---
    f_vto_cer = pd.to_datetime(row_cer['Vencimiento']).date()
    f_cobro_cer = proximo_habil(f_vto_cer, feriados)
    dias_cartera_cer = (f_cobro_cer - f_liq).days
    if dias_cartera_cer <= 0:
        return None

    pesos_requeridos_cer = p_cer * ((1.0 + tir_lec) ** (dias_cartera_cer / 365.0))

    cer_inicial_emision = float(row_cer['CER inicial'])
    cupon_cer = row_cer['Tasa']
    vpv_real = 100.0 * (1.0 + cupon_cer)
    cer_vto_req = (pesos_requeridos_cer / vpv_real) * cer_inicial_emision

    # CER de referencia al vencimiento (rezago 10 hábiles)
    f_ref_vto = restar_dias_habiles(f_cobro_cer, 10, feriados)

    # CER de referencia hoy y tasa real (con CER constante)
    f_ref_hoy = restar_dias_habiles(f_liq, 10, feriados)
    df_hist_hoy = df_indice[df_indice['Fecha'] <= f_ref_hoy]
    cer_hoy_ref = float(df_hist_hoy['CER'].iloc[-1]) if not df_hist_hoy.empty else cer_inicial_emision
    factor_real = (vpv_real * cer_hoy_ref) / (p_cer * cer_inicial_emision)
    tir_real_cer = (factor_real ** (365.0 / dias_cartera_cer)) - 1.0 if factor_real > 0 else np.nan

    # ¿Cuánto del CER ya se conoce?
    if modo_historico:
        f_limite = obtener_limite_cer_historico(f_op)
    else:
        f_limite = fin_periodo_cer(df_indice['Fecha'].max())

    f_corte = min(f_ref_vto, f_limite)
    f_cer_ult, cer_ult = cer_conocido(df_indice, f_corte)
    if f_cer_ult is None or cer_ult <= 0:
        return None

    cer_es_conocido = f_ref_vto <= f_limite
    tem_be = tea_be = n_meses = infla_desconocida = tir_cer_fija = np.nan

    if cer_es_conocido:
        if f_cer_ult < f_ref_vto:
            return None  # faltan datos de CER para completar el tramo conocido
        pesos_vto_cer = vpv_real * cer_ult / cer_inicial_emision
        tir_cer_fija = (pesos_vto_cer / p_cer) ** (365.0 / dias_cartera_cer) - 1.0
    else:
        n_meses = calcular_meses_efectivos_bcra(f_cer_ult, f_ref_vto)
        if n_meses <= 0 or cer_vto_req <= 0:
            return None
        infla_desconocida = (cer_vto_req / cer_ult) - 1.0
        tem_be = ((cer_vto_req / cer_ult) ** (1.0 / n_meses)) - 1.0
        tea_be = ((1.0 + tem_be) ** 12.0) - 1.0

    return {
        'fecha_op': f_op, 'vpv_lec': vpv_lec, 'tir_lec': tir_lec,
        'pesos_req_cer': pesos_requeridos_cer, 'cer_hoy_ref': cer_hoy_ref,
        'cer_vto_req': cer_vto_req, 'f_ref_hoy': f_ref_hoy, 'f_ref_vto': f_ref_vto,
        'infla_total': (cer_vto_req / cer_hoy_ref) - 1.0,
        'infla_desconocida': infla_desconocida,
        'f_cer_ult': f_cer_ult, 'cer_ult': cer_ult, 'n_meses': n_meses,
        'tem_be': tem_be, 'tea_be': tea_be, 'tir_real_cer': tir_real_cer,
        'cer_conocido': cer_es_conocido, 'tir_cer_fija': tir_cer_fija,
    }

# =========================================================
# MOTOR TAMAR
# =========================================================
def tem_desde_tasa_comb(tasa_comb):
    """TAMAR TEM según prospecto, con tasa_comb = TAMAR + MARGEN (TNA decimal)."""
    return ((1 + tasa_comb / (365 / 32)) ** (365 / 32)) ** (1 / 12) - 1

def calcular_tamar_be(ticker_lec, ticker_let, p_lec, p_let, f_op,
                      df_lec, df_let, df_tamar, feriados, plazo_t=1):
    """Devuelve dict de resultados, o None si el cálculo no es válido para esa fecha."""
    if pd.isna(p_lec) or pd.isna(p_let) or p_lec <= 0 or p_let <= 0:
        return None

    row_lec = df_lec[df_lec['Ticker'] == ticker_lec].iloc[0]
    row_let = df_let[df_let['Ticker'] == ticker_let].iloc[0]
    f_liq = calcular_fecha_liq(f_op, plazo_t, feriados)

    # --- LECAP ---
    f_vto_lec = pd.to_datetime(row_lec['Vencimiento']).date()
    f_em_lec = pd.to_datetime(row_lec['Emisión']).date()
    tem_lec = row_lec['Tasa']
    d360_lec = dias360_excel(f_em_lec, f_vto_lec)

    if ticker_lec in VPV_OVERRIDE:
        vpv_lec = VPV_OVERRIDE[ticker_lec]
    elif 'VPV' in row_lec.index and pd.notnull(row_lec['VPV']):
        vpv_lec = float(row_lec['VPV'])
    else:
        vpv_lec = 100.0 * ((1 + tem_lec) ** (d360_lec / 30.0))

    f_cobro_lec = proximo_habil(f_vto_lec, feriados)
    dias_cartera_lec = (f_cobro_lec - f_liq).days
    if dias_cartera_lec <= 0:
        return None

    rend_lec = (vpv_lec / p_lec) - 1
    tir_lec = ((1 + rend_lec) ** (365 / dias_cartera_lec)) - 1
    tna_lec = rend_lec * (365 / dias_cartera_lec)

    # --- LETAM ---
    f_vto_let = pd.to_datetime(row_let['Vencimiento']).date()
    f_em_let = pd.to_datetime(row_let['Emisión']).date()
    margen_let = row_let['Tasa']
    d360_let = dias360_excel(f_em_let, f_vto_let)
    t_meses_let = (d360_let / 360) * 12
    f_cobro_let = proximo_habil(f_vto_let, feriados)

    dias_cartera_let = (f_cobro_let - f_liq).days
    if dias_cartera_let <= 0:
        return None

    vpv_objetivo = p_let * ((1 + tir_lec) ** (dias_cartera_let / 365))

    tem_cupon_target = ((vpv_objetivo / 100.0) ** (1.0 / t_meses_let)) - 1
    tea_target = (1 + tem_cupon_target) ** 12
    tasa_comb_target = ((tea_target ** (32 / 365)) - 1) * (365 / 32)  # = TAMAR_prom + MARGEN

    # --- FIXING TAMAR: ventana [emisión-10h, vto-10h], ambos inclusive ---
    f_inicio_fixing = restar_dias_habiles(f_em_let, 10, feriados)
    f_fin_fixing = restar_dias_habiles(f_vto_let, 10, feriados)
    n_total = contar_dias_habiles(f_inicio_fixing - timedelta(days=1), f_fin_fixing, feriados)

    f_corte_dev = min(f_op, f_fin_fixing)
    if f_corte_dev >= f_inicio_fixing:
        n_dev = contar_dias_habiles(f_inicio_fixing - timedelta(days=1), f_corte_dev, feriados)
        df_dev = df_tamar[(df_tamar['Fecha'] >= f_inicio_fixing) & (df_tamar['Fecha'] <= f_corte_dev)]
        dev_tna = df_dev['Tamar'].mean() / 100.0 if not df_dev.empty else np.nan
    else:
        n_dev, dev_tna = 0, 0.0

    n_proy = n_total - n_dev

    if n_proy > 0:
        tamar_be = ((tasa_comb_target - margen_let) * n_total - n_dev * dev_tna) / n_proy
    else:
        tamar_be = np.nan

    if n_proy > 0 and np.isnan(tamar_be):
        return None

    df_hist_mercado = df_tamar[df_tamar['Fecha'] <= f_op]
    if df_hist_mercado.empty:
        return None
    tamar_mercado = df_hist_mercado['Tamar'].iloc[-1] / 100.0

    spread = (tamar_mercado - tamar_be) * 100 if n_proy > 0 else 0.0

    tir_let = np.nan
    if n_proy == 0 and not np.isnan(dev_tna):
        tem_fija = tem_desde_tasa_comb(margen_let + dev_tna)
        vpv_let = 100.0 * (1 + tem_fija) ** t_meses_let
        rend_let = vpv_let / p_let - 1
        tir_let = (1 + rend_let) ** (365 / dias_cartera_let) - 1

    return {
        'fecha': f_op, 'tir_lec': tir_lec, 'tna_lec': tna_lec, 'vpv_objetivo': vpv_objetivo,
        'tamar_be': tamar_be, 'tamar_mercado': tamar_mercado, 'spread_pkt': spread,
        'es_fija': n_proy == 0, 'tir_let': tir_let,
        'dias_cartera_let': dias_cartera_let, 't_meses_let': t_meses_let, 'margen_let': margen_let,
        'f_inicio_fixing': f_inicio_fixing, 'f_fin_fixing': f_fin_fixing,
        'n_total': n_total, 'n_dev': n_dev, 'n_proy': n_proy,
    }

# =========================================================
# ESTILO COMPARTIDO (tarjetas KPI + gráficos)
# =========================================================
# Convención de color fija en toda la app (el color sigue a la serie, no a su posición):
#   azul    = Break-Even (la serie protagonista)
#   naranja = serie de comparación (TAMAR de mercado / tasa real CER / TAMAR)
COLORES = {
    'be':   {'claro': '#2a78d6', 'oscuro': '#3987e5'},
    'comp': {'claro': '#eb6834', 'oscuro': '#d95926'},
}
SUPERFICIE = {'claro': '#ffffff', 'oscuro': '#0e1117'}

CSS = """
<style>
.block-container {max-width: 1400px; padding-top: 2.2rem;}
.kpi, .kpi * {font-family: system-ui, -apple-system, "Segoe UI", sans-serif;}
.kpi {border: 1px solid rgba(128,128,128,.28); border-radius: 12px; padding: 12px 16px;}
.kpi .lbl {font-size: .78rem; opacity: .72;}
.kpi .val {font-size: 1.35rem; font-weight: 600; line-height: 1.25; margin-top: 2px;}
.kpi .sub {font-size: .78rem; opacity: .65; margin-top: 2px;}
.kpi.hero {padding: 18px 24px; border-left: 6px solid var(--kpi-accent, #2a78d6);}
.kpi.hero .lbl {font-size: .95rem;}
.kpi.hero .val {font-size: 3.6rem; font-weight: 700; line-height: 1.1; letter-spacing: -.01em;}
.kpi.hero .sub {font-size: .9rem; margin-top: 6px;}
</style>
"""

def _modo():
    """'oscuro' o 'claro' según el tema activo de Streamlit (claro si no se puede detectar)."""
    try:
        return 'oscuro' if st.context.theme.type == 'dark' else 'claro'
    except Exception:
        return 'claro'

def color(nombre):
    return COLORES[nombre][_modo()]

def pct(x, d=2, signo=False):
    """Fracción decimal -> texto en %. NaN-safe."""
    if x is None or pd.isna(x):
        return "—"
    return f"{x * 100:{'+' if signo else ''}.{d}f}%"

def kpi(label, valor, sub="", hero=False):
    """Tarjeta de indicador. hero=True: la cifra principal de la vista (una sola por vista)."""
    cls = "kpi hero" if hero else "kpi"
    accent = f' style="--kpi-accent:{color("be")}"' if hero else ""
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    st.markdown(f'<div class="{cls}"{accent}><div class="lbl">{label}</div>'
                f'<div class="val">{valor}</div>{sub_html}</div>', unsafe_allow_html=True)

def encabezado(titulo, bajada, fecha_sim, plazo_label, fuente, extra=""):
    st.title(titulo)
    st.caption(f"{bajada} · Al {fecha_sim:%d/%m/%Y} · {plazo_label} · Precios: {fuente}{extra}")

def nueva_fig(titulo, y_dec=1, leyenda=True, alto=420):
    """Figura con el estilo común: una sola escala Y, grilla recesiva, leyenda arriba a la derecha."""
    grilla = "rgba(128,128,128,.18)"
    fig = go.Figure()
    fig.update_layout(
        title=dict(text=f"<b>{titulo}</b>", x=0, xanchor="left", font=dict(size=16)),
        height=alto, hovermode="x unified", showlegend=leyenda,
        margin=dict(l=10, r=70, t=70, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(title=None, showgrid=False, showline=True, linecolor=grilla, tickformat="%d/%m/%y")
    fig.update_yaxes(title=None, ticksuffix="%", tickformat=f".{y_dec}f", gridcolor=grilla,
                     gridwidth=1, zeroline=False)
    return fig

def agregar_linea(fig, x, y, nombre, clave_color, etiqueta_final=False, eje_derecho=False):
    """Línea de 2px. etiqueta_final=True: punto + valor al final (solo para la serie protagonista).
    eje_derecho=True: la serie se grafica contra el eje Y secundario (derecha)."""
    d = pd.DataFrame({'x': list(x), 'y': list(y)}).dropna()
    if d.empty:
        return
    c = color(clave_color)
    fig.add_trace(go.Scatter(
        x=d['x'], y=d['y'], mode="lines", name=nombre, line=dict(color=c, width=2),
        yaxis="y2" if eje_derecho else "y",
        hovertemplate=f"{nombre}: %{{y:.2f}}%<extra></extra>"))
    if etiqueta_final:
        fig.add_trace(go.Scatter(
            x=[d['x'].iloc[-1]], y=[d['y'].iloc[-1]], mode="markers+text",
            text=[f"{d['y'].iloc[-1]:.2f}%"], textposition="middle right", cliponaxis=False,
            showlegend=False, hoverinfo="skip",
            marker=dict(size=9, color=c, line=dict(width=2, color=SUPERFICIE[_modo()]))))

# =========================================================
# VISTA: INFLACIÓN BREAK-EVEN (CER)
# =========================================================
def render_cer(D, precios_vivo, hoy, plazo_t, plazo_label):
    titulo, bajada = "Inflación implícita Break-Even", "Lecap (tasa fija) vs LECER (CER)"

    if D is None:
        st.title(titulo)
        st.error("No se encontró el archivo 'Inputs.xlsx' en el directorio.")
        return
    df_lec, df_cer, df_precios, df_indice, df_tamar, feriados, pares, sin_par = D
    if not pares:
        st.title(titulo)
        st.error("No se encontraron pares de Lecap y LECER con el mismo vencimiento.")
        return

    ultima_fila = df_precios.iloc[-1]
    fecha_sim = hoy if precios_vivo else ultima_fila['Fecha_dt']

    def precio_y_fuente(ticker):
        """(precio base 100, fuente). API si existe; si no, última fila del Excel."""
        v = precios_vivo.get(ticker) if precios_vivo else None
        if v is not None and not pd.isna(v):
            return normalizar_precio(v), "API"
        v = ultima_fila[ticker] if ticker in ultima_fila.index else np.nan
        return normalizar_precio(v), "Excel"

    pares_ordenados = sorted(pares, key=lambda x: x['vencimiento'])

    # ---- Barra lateral ----
    st.sidebar.subheader("Par a analizar")
    opciones_labels = [p['label'] for p in pares_ordenados]
    default_idx = next(
        (i for i, p in enumerate(pares_ordenados) if p['ticker_lec'] == 'S30N6' and p['ticker_cer'] == 'X30N6'), 0)
    par_label = st.sidebar.selectbox("Par (Lecap vs LECER)", opciones_labels, index=default_idx)
    par_info = next(p for p in pares_ordenados if p['label'] == par_label)
    lec_ticker, cer_ticker = par_info['ticker_lec'], par_info['ticker_cer']

    def_p_l = precio_y_fuente(lec_ticker)[0]
    def_p_c = precio_y_fuente(cer_ticker)[0]
    def_p_l = 100.0 if pd.isna(def_p_l) else float(def_p_l)
    def_p_c = 100.0 if pd.isna(def_p_c) else float(def_p_c)

    precio_lec_sim = st.sidebar.number_input(f"Precio {lec_ticker}", value=def_p_l, step=0.10, format="%.2f")
    precio_cer_sim = st.sidebar.number_input(f"Precio {cer_ticker}", value=def_p_c, step=0.10, format="%.2f")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Rango histórico")
    fecha_min_gen = df_precios['Fecha_dt'].min()
    fecha_max_gen = df_precios['Fecha_dt'].max()
    rango_fechas = st.sidebar.date_input("Período a graficar", value=(fecha_min_gen, fecha_max_gen),
                                         min_value=fecha_min_gen, max_value=fecha_max_gen)
    if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
        f_desde, f_hasta = rango_fechas
    else:
        f_desde, f_hasta = fecha_min_gen, fecha_max_gen

    # ---- Encabezado ----
    encabezado(titulo, bajada, fecha_sim, plazo_label, "API 1816 en vivo" if precios_vivo else FUENTE_HIST,
               extra=f" · Último CER cargado: {df_indice['Fecha'].max():%d/%m/%Y}")

    # ---- Resumen de mercado ----
    resumen_list, excluidos = [], []
    for p in pares_ordenados:
        t_lec, t_cer = p['ticker_lec'], p['ticker_cer']
        venc = p['vencimiento'].strftime('%d/%m/%Y')
        (px_l, f_l), (px_c, f_c) = precio_y_fuente(t_lec), precio_y_fuente(t_cer)
        if pd.isna(px_l) or pd.isna(px_c):
            faltan = [t for t, px in ((t_lec, px_l), (t_cer, px_c)) if pd.isna(px)]
            excluidos.append({'Tasa Fija': t_lec, 'CER': t_cer, 'Vencimiento': venc,
                              'Motivo': "Sin precio para " + ", ".join(faltan) +
                                        " (no está en la API ni como columna en la hoja Precios)"})
            continue
        try:
            r_all = calcular_be_cer(t_lec, t_cer, px_l, px_c, fecha_sim, df_lec, df_cer, df_indice, feriados,
                                    plazo_t=plazo_t, modo_historico=False)
        except Exception as e:
            excluidos.append({'Tasa Fija': t_lec, 'CER': t_cer, 'Vencimiento': venc,
                              'Motivo': f"Error de cálculo ({type(e).__name__}: {e}): revisá las columnas "
                                        "de las hojas Lecaps / LECER para estos tickers"})
            continue
        if not r_all:
            excluidos.append({'Tasa Fija': t_lec, 'CER': t_cer, 'Vencimiento': venc,
                              'Motivo': "No se pudo calcular: ya venció, precio inválido o faltan datos "
                                        "de CER hasta el vencimiento"})
            continue
        resumen_list.append({
            'Vencimiento': venc, 'Tasa Fija': t_lec, 'CER': t_cer,
            'TEM BE (%)': r_all['tem_be'] * 100,
            'TEA BE (%)': r_all['tea_be'] * 100,
            'N meses': r_all['n_meses'],
            'Fuente': f_l if f_l == f_c else 'Mixto',
        })
    df_resumen = pd.DataFrame(resumen_list)

    st.subheader("Resumen de mercado")
    st.caption("Break-even por vencimiento. Si el CER del vencimiento ya está publicado no hay inflación "
               "que implicar (se indica 'CER conocido').")
    if not df_resumen.empty:
        # La columna Fuente solo aporta información cuando hay precios en vivo
        df_mostrar = df_resumen if precios_vivo else df_resumen.drop(columns=['Fuente'])
        st.dataframe(
            df_mostrar.style.format({'TEM BE (%)': '{:.2f}%', 'TEA BE (%)': '{:.2f}%', 'N meses': '{:.2f}'},
                                    na_rep="CER conocido"),
            hide_index=True)
    else:
        st.warning("No se pudo calcular ningún par (¿faltan precios o ya vencieron?).")
    if precios_vivo and not df_resumen.empty:
        n_no_api = int((df_resumen['Fuente'] != 'API').sum())
        if n_no_api:
            st.warning(f"{n_no_api} de {len(df_resumen)} pares usan (total o parcialmente) precios del Excel, "
                       "no de la API: ese precio puede ser de una fecha anterior a hoy. Ver columna 'Fuente'.")
    st.markdown("---")

    # ---- Detalle del par ----
    res_sim = calcular_be_cer(lec_ticker, cer_ticker, precio_lec_sim, precio_cer_sim, fecha_sim,
                              df_lec, df_cer, df_indice, feriados,
                              plazo_t=plazo_t, modo_historico=False)

    st.subheader(f"Detalle: {lec_ticker} vs {cer_ticker}")
    st.caption(f"Vencimiento {par_info['vencimiento']:%d/%m/%Y}")

    if res_sim is None:
        st.warning("No se pudo calcular el par: revisá precios, que no haya vencido, y que el CER cargado "
                   "cubra hasta el último dato conocido.")
    elif res_sim['cer_conocido']:
        c = st.columns([2, 1, 1])
        with c[0]:
            kpi(f"TIR {cer_ticker} (CER ya conocido)", pct(res_sim['tir_cer_fija']),
                sub=f"El CER de referencia al vencimiento ({res_sim['f_ref_vto']:%d/%m/%Y}) ya está publicado: "
                    "la LECER opera como tasa fija y se comparan las TIR", hero=True)
        with c[1]:
            kpi(f"TEA fija {lec_ticker}", pct(res_sim['tir_lec']))
        with c[2]:
            kpi("TIR CER − TIR fija", f"{(res_sim['tir_cer_fija'] - res_sim['tir_lec']) * 100:+.2f} pp")
    else:
        c = st.columns([2, 1, 1, 1])
        with c[0]:
            kpi("Inflación mensual implícita (TEM BE)", pct(res_sim['tem_be']),
                sub=f"TEA BE {pct(res_sim['tea_be'])} · tramo desconocido de {res_sim['n_meses']:.2f} meses",
                hero=True)
        with c[1]:
            kpi(f"TEA fija {lec_ticker}", pct(res_sim['tir_lec']))
        with c[2]:
            kpi(f"Tasa real CER {cer_ticker}", pct(res_sim['tir_real_cer']))
        with c[3]:
            kpi(f"Pesos target {cer_ticker}", f"${res_sim['pesos_req_cer']:.2f}")
        st.caption(f"Tramo desconocido: desde {res_sim['f_cer_ult']:%d/%m/%Y} hasta "
                   f"{res_sim['f_ref_vto']:%d/%m/%Y}. Inflación acumulada implícita en ese tramo: "
                   f"{res_sim['infla_desconocida'] * 100:.2f}%.")
        if res_sim['n_meses'] < 1.0:
            st.warning("El tramo desconocido es menor a 1 mes: la TEM BE se calcula elevando a 1/N con N chico, "
                       "por lo que pequeñas diferencias de precio la mueven mucho. Mirá la inflación acumulada "
                       "implícita del tramo como referencia más estable.")

    # ---- Serie histórica ----
    df_precios_filt = df_precios[(df_precios['Fecha_dt'] >= f_desde) & (df_precios['Fecha_dt'] <= f_hasta)]
    historico = []
    for _, row in df_precios_filt.iterrows():
        f_o = row['Fecha_dt']
        p_l = normalizar_precio(row[lec_ticker]) if lec_ticker in row.index else np.nan
        p_c = normalizar_precio(row[cer_ticker]) if cer_ticker in row.index else np.nan
        if pd.isna(p_l) or pd.isna(p_c):
            continue
        r = calcular_be_cer(lec_ticker, cer_ticker, p_l, p_c, f_o, df_lec, df_cer, df_indice, feriados,
                            plazo_t=plazo_t, modo_historico=True)
        if r:
            if not df_tamar.empty:
                tamar_sub = df_tamar[df_tamar['Fecha_dt'] <= f_o]
                r['tamar'] = tamar_sub['Tamar_pct'].iloc[-1] if not tamar_sub.empty else np.nan
            else:
                r['tamar'] = np.nan
            historico.append(r)
    df_hist = pd.DataFrame(historico)

    st.markdown("---")
    st.subheader("Evolución histórica")

    if not df_hist.empty and df_hist['tem_be'].notna().any():
        df_hist['tem_be_pct'] = df_hist['tem_be'] * 100
        df_hist['tea_be_pct'] = df_hist['tea_be'] * 100
        df_hist['tir_real_cer_pct'] = df_hist['tir_real_cer'] * 100
        st.caption("Los días en que el CER del vencimiento ya era conocido no se grafican (no hay break-even).")

        tab1, tab2, tab3 = st.tabs(["Inflación mensual (TEM)", "TEA BE vs tasa real CER", "TEA BE vs TAMAR"])

        with tab1:
            fig1 = nueva_fig(f"Inflación mensual implícita (TEM BE): {lec_ticker} vs {cer_ticker}",
                             y_dec=2, leyenda=False)
            agregar_linea(fig1, df_hist['fecha_op'], df_hist['tem_be_pct'], "TEM BE", 'be', etiqueta_final=True)
            st.plotly_chart(fig1)

        with tab2:
            fig2 = nueva_fig(f"Inflación anualizada implícita (TEA BE) vs tasa real CER ({cer_ticker})")
            # Dos escalas distintas: TEA BE a la izquierda, tasa real CER a la derecha.
            # Cada eje lleva el color de su serie (línea del eje) para saber cuál es cuál.
            agregar_linea(fig2, df_hist['fecha_op'], df_hist['tea_be_pct'], "TEA BE", 'be')
            agregar_linea(fig2, df_hist['fecha_op'], df_hist['tir_real_cer_pct'],
                          f"Tasa real CER ({cer_ticker})", 'comp', eje_derecho=True)
            fig2.update_layout(
                margin=dict(l=10, r=10, t=70, b=10),
                yaxis=dict(title=dict(text="TEA BE", font=dict(size=12)), automargin=True,
                           showline=True, linewidth=2, linecolor=color('be')),
                yaxis2=dict(title=dict(text="Tasa real CER", font=dict(size=12)), automargin=True,
                            overlaying="y", side="right", showgrid=False, zeroline=False,
                            ticksuffix="%", tickformat=".1f",
                            showline=True, linewidth=2, linecolor=color('comp')))
            st.plotly_chart(fig2)

        with tab3:
            if df_hist['tamar'].notna().any():
                fig3 = nueva_fig("Inflación anualizada implícita (TEA BE) vs TAMAR")
                agregar_linea(fig3, df_hist['fecha_op'], df_hist['tea_be_pct'], "TEA BE", 'be', etiqueta_final=True)
                agregar_linea(fig3, df_hist['fecha_op'], df_hist['tamar'], "TAMAR", 'comp')
                st.plotly_chart(fig3)
            else:
                st.info("No hay datos de TAMAR en el Excel para este período.")
    else:
        st.warning("No hay datos con inflación break-even para el rango seleccionado "
                   "(o el CER de todo el período ya era conocido al vencimiento).")

# =========================================================
# VISTA: TAMAR BREAK-EVEN
# =========================================================
def render_tamar(D, precios_vivo, hoy, plazo_t, plazo_label):
    titulo, bajada = "TAMAR Break-Even", "Lecap (tasa fija) vs Letam (tasa variable)"

    if D is None:
        st.title(titulo)
        st.error("No se encontró el archivo 'Inputs.xlsx' en el directorio.")
        return
    df_lec, df_let, df_precios, df_tamar, feriados = D

    # ---- Barra lateral ----
    st.sidebar.subheader("Especies")
    lecap_elegida = st.sidebar.selectbox("Tasa fija (Lecap/Boncap)", df_lec['Ticker'].tolist())
    letam_elegida = st.sidebar.selectbox("Tasa variable (Letam)", df_let['Ticker'].tolist())

    px_l_api = precios_vivo.get(_k(lecap_elegida))
    px_t_api = precios_vivo.get(_k(letam_elegida))
    usa_vivo = px_l_api is not None and px_t_api is not None

    if usa_vivo:
        def_p_l = float(normalizar_precio(px_l_api))
        def_p_t = float(normalizar_precio(px_t_api))
        fecha_sim = hoy
    else:
        if precios_vivo:
            st.sidebar.warning("La API no devolvió precio para este par. Usando Excel.")
        df_validos = df_precios.dropna(subset=[lecap_elegida, letam_elegida])
        if df_validos.empty:
            st.title(titulo)
            st.error("No hay fechas con precio para ambas especies seleccionadas.")
            return
        ultima_fila = df_validos.iloc[-1]
        fecha_sim = ultima_fila['Fecha_dt']
        def_p_l = float(normalizar_precio(ultima_fila[lecap_elegida]))
        def_p_t = float(normalizar_precio(ultima_fila[letam_elegida]))

    st.sidebar.markdown("---")
    st.sidebar.subheader("Precios de simulación")
    precio_lecap_sim = st.sidebar.number_input(f"Precio {lecap_elegida}", value=def_p_l, step=0.10, format="%.2f")
    precio_letam_sim = st.sidebar.number_input(f"Precio {letam_elegida}", value=def_p_t, step=0.10, format="%.2f")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Rango histórico")
    fecha_min_general = df_precios['Fecha_dt'].min()
    fecha_max_general = df_precios['Fecha_dt'].max()
    rango_fechas = st.sidebar.date_input("Período a graficar", value=(fecha_min_general, fecha_max_general),
                                         min_value=fecha_min_general, max_value=fecha_max_general)
    if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
        f_desde, f_hasta = rango_fechas
    else:
        f_desde, f_hasta = fecha_min_general, fecha_max_general

    # ---- Encabezado ----
    encabezado(titulo, bajada, fecha_sim, plazo_label, "API 1816 en vivo" if usa_vivo else FUENTE_HIST)

    res_sim = calcular_tamar_be(lecap_elegida, letam_elegida, precio_lecap_sim, precio_letam_sim, fecha_sim,
                                df_lec, df_let, df_tamar, feriados, plazo_t=plazo_t)

    if res_sim is None:
        st.error("No se pudo calcular la simulación: revisá que las especies no hayan vencido a la fecha "
                 "y que haya datos de TAMAR para la ventana de fixing.")
        return

    # ---- Indicadores ----
    if res_sim['es_fija']:
        c = st.columns([2, 1, 1, 1])
        with c[0]:
            kpi(f"TIR {letam_elegida} (ya conocida)", pct(res_sim['tir_let']),
                sub="El fixing finalizó: la Letam ya devengó toda la TAMAR y opera como tasa fija, "
                    "por lo que no corresponde calcular un break-even futuro", hero=True)
        with c[1]:
            kpi(f"TIR {lecap_elegida}", pct(res_sim['tir_lec']))
        with c[2]:
            kpi(f"TNA {lecap_elegida}", pct(res_sim['tna_lec']))
        with c[3]:
            kpi(f"VPV requerido {letam_elegida}", f"${res_sim['vpv_objetivo']:.2f}")
    else:
        c = st.columns([2, 1, 1, 1, 1])
        with c[0]:
            kpi("TAMAR Break-Even", pct(res_sim['tamar_be']),
                sub="TNA constante sobre los días a proyectar que iguala la TIR de la Lecap", hero=True)
        with c[1]:
            kpi("TAMAR mercado hoy", pct(res_sim['tamar_mercado']))
        with c[2]:
            kpi("Spread BE vs mercado", f"{res_sim['spread_pkt']:+.2f} pkt")
        with c[3]:
            kpi(f"TNA {lecap_elegida}", pct(res_sim['tna_lec']))
        with c[4]:
            kpi(f"VPV requerido {letam_elegida}", f"${res_sim['vpv_objetivo']:.2f}")
        st.caption(
            f"Ventana de fixing: {res_sim['f_inicio_fixing']:%d/%m/%Y} a {res_sim['f_fin_fixing']:%d/%m/%Y} "
            f"({res_sim['n_total']} hábiles; {res_sim['n_dev']} conocidos, {res_sim['n_proy']} a proyectar). "
            "Supone reinversión a la TIR de la Lecap si ésta vence antes que la Letam.")

    # ---- Serie histórica (siempre con precios del Excel) ----
    df_precios_filtrado = df_precios[(df_precios['Fecha_dt'] >= f_desde) & (df_precios['Fecha_dt'] <= f_hasta)]
    historico = []
    for _, row in df_precios_filtrado.iterrows():
        p_l = normalizar_precio(row[lecap_elegida])
        p_t = normalizar_precio(row[letam_elegida])
        if pd.isna(p_l) or pd.isna(p_t):
            continue
        r = calcular_tamar_be(lecap_elegida, letam_elegida, p_l, p_t, row['Fecha_dt'],
                              df_lec, df_let, df_tamar, feriados, plazo_t=plazo_t)
        if r is not None:
            historico.append(r)
    df_hist = pd.DataFrame(historico)

    st.markdown("---")
    st.subheader("Evolución histórica")

    if not df_hist.empty:
        # Se excluyen los días con fixing cerrado para que la línea no caiga a 0% y arruine la escala
        df_hist_plot = df_hist[~df_hist['es_fija']].copy()
        if not df_hist_plot.empty:
            df_hist_plot['be_pct'] = df_hist_plot['tamar_be'] * 100
            df_hist_plot['mkt_pct'] = df_hist_plot['tamar_mercado'] * 100

            st.caption("Los días en que el fixing ya había cerrado no se grafican (no hay break-even).")
            fig = nueva_fig(f"TAMAR Break-Even vs TAMAR de mercado: {lecap_elegida} vs {letam_elegida}")
            agregar_linea(fig, df_hist_plot['fecha'], df_hist_plot['be_pct'], "TAMAR Break-Even", 'be',
                          etiqueta_final=True)
            agregar_linea(fig, df_hist_plot['fecha'], df_hist_plot['mkt_pct'], "TAMAR de mercado", 'comp')
            st.plotly_chart(fig)
        else:
            st.warning("Todos los datos en este rango corresponden al período donde la Letam ya operaba "
                       "como tasa fija (cierre de fixing).")
    else:
        st.warning("No hay datos válidos en el rango de fechas seleccionado.")

# =========================================================
# MAIN: NAVEGACIÓN + API COMPARTIDA
# =========================================================
st.markdown(CSS, unsafe_allow_html=True)

_firma = firma_archivos()
D_cer = cargar_datos_cer(_firma)
D_tamar = cargar_datos_tamar(_firma)

# Feriados para saber si el mercado está operando (se toman de cualquiera de las dos cargas)
feriados_g = D_cer[5] if D_cer is not None else (D_tamar[4] if D_tamar is not None else set())
ahora = datetime.now(TZ_AR)
hoy = ahora.date()
mercado_abierto = es_habil(hoy, feriados_g) and 11 <= ahora.hour < 17  # horario aproximado BYMA

st.sidebar.title("Renta fija")
vista = st.sidebar.radio("Calculadora", ["Inflación BE (CER)", "TAMAR BE"], key="vista")
st.sidebar.markdown("---")

st.sidebar.subheader("Precios y liquidación")
usar_api = st.sidebar.checkbox("Usar precios en vivo (API 1816)", value=False)
if usar_api and st.sidebar.button("Actualizar precios"):
    obtener_precios_1816.clear()  # fuerza una nueva consulta a la API
precios_vivo = {}
if usar_api:
    api_key = _get_api_key()
    if not api_key:
        st.sidebar.error("Falta API_1816_KEY (en st.secrets o variable de entorno). Usando Excel.")
    else:
        todos = set()
        if D_cer is not None:
            todos |= {_k(t) for t in D_cer[0]['Ticker']} | {_k(t) for t in D_cer[1]['Ticker']}
        if D_tamar is not None:
            todos |= {_k(t) for t in D_tamar[0]['Ticker']} | {_k(t) for t in D_tamar[1]['Ticker']}
        try:
            with st.sidebar.status("Conectando con 1816..."):
                precios_vivo, ts_api = obtener_precios_1816(tuple(sorted(todos)), api_key)
            st.sidebar.success(f"1816 conectado: {len(precios_vivo)} especies, consulta de las {ts_api:%H:%M:%S}")
            if not mercado_abierto:
                st.sidebar.warning("Mercado cerrado (fuera de horario, fin de semana o feriado): los precios "
                                   "pueden ser del último cierre.")
        except Exception as e:
            st.sidebar.error(f"Falló la API: {e}. Usando Excel.")
            precios_vivo = {}

# Plazo de liquidación global (aplica a las dos calculadoras)
plazo_opciones = {"T+0 (contado inmediato)": 0, "T+1 (24hs)": 1}
plazo_label = st.sidebar.selectbox("Plazo de liquidación", list(plazo_opciones.keys()), index=1)
plazo_t = plazo_opciones[plazo_label]
st.sidebar.markdown("---")

if vista == "Inflación BE (CER)":
    render_cer(D_cer, precios_vivo, hoy, plazo_t, plazo_label)
else:
    render_tamar(D_tamar, precios_vivo, hoy, plazo_t, plazo_label)