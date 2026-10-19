import os
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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

@st.cache_data(ttl=300)
def obtener_precios_1816(lista_tickers, api_key):
    """Lanza excepción si algo falla (así no se cachea el error)."""
    auth = requests.post("https://api.1816.com.ar/v1/auth/token",
                         json={"apiKey": api_key, "module": "mercado"}, timeout=10)
    auth.raise_for_status()
    token = auth.json()["token"]

    url = "https://api.1816.com.ar/v1/mercado/indicadores"
    params = {"tickers": [t.strip().lower() for t in lista_tickers], "campos": ["precioDirty"]}
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=10)
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

@st.cache_data
def cargar_datos_cer():
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

    return df_lec, df_cer, df_precios, df_indice_clean, df_tamar, feriados, pares

@st.cache_data
def cargar_datos_tamar():
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

def test_sintetico(ticker_lec, ticker_let, p_lec, x_tamar, df_lec, df_let, feriados):
    """TAMAR constante = x_tamar: el BE debe devolver x_tamar al precio de equilibrio."""
    row_let = df_let[df_let['Ticker'] == ticker_let].iloc[0]
    f_em = pd.to_datetime(row_let['Emisión']).date()
    f_vto = pd.to_datetime(row_let['Vencimiento']).date()
    ini = restar_dias_habiles(f_em, 10, feriados)
    fin = restar_dias_habiles(f_vto, 10, feriados)

    dias = []
    d = ini - timedelta(days=30)
    while d <= fin + timedelta(days=30):
        if es_habil(d, feriados):
            dias.append(d)
        d += timedelta(days=1)
    df_syn = pd.DataFrame({'Fecha': dias, 'Tamar': x_tamar * 100.0})

    f_op = proximo_habil(ini + (fin - ini) / 2, feriados)

    r0 = calcular_tamar_be(ticker_lec, ticker_let, p_lec, 100.0, f_op, df_lec, df_let, df_syn, feriados)
    if r0 is None:
        return None, "No se pudo calcular el caso base (¿fechas/datos fuera de rango?)."

    tem = tem_desde_tasa_comb(r0['margen_let'] + x_tamar)
    vpv_x = 100.0 * (1 + tem) ** r0['t_meses_let']
    p_eq = vpv_x / ((1 + r0['tir_lec']) ** (r0['dias_cartera_let'] / 365))

    r1 = calcular_tamar_be(ticker_lec, ticker_let, p_lec, p_eq, f_op, df_lec, df_let, df_syn, feriados)
    if r1 is None:
        return None, "No se pudo calcular el caso de equilibrio."
    return r1, f_op

# =========================================================
# VISTA: INFLACIÓN BREAK-EVEN (CER)
# =========================================================
def render_cer(D, precios_vivo, hoy, plazo_t_selected, plazo_label):
    st.title("Dashboard: Inflación Implícita Break-Even (Tasa Fija vs CER)")

    if D is None:
        st.error("Error: No se encontró el archivo 'Inputs.xlsx' en el directorio.")
        return
    df_lec, df_cer, df_precios, df_indice, df_tamar, feriados, pares = D
    if not pares:
        st.error("Error: No se encontraron pares de Lecap y LECER con la misma fecha de vencimiento.")
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

    def precio_actual(ticker):
        return precio_y_fuente(ticker)[0]

    st.sidebar.write(f"**Fecha del último dato:** {fecha_sim.strftime('%d/%m/%Y')}")
    st.sidebar.write(f"**Último CER cargado:** {df_indice['Fecha'].max().strftime('%d/%m/%Y')}")

    # Control de unidades: API vs Excel
    if precios_vivo:
        tickers_cer = {p['ticker_lec'] for p in pares} | {p['ticker_cer'] for p in pares}
        with st.expander("🔎 Control de precios: API vs Excel (verificar unidades)"):
            filas = []
            for t in sorted(t for t in precios_vivo if t in tickers_cer):
                px_api = normalizar_precio(precios_vivo[t])
                px_xls = normalizar_precio(ultima_fila[t]) if t in ultima_fila.index else np.nan
                filas.append({'Ticker': t, 'API': px_api, 'Excel (última fila)': px_xls,
                              'Dif %': (px_api / px_xls - 1) * 100 if pd.notnull(px_xls) and px_xls else np.nan})
            st.dataframe(pd.DataFrame(filas).style.format(
                {'API': '{:.2f}', 'Excel (última fila)': '{:.2f}', 'Dif %': '{:+.2f}%'}, na_rep="-"),
                use_container_width=True, hide_index=True)

    # ---- Tabla de resumen ----
    pares_ordenados = sorted(pares, key=lambda x: x['vencimiento'])
    resumen_list, omitidos = [], []
    for p in pares_ordenados:
        t_lec, t_cer = p['ticker_lec'], p['ticker_cer']
        (px_l, f_l), (px_c, f_c) = precio_y_fuente(t_lec), precio_y_fuente(t_cer)
        if pd.isna(px_l) or pd.isna(px_c):
            omitidos.append(f"{t_lec}/{t_cer}")
            continue
        r_all = calcular_be_cer(t_lec, t_cer, px_l, px_c, fecha_sim, df_lec, df_cer, df_indice, feriados,
                                plazo_t=plazo_t_selected, modo_historico=False)
        if r_all:
            resumen_list.append({
                'Vencimiento': p['vencimiento'].strftime('%d/%m/%Y'),
                'Tasa Fija': t_lec, 'CER': t_cer,
                'TEM BE (%)': r_all['tem_be'] * 100,
                'TEA BE (%)': r_all['tea_be'] * 100,
                'N meses': r_all['n_meses'],
                'Fuente': f_l if f_l == f_c else 'Mixto',
            })
    df_resumen = pd.DataFrame(resumen_list)

    st.subheader("Resumen de Mercado: Break-Even por Vencimiento")
    st.caption(f"Calculado al {fecha_sim.strftime('%d/%m/%Y')} en **{plazo_label}**. "
               "Si el CER del vencimiento ya está publicado no hay inflación que implicar (se indica 'CER conocido').")
    if not df_resumen.empty:
        st.dataframe(
            df_resumen.style.format({'TEM BE (%)': '{:.2f}%', 'TEA BE (%)': '{:.2f}%', 'N meses': '{:.2f}'},
                                    na_rep="CER conocido"),
            use_container_width=True, hide_index=True)
    else:
        st.warning("No se pudo calcular ningún par (¿faltan precios o ya vencieron?).")
    if precios_vivo and not df_resumen.empty:
        n_no_api = int((df_resumen['Fuente'] != 'API').sum())
        if n_no_api:
            st.warning(f"{n_no_api} de {len(df_resumen)} pares usan (total o parcialmente) precios del Excel, "
                       "no de la API: ese precio puede ser de una fecha anterior a hoy. Ver columna 'Fuente'.")
    if omitidos:
        st.caption("Pares omitidos por falta de precio: " + ", ".join(omitidos))

    st.markdown("---")

    # ---- Detalle de especie ----
    opciones_labels = [p['label'] for p in pares_ordenados]
    default_idx = next(
        (i for i, p in enumerate(pares_ordenados) if p['ticker_lec'] == 'S30N6' and p['ticker_cer'] == 'X30N6'), 0)
    par_label = st.sidebar.selectbox("1. Seleccione el par a comparar:", opciones_labels, index=default_idx)
    par_info = next(p for p in pares_ordenados if p['label'] == par_label)
    lec_ticker, cer_ticker = par_info['ticker_lec'], par_info['ticker_cer']

    def_p_l = precio_actual(lec_ticker)
    def_p_c = precio_actual(cer_ticker)
    def_p_l = 100.0 if pd.isna(def_p_l) else float(def_p_l)
    def_p_c = 100.0 if pd.isna(def_p_c) else float(def_p_c)

    precio_lec_sim = st.sidebar.number_input(f"Precio {lec_ticker} (Tasa Fija):", value=def_p_l, step=0.10, format="%.2f")
    precio_cer_sim = st.sidebar.number_input(f"Precio {cer_ticker} (CER):", value=def_p_c, step=0.10, format="%.2f")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Rango Histórico a Graficar")
    fecha_min_gen = df_precios['Fecha_dt'].min()
    fecha_max_gen = df_precios['Fecha_dt'].max()
    rango_fechas = st.sidebar.date_input("Seleccione el período:", value=(fecha_min_gen, fecha_max_gen),
                                         min_value=fecha_min_gen, max_value=fecha_max_gen)
    if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
        f_desde, f_hasta = rango_fechas
    else:
        f_desde, f_hasta = fecha_min_gen, fecha_max_gen

    res_sim = calcular_be_cer(lec_ticker, cer_ticker, precio_lec_sim, precio_cer_sim, fecha_sim,
                              df_lec, df_cer, df_indice, feriados,
                              plazo_t=plazo_t_selected, modo_historico=False)

    st.subheader(f"Análisis Detallado: {lec_ticker} vs {cer_ticker} "
                 f"(Vto: {par_info['vencimiento'].strftime('%d/%m/%Y')} - {plazo_label})")

    if res_sim is None:
        st.warning("No se pudo calcular el par: revisá precios, que no haya vencido, y que el CER cargado "
                   "cubra hasta el último dato conocido.")
    elif res_sim['cer_conocido']:
        c1, c2, c3 = st.columns(3)
        c1.metric(f"TEA Fija ({lec_ticker})", f"{res_sim['tir_lec']*100:.2f}%")
        c2.metric(f"TIR {cer_ticker} (CER ya conocido)", f"{res_sim['tir_cer_fija']*100:.2f}%")
        c3.metric("TIR CER − TIR Fija", f"{(res_sim['tir_cer_fija'] - res_sim['tir_lec'])*100:+.2f} pp")
        st.info(f"El CER de referencia al vencimiento ({res_sim['f_ref_vto'].strftime('%d/%m/%Y')}) ya está "
                "publicado: la LECER tiene pago conocido y opera como tasa fija. No corresponde calcular una "
                "inflación break-even; se comparan directamente las TIR.")
    else:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric(f"TEA Fija ({lec_ticker})", f"{res_sim['tir_lec']*100:.2f}%")
        c2.metric(f"Tasa Real CER ({cer_ticker})", f"{res_sim['tir_real_cer']*100:.2f}%")
        c3.metric(f"Pesos Target ({cer_ticker})", f"${res_sim['pesos_req_cer']:.2f}")
        c4.metric("TEM BE (Mensual)", f"{res_sim['tem_be']*100:.2f}%")
        c5.metric("TEA BE (TEM Anualizada)", f"{res_sim['tea_be']*100:.2f}%")
        st.caption(f"Tramo desconocido: desde {res_sim['f_cer_ult'].strftime('%d/%m/%Y')} hasta "
                   f"{res_sim['f_ref_vto'].strftime('%d/%m/%Y')} ({res_sim['n_meses']:.2f} meses efectivos). "
                   f"Inflación acumulada implícita en ese tramo: {res_sim['infla_desconocida']*100:.2f}%.")
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
                            plazo_t=plazo_t_selected, modo_historico=True)
        if r:
            if not df_tamar.empty:
                tamar_sub = df_tamar[df_tamar['Fecha_dt'] <= f_o]
                r['tamar'] = tamar_sub['Tamar_pct'].iloc[-1] if not tamar_sub.empty else np.nan
            else:
                r['tamar'] = np.nan
            historico.append(r)
    df_hist = pd.DataFrame(historico)

    st.markdown("---")
    st.subheader("Evolución Histórica de Métricas")

    if not df_hist.empty and df_hist['tem_be'].notna().any():
        df_hist['tem_be_pct'] = df_hist['tem_be'] * 100
        df_hist['tea_be_pct'] = df_hist['tea_be'] * 100
        df_hist['tir_real_cer_pct'] = df_hist['tir_real_cer'] * 100
        st.caption("Los días en que el CER del vencimiento ya era conocido no se grafican (no hay BE).")

        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tem_be_pct'],
            name="TEM BE (%)", line=dict(color='#d62728', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEM BE:</b> %{y:.2f}%<extra></extra>"))
        fig1.update_layout(
            title=f"<b>1. Inflación Mensual Implícita (TEM BE): {lec_ticker} vs {cer_ticker}</b>",
            xaxis_title="Fecha", yaxis_title="TEM (%)", template="plotly_white", hovermode="x unified")
        st.plotly_chart(fig1, use_container_width=True)

        fig2 = make_subplots(specs=[[{"secondary_y": True}]])
        fig2.add_trace(go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tea_be_pct'], name="TEA BE (%)",
            line=dict(color='#1f77b4', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEA BE:</b> %{y:.2f}%<extra></extra>"),
            secondary_y=False)
        fig2.add_trace(go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tir_real_cer_pct'], name=f"Tasa Real CER ({cer_ticker})",
            line=dict(color='#2ca02c', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Tasa Real CER:</b> %{y:.2f}%<extra></extra>"),
            secondary_y=True)
        fig2.update_layout(
            title=f"<b>2. Inflación Anualizada Implícita (TEA BE) vs Tasa Real CER ({cer_ticker})</b>",
            template="plotly_white", hovermode="x unified")
        fig2.update_xaxes(title_text="Fecha")
        fig2.update_yaxes(title_text="<b>TEA BE (%)</b>", secondary_y=False, title_font_color='#1f77b4')
        fig2.update_yaxes(title_text="<b>Tasa Real CER (%)</b>", secondary_y=True, title_font_color='#2ca02c')
        st.plotly_chart(fig2, use_container_width=True)

        fig3 = make_subplots(specs=[[{"secondary_y": True}]])
        fig3.add_trace(go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tea_be_pct'], name="TEA BE (%)",
            line=dict(color='#1f77b4', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TEA BE:</b> %{y:.2f}%<extra></extra>"),
            secondary_y=False)
        fig3.add_trace(go.Scatter(
            x=df_hist['fecha_op'], y=df_hist['tamar'], name="Tasa TAMAR (%)",
            line=dict(color='#ff7f0e', width=2.5),
            hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR:</b> %{y:.2f}%<extra></extra>"),
            secondary_y=True)
        fig3.update_layout(title="<b>3. Inflación Anualizada Implícita (TEA BE) vs Tasa TAMAR</b>",
                           template="plotly_white", hovermode="x unified")
        fig3.update_xaxes(title_text="Fecha")
        fig3.update_yaxes(title_text="<b>TEA BE (%)</b>", secondary_y=False, title_font_color='#1f77b4')
        fig3.update_yaxes(title_text="<b>Tasa TAMAR (%)</b>", secondary_y=True, title_font_color='#ff7f0e')
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.warning("No hay datos con inflación break-even para el rango seleccionado "
                   "(o el CER de todo el período ya era conocido al vencimiento).")

# =========================================================
# VISTA: TAMAR BREAK-EVEN
# =========================================================
def render_tamar(D, precios_vivo, hoy, plazo_t, plazo_label):
    st.title("📊 Dashboard: TAMAR Break-Even")

    if D is None:
        st.error("❌ Archivo 'Inputs.xlsx' no encontrado en el directorio.")
        return
    df_lec, df_let, df_precios, df_tamar, feriados = D

    lecap_elegida = st.sidebar.selectbox("1️⃣ Tasa Fija (Lecap/Boncap):", df_lec['Ticker'].tolist())
    letam_elegida = st.sidebar.selectbox("2️⃣ Tasa Variable (Letam):", df_let['Ticker'].tolist())

    px_l_api = precios_vivo.get(_k(lecap_elegida))
    px_t_api = precios_vivo.get(_k(letam_elegida))
    usa_vivo = px_l_api is not None and px_t_api is not None

    if usa_vivo:
        def_p_l = float(normalizar_precio(px_l_api))
        def_p_t = float(normalizar_precio(px_t_api))
        fecha_sim = hoy
        st.sidebar.write(f"**Precios en vivo al:** {fecha_sim.strftime('%d/%m/%Y')}")
    else:
        if precios_vivo:
            st.sidebar.warning("La API no devolvió precio para este par. Usando Excel.")
        df_validos = df_precios.dropna(subset=[lecap_elegida, letam_elegida])
        if df_validos.empty:
            st.error("❌ No hay fechas con precio para ambas especies seleccionadas.")
            return
        ultima_fila = df_validos.iloc[-1]
        fecha_sim = ultima_fila['Fecha_dt']
        def_p_l = float(normalizar_precio(ultima_fila[lecap_elegida]))
        def_p_t = float(normalizar_precio(ultima_fila[letam_elegida]))
        st.sidebar.write(f"**Última fecha con ambos precios:** {fecha_sim.strftime('%d/%m/%Y')}")

    st.sidebar.markdown("---")
    precio_lecap_sim = st.sidebar.number_input(f"Precio {lecap_elegida}:", value=def_p_l, step=0.10, format="%.2f")
    precio_letam_sim = st.sidebar.number_input(f"Precio {letam_elegida}:", value=def_p_t, step=0.10, format="%.2f")

    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Rango Histórico a Graficar")
    fecha_min_general = df_precios['Fecha_dt'].min()
    fecha_max_general = df_precios['Fecha_dt'].max()
    rango_fechas = st.sidebar.date_input("Seleccioná el período:", value=(fecha_min_general, fecha_max_general),
                                         min_value=fecha_min_general, max_value=fecha_max_general)
    if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
        f_desde, f_hasta = rango_fechas
    else:
        f_desde, f_hasta = fecha_min_general, fecha_max_general

    # Control de unidades: API vs Excel
    if precios_vivo:
        tickers_tam = {_k(t) for t in df_lec['Ticker']} | {_k(t) for t in df_let['Ticker']}
        with st.expander("🔎 Control de precios: API vs Excel (verificar unidades)"):
            ult_xls = df_precios.iloc[-1]
            filas = []
            for t in sorted(t for t in precios_vivo if t in tickers_tam):
                px_api = normalizar_precio(precios_vivo[t])
                col = next((c for c in ult_xls.index if _k(c) == t), None)
                px_xls = normalizar_precio(ult_xls[col]) if col is not None else np.nan
                filas.append({'Ticker': t, 'API': px_api, 'Excel (última fila)': px_xls,
                              'Dif %': (px_api / px_xls - 1) * 100 if pd.notnull(px_xls) and px_xls else np.nan})
            st.dataframe(pd.DataFrame(filas).style.format(
                {'API': '{:.2f}', 'Excel (última fila)': '{:.2f}', 'Dif %': '{:+.2f}%'}, na_rep="-"),
                use_container_width=True, hide_index=True)

    res_sim = calcular_tamar_be(lecap_elegida, letam_elegida, precio_lecap_sim, precio_letam_sim, fecha_sim,
                                df_lec, df_let, df_tamar, feriados, plazo_t=plazo_t)

    if res_sim is None:
        st.error("⚠️ No se pudo calcular la simulación: revisá que las especies no hayan vencido a la fecha "
                 "y que haya datos de TAMAR para la ventana de fixing.")
        return

    st.caption(f"Simulación al {fecha_sim.strftime('%d/%m/%Y')} en **{plazo_label}** con precios "
               f"{'en VIVO (1816)' if usa_vivo else 'del Excel'}.")

    col1, col2, col3 = st.columns(3)
    col1.metric(label=f"TNA {lecap_elegida} (Fija)", value=f"{res_sim['tna_lec']*100:.2f}%")
    col2.metric(label=f"VPV Req {letam_elegida} (Var)", value=f"${res_sim['vpv_objetivo']:.2f}")

    if res_sim['es_fija']:
        col3.metric(label="Spread TAMAR BE vs Mkt", value="Ya es Fija")
        st.markdown("---")
        st.info("⚠️ **El período de fixing finalizó:** la Letam ya devengó toda la TAMAR y opera como "
                "instrumento a tasa fija, por lo que no corresponde calcular una Break-Even futura.")
        if not np.isnan(res_sim['tir_let']):
            ca, cb = st.columns(2)
            ca.info(f"**TIR {letam_elegida} (ya conocida):** {res_sim['tir_let']*100:.2f}%")
            cb.info(f"**TIR {lecap_elegida}:** {res_sim['tir_lec']*100:.2f}%")
    else:
        col3.metric(label="Spread TAMAR BE vs Mkt", value=f"{res_sim['spread_pkt']:+.2f} pkt")
        st.markdown("---")
        col_a, col_b = st.columns(2)
        col_a.info(f"**TAMAR Break-Even Futura:** {res_sim['tamar_be']*100:.2f}%")
        col_b.info(f"**TAMAR Mercado Hoy:** {res_sim['tamar_mercado']*100:.2f}%")
        st.caption(
            f"Ventana de fixing: {res_sim['f_inicio_fixing']:%d/%m/%Y} a {res_sim['f_fin_fixing']:%d/%m/%Y} "
            f"({res_sim['n_total']} hábiles; {res_sim['n_dev']} conocidos, {res_sim['n_proy']} a proyectar). "
            "El BE es la TAMAR constante sobre los días a proyectar que iguala la TIR de la Lecap. "
            "Supone reinversión a la TIR de la Lecap si ésta vence antes que la Letam.")

    # ---- Histórica (siempre con precios del Excel) ----
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

    st.subheader("📈 Evolución Histórica: BE vs Mercado")
    if not df_hist.empty:
        df_hist_plot = df_hist[~df_hist['es_fija']].copy()
        if not df_hist_plot.empty:
            df_hist_plot['be_pct'] = df_hist_plot['tamar_be'] * 100
            df_hist_plot['mkt_pct'] = df_hist_plot['tamar_mercado'] * 100

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=df_hist_plot['fecha'], y=df_hist_plot['be_pct'], mode='lines',
                name=f'TAMAR BE ({lecap_elegida} vs {letam_elegida})',
                line=dict(color='#d62728', width=2.5),
                hovertemplate='<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR BE:</b> %{y:.2f}%<extra></extra>'))
            fig.add_trace(go.Scatter(
                x=df_hist_plot['fecha'], y=df_hist_plot['mkt_pct'], mode='lines',
                name='TAMAR Mkt Real',
                line=dict(color='#1f77b4', width=2.5, dash='dash'),
                hovertemplate='<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR Mkt Real:</b> %{y:.2f}%<extra></extra>'))
            fig.update_layout(
                hovermode='x unified', yaxis_title='TNA (%)', template='plotly_white',
                height=450, margin=dict(l=10, r=10, t=30, b=10),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("⚠️ Todos los datos en este rango corresponden al período donde la Letam ya operaba "
                       "como tasa fija (cierre de fixing).")
    else:
        st.warning("No hay datos válidos en el rango de fechas seleccionado.")

    # ---- Test de validación del motor ----
    with st.expander("🧪 Test de validación del motor (TAMAR sintética constante)"):
        x_test = st.number_input("TAMAR constante a testear (TNA %):", value=30.0, step=1.0, format="%.2f") / 100.0
        if st.button("Correr test"):
            r_test, info = test_sintetico(lecap_elegida, letam_elegida, precio_lecap_sim, x_test,
                                          df_lec, df_let, feriados)
            if r_test is None:
                st.error(info)
            else:
                err = abs(r_test['tamar_be'] - x_test) * 100
                st.write(f"Fecha de operación del test: {info:%d/%m/%Y} "
                         f"(n_total={r_test['n_total']}, n_dev={r_test['n_dev']}, n_proy={r_test['n_proy']})")
                st.write(f"BE obtenido: **{r_test['tamar_be']*100:.6f}%** vs esperado **{x_test*100:.6f}%** "
                         f"(error: {err:.2e} pp)")
                if err < 1e-6:
                    st.success("✅ El motor reproduce la TAMAR constante.")
                else:
                    st.error("❌ El BE no coincide con la TAMAR sintética: hay un problema en el motor.")

# =========================================================
# MAIN: NAVEGACIÓN + API COMPARTIDA
# =========================================================
D_cer = cargar_datos_cer()
D_tamar = cargar_datos_tamar()

# Feriados para saber si el mercado está operando (se toman de cualquiera de los dos Excel)
feriados_g = D_cer[5] if D_cer is not None else (D_tamar[4] if D_tamar is not None else set())
ahora = datetime.now(TZ_AR)
hoy = ahora.date()
mercado_abierto = es_habil(hoy, feriados_g) and 11 <= ahora.hour < 17  # horario aproximado BYMA

st.sidebar.title("Renta Fija")
vista = st.sidebar.radio("Calculadora:", ["Inflación BE (CER)", "TAMAR BE"], key="vista")
st.sidebar.markdown("---")
st.sidebar.header("Parámetros de Simulación")

usar_api = st.sidebar.checkbox("⚡ Usar precios en VIVO (API 1816)", value=False)
if usar_api and st.sidebar.button("🔄 Actualizar precios"):
    obtener_precios_1816.clear()  # fuerza una nueva consulta a la API
precios_vivo = {}
if usar_api:
    api_key = _get_api_key()
    if not api_key:
        st.sidebar.error("Falta API_1816_KEY (en st.secrets o variable de entorno). Usando Excel.")
    else:
        todos = set()
        if D_cer is not None:
            df_lec_c, df_cer_c = D_cer[0], D_cer[1]
            todos |= {_k(t) for t in df_lec_c['Ticker']} | {_k(t) for t in df_cer_c['Ticker']}
        if D_tamar is not None:
            df_lec_t, df_let_t = D_tamar[0], D_tamar[1]
            todos |= {_k(t) for t in df_lec_t['Ticker']} | {_k(t) for t in df_let_t['Ticker']}
        try:
            with st.sidebar.status("Conectando con 1816..."):
                precios_vivo, ts_api = obtener_precios_1816(tuple(sorted(todos)), api_key)
            st.sidebar.success(f"✅ Precios 1816 obtenidos ({len(precios_vivo)} especies).")
            st.sidebar.caption(f"Última consulta: {ts_api:%H:%M:%S} (se reutiliza hasta 5 min)")
            if not mercado_abierto:
                st.sidebar.warning("Mercado cerrado (fuera de horario, fin de semana o feriado): los precios "
                                   "pueden ser del último cierre, no de operaciones de hoy.")
        except Exception as e:
            st.sidebar.error(f"❌ Falló API: {e}. Usando Excel.")
            precios_vivo = {}

st.sidebar.markdown("---")

# Plazo de liquidación global (aplica a las dos calculadoras)
plazo_opciones = {"T+0 (Contado Inmediato)": 0, "T+1 (24hs)": 1}
plazo_label = st.sidebar.selectbox("Plazo de Liquidación:", list(plazo_opciones.keys()), index=1)
plazo_t = plazo_opciones[plazo_label]
st.sidebar.markdown("---")

if vista == "Inflación BE (CER)":
    render_cer(D_cer, precios_vivo, hoy, plazo_t, plazo_label)
else:
    render_tamar(D_tamar, precios_vivo, hoy, plazo_t, plazo_label)