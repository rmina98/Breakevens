"""
Guarda el cierre diario de 1816 (precioDirty) de todas las especies de las hojas
Lecaps, LECER y Letam de Inputs.xlsx en precios_1816.csv.

- Lo corre GitHub Actions todos los días hábiles después del cierre (ver .github/workflows/cierres_be.yml).
- NO escribe el Excel: Precios/CER/Tamar tienen fórmulas a libros externos y si se guarda el
  archivo desde Python se pierden los valores. La app combina Excel + este CSV al cargar.
- Si ya existe la fila del día, la reemplaza (se puede correr varias veces sin duplicar).

Uso local:  API_1816_KEY=xxxx python guardar_cierres_be.py  [--forzar]
"""
import os
import sys
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")
BASE = "https://api.1816.com.ar/v1"
ARCHIVO_CSV = "precios_1816.csv"
HOJAS_ESPECIES = ("Lecaps", "LECER", "Letam")


def ruta_inputs():
    for nombre in ("Inputs.xlsx", "inputs.xlsx", "INPUTS.xlsx"):
        if os.path.exists(nombre):
            return nombre
    sys.exit("No se encontró Inputs.xlsx en el directorio.")


def leer_tickers_y_feriados(path):
    xls = pd.ExcelFile(path)
    tickers = set()
    for hoja in HOJAS_ESPECIES:
        if hoja in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=hoja)
            df.columns = df.columns.str.strip()
            tickers |= {str(t).strip().upper() for t in df["Ticker"].dropna()}
    feriados = set()
    if "Feriados" in xls.sheet_names:
        f = pd.read_excel(xls, sheet_name="Feriados")
        feriados = set(pd.to_datetime(f["Feriados"], errors="coerce").dropna().dt.date)
    return sorted(tickers), feriados


def precios_1816(tickers, api_key):
    auth = requests.post(f"{BASE}/auth/token", json={"apiKey": api_key, "module": "mercado"}, timeout=15)
    auth.raise_for_status()
    token = auth.json()["token"]
    time.sleep(1.1)  # plan Mercado-Base: máx. 1 petición por segundo

    r = requests.get(f"{BASE}/mercado/indicadores", headers={"Authorization": f"Bearer {token}"},
                     params={"tickers": [t.lower() for t in tickers], "campos": ["precioDirty"]}, timeout=30)
    r.raise_for_status()
    data = r.json()
    precios = {}
    for ticker, valores in (data.get("instrumentos") or {}).items():
        if isinstance(valores, dict) and valores.get("precioDirty") is not None:
            precios[ticker.upper()] = float(valores["precioDirty"])
    return precios


def main():
    forzar = "--forzar" in sys.argv
    api_key = os.environ.get("API_1816_KEY")
    if not api_key:
        sys.exit("Falta la variable de entorno API_1816_KEY.")

    hoy = datetime.now(TZ_AR).date()
    tickers, feriados = leer_tickers_y_feriados(ruta_inputs())
    if not forzar and (hoy.weekday() >= 5 or hoy in feriados):
        print(f"{hoy:%d/%m/%Y} no es día hábil: no se guarda nada.")
        return

    precios = precios_1816(tickers, api_key)
    if not precios:
        sys.exit("1816 no devolvió precioDirty para ninguna especie.")  # falla el Action y te llega el mail

    faltan = [t for t in tickers if t not in precios]
    if faltan:
        print("Sin precio en 1816 para:", ", ".join(faltan))

    fila = pd.DataFrame([{"Fecha": pd.Timestamp(hoy), **precios}])
    if os.path.exists(ARCHIVO_CSV):
        hist = pd.read_csv(ARCHIVO_CSV, parse_dates=["Fecha"])
        hist = hist[hist["Fecha"].dt.date != hoy]
        hist = pd.concat([hist, fila], ignore_index=True)
    else:
        hist = fila
    hist = hist.sort_values("Fecha")
    cols = ["Fecha"] + sorted(c for c in hist.columns if c != "Fecha")
    hist[cols].to_csv(ARCHIVO_CSV, index=False, date_format="%Y-%m-%d")
    print(f"Guardado cierre {hoy:%d/%m/%Y}: {len(precios)} especies -> {ARCHIVO_CSV}")


if __name__ == "__main__":
    main()
