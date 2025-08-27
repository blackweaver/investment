# https://polygon.io/dashboard/keys/4040a189-9e0d-4b0a-8c8e-5a1871da6bcf?values=flat-files
import os
import sys
import json
import ast
import requests
import argparse
import time
import math
import html
from tabulate import tabulate
import pandas as pd
from typing import Dict, Any, List
from typing import Optional, List
from datetime import datetime, date, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import PatternFill, Font
from dotenv import load_dotenv
load_dotenv()

# Variables de entorno (tal como pediste)
API_KEY_POLYGON = os.getenv("API_KEY_POLYGON")
TICKETS = os.getenv("TICKETS")

parser = argparse.ArgumentParser()
parser.add_argument('--top10', action='store_true', help='Ejecutar lógica para Top 10')
parser.add_argument('--bajas', action='store_true', help='Ejecutar lógica para Bajas')
parser.add_argument('--telegram', action='store_true', help='Enviar un mensaje al bot de telegram con la tabla de totales del excel')
parser.add_argument(
    "--date",
    type=lambda s: datetime.strptime(s, "%Y-%m-%d").strftime("%Y-%m-%d"),
    default=date.today().strftime("%Y-%m-%d"),
    help="Fecha en formato YYYY-MM-DD (por defecto: hoy)"
)


args = parser.parse_args()
top10 = args.top10
bajas = args.bajas
telegram = args.telegram

# Endpoints Polygon
SNAPSHOT_URL = "https://api.polygon.io/v2/snapshot/locale/us/markets/stocks/tickers"
PREV_URL_TMPL = "https://api.polygon.io/v2/aggs/ticker/{ticker}/prev"
REF_TICKER_URL_TMPL = "https://api.polygon.io/v3/reference/tickers/{ticker}"
HIST_URL_TMPL = "https://api.polygon.io/v2/aggs/ticker/{ticker}/range/1/day/{date}/{date}"


EXCEL_FILE = "actions.xlsx"

# Cargar o crear archivo
if os.path.exists(EXCEL_FILE):
    book = load_workbook(EXCEL_FILE)
else:
    book = Workbook()
    book.remove(book.active)

try:
    # SDK 1.x
    from openai import OpenAI
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    def chat(model, messages):
        return client.chat.completions.create(model=model, messages=messages)
except ImportError:
    # SDK 0.28 (legacy)
    import openai
    openai.api_key = os.getenv("OPENAI_API_KEY")
    def chat(model, messages):
        return openai.ChatCompletion.create(model=model, messages=messages)

# Uso:
resp = chat("gpt-4o-mini", [{"role":"user","content":"Hola"}])


# Totales
# Crear hoja solo si no existe (se regenera cada vez)
if "Totales" in book.sheetnames:
    del book["Totales"]

sheet = book.create_sheet("Totales")

# Top 10
if top10:
     # Crear hoja "top10"
    if "Top 10 empresas del momento" in book.sheetnames:
        del book["Top 10 empresas del momento"]

    sheet = book.create_sheet("Top 10 empresas del momento")

    def get_top10_crypto(max_retries=3, delay=2):
        prompt_messages = [
            {"role": "system", "content": (
                "I am aware that you are not a financial analyst, far from it."
                "I also know that no serious financial analyst would dare to predict the future or induce people to invest based on their recommendations."
                "However, I want you to gather the opinions of recognized analysts and tell me which 10 companies are presumed to have their stocks on the rise in the next six months."
            )},
            {"role": "user", "content": (
                "Your only response should be a valid Python list, for example ['Figma','Monster','Itaú','Xiaomi', ... 'reflection...'], with no explanations, no additional text."
                "Return them ordered, first the one expected to rise the most and so on."
                "The last item in the array should not be one of the rising companies, but a random suggestion of a good trading practice to improve, considering general investment knowledge, but also practical suggestions for eToro tools."
            )}
        ]

        for attempt in range(1, max_retries + 1):
            try:
                response = client.chat.completions.create(model="gpt-4",
                messages=prompt_messages)
                top10_text = response.choices[0].message.content

                # Intentar interpretar como lista Python
                top10 = ast.literal_eval(top10_text)

                # Validar que sea lista con al menos un string
                if isinstance(top10, list) and all(isinstance(x, str) for x in top10) and len(top10) > 0:
                    return top10
                else:
                    print("⚠️ La respuesta no es una lista válida de strings no vacía.")

            except Exception as e:
                print(f"⚠️ Error al interpretar la respuesta (intento {attempt}): {e}")

            if attempt < max_retries:
                print("🔁 Reintentando...")
                time.sleep(delay)

        # Si todo falla, retornar None
        return None

    # Llamar a la función
    top10 = get_top10_crypto()

    if top10:
        print("✅ Top 10 obtenido: ", top10)
        top10_telegram = "✅ Top 3 obtenido: " + ", ".join(top10[:3])
        df_top10 = pd.DataFrame({"Top 10 empresas del momento": top10})

        for r in dataframe_to_rows(df_top10, index=False, header=True):
            sheet.append(r)
            
         # --- Función para buscar si ya existe una fila con cierto texto ---
        def find_row_index(sheet, target_text):
            for i, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                if row and row[0] == target_text:
                    return i
            return None

        # --- Insertar o sobrescribir Top 10 ---
        top10_header = "Top 10 empresas del momento"
        top10_row_idx = find_row_index(sheet, top10_header)

        if top10_row_idx:
            for i, r in enumerate(dataframe_to_rows(df_top10, index=False, header=True)):
                for j, value in enumerate(r, start=1):
                    sheet.cell(row=top10_row_idx + i, column=j, value=value)
        else:
            for r in dataframe_to_rows(df_top10, index=False, header=True):
                sheet.append(r)
    else:
        print("❌ No se pudo obtener un Top 10 válido.")

if bajas:
     # Crear hoja "top10"
    if "Top 10 acciones bajas" in book.sheetnames:
        del book["Top 10 acciones bajas"]

    sheet = book.create_sheet("Top 10 acciones bajas")

    def get_top10_bajas_acciones(max_retries=3, delay=2):
        prompt_messages = [
            {"role": "system", "content": (
                "I am aware that you are not a financial analyst, far from it."
                "I also know that no serious financial analyst would dare to predict the future or induce people to invest based on their recommendations."
                "However, I want you to gather the opinions of the most prestigious analysts and tell me which 10 companies that start at a very low actions price are presumed to have potential to rise in the next six months."
            )},
            {"role": "user", "content": (
                "Your only response should be a valid Python list, for example ['Figma','Monster','Itaú','Xiaomi', ... 'reflection...'], with no explanations, no additional text."
                "The last item in the array should not be one of the rising companies, but a random suggestion of a good trading practice to improve, considering general investment knowledge, but also practical suggestions for eToro tools."
            )}
        ]

        for attempt in range(1, max_retries + 1):
            try:
                response = client.chat.completions.create(model="gpt-4",
                messages=prompt_messages)
                top10_bajas_text = response.choices[0].message.content

                # Intentar interpretar como lista Python
                top10_bajas = ast.literal_eval(top10_bajas_text)

                # Validar que sea lista con al menos un string
                if isinstance(top10_bajas, list) and all(isinstance(x, str) for x in top10_bajas) and len(top10_bajas) > 0:
                    return top10_bajas
                else:
                    print("⚠️ La respuesta no es una lista válida de strings no vacía.")

            except Exception as e:
                print(f"⚠️ Error al interpretar la respuesta (intento {attempt}): {e}")

            if attempt < max_retries:
                print("🔁 Reintentando...")
                time.sleep(delay)

        # Si todo falla, retornar None
        return None

    # Llamar a la función
    top10_bajas = get_top10_bajas_acciones()

    if top10_bajas:
        print("✅ Top 10 de bajas obtenido:", top10_bajas)
        bajas_telegram = "✅ Top 3 bajas con potencial: " + ", ".join(top10_bajas[:3])
        df_top10_bajas = pd.DataFrame({"Top 10 acciones a bajo costo con potencial ": top10_bajas})

        for r in dataframe_to_rows(df_top10_bajas, index=False, header=True):
            sheet.append(r)
        
         # --- Función para buscar si ya existe una fila con cierto texto ---
        def find_row_index(sheet, target_text):
            for i, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                if row and row[0] == target_text:
                    return i
            return None

        # --- Insertar o sobrescribir Top 10 ---
        top10_bajas_header = "Top 10 acciones a bajo costo con potencial"
        top10_bajas_row_idx = find_row_index(sheet, top10_bajas_header)

        if top10_bajas_row_idx:
            for i, r in enumerate(dataframe_to_rows(df_top10_bajas, index=False, header=True)):
                for j, value in enumerate(r, start=1):
                    sheet.cell(row=top10_bajas_row_idx + i, column=j, value=value)
        else:
            for r in dataframe_to_rows(df_top10_bajas, index=False, header=True):
                sheet.append(r)
    else:
        print("❌ No se pudo obtener un Top 10 válido.")

# ---------------------------
# Excepciones y utilidades
# ---------------------------

class PolygonForbidden(Exception):
    """HTTP 403: tu plan no está habilitado para el endpoint/dato solicitado."""

def _safe_get(d: dict, path: List[str], default=None):
    cur = d
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur

def _http_get_json(url: str, params: dict) -> dict:
    """GET con manejo de errores y retorno JSON."""
    r = requests.get(url, params=params, timeout=15)
    try:
        r.raise_for_status()
        return r.json()
    except requests.HTTPError as e:
        # mensaje útil si viene JSON
        try:
            msg = r.json().get("message") or r.text[:400]
        except Exception:
            msg = r.text[:400]
        if r.status_code == 403:
            raise PolygonForbidden(f"403 Forbidden: {msg}") from e
        raise RuntimeError(f"HTTP {r.status_code}: {msg}") from e
    except requests.RequestException as e:
        raise RuntimeError(f"Error de red: {e}") from e

# ---------------------------
# Nombres de compañías
# ---------------------------

def _company_name(ticker: str, api_key: str, cache: Dict[str, str]) -> str:
    """
    Busca el nombre oficial de la compañía para un ticker.
    Si falla o no hay acceso, devuelve el mismo ticker.
    """
    if ticker in cache:
        return cache[ticker]
    try:
        data = _http_get_json(REF_TICKER_URL_TMPL.format(ticker=ticker), {"apiKey": api_key})
        name = _safe_get(data, ["results", "name"])
        cache[ticker] = (name or "").strip() or ticker
    except Exception:
        cache[ticker] = ticker
    return cache[ticker]

# ---------------------------
# Extracción de precios
# ---------------------------

def _extract_prices_from_snapshot(snapshot_payload: dict) -> Dict[str, float]:
    """
    Desde /v2/snapshot...:
    - Prioriza lastTrade.p (último precio)
    - Fallback: day.c (close del día)
    Devuelve {ticker: precio}
    """
    out: Dict[str, float] = {}
    for item in snapshot_payload.get("tickers", []):
        t = item.get("ticker")
        if not t:
            continue
        price = _safe_get(item, ["lastTrade", "p"])
        if price is None:
            price = _safe_get(item, ["day", "c"])
        if price is not None:
            try:
                out[t] = float(price)
            except (TypeError, ValueError):
                pass
    return out

def _extract_prices_from_prev(prev_dict: Dict[str, Any]) -> Dict[str, float]:
    """
    Desde /v2/aggs/{ticker}/prev:
    - Toma results[0].c (cierre del día anterior)
    Devuelve {ticker: precio}
    """
    out: Dict[str, float] = {}
    for t, payload in prev_dict.items():
        results = payload.get("results") or []
        if results:
            c = results[0].get("c")
            if c is not None:
                try:
                    out[t] = float(c)
                except (TypeError, ValueError):
                    pass
    return out

# ---------------------------
# Lógica principal API Polygon
# ---------------------------

def fetch_price_on_date(ticker: str, date: str, api_key: str) -> dict:
    """
    Devuelve OHLC de un ticker en una fecha específica (YYYY-MM-DD).
    """
    url = HIST_URL_TMPL.format(ticker=ticker, date=date)
    return _http_get_json(url, {"adjusted": "true", "apiKey": api_key})

def fetch_quotes_polygon(symbols_csv: str, api_key: str, date: Optional[str] = None) -> dict:
    """
    Si se pasa una fecha (YYYY-MM-DD), consulta /open-close para cada ticker.
    Si no, intenta una sola llamada al snapshot múltiple.
    Si 403 (no habilitado), hace fallback consultando /prev por cada ticker.
    Retorna: {"mode": "snapshot"|"prev"|"historical", "data": <payload>}
    """
    symbols: List[str] = [s.strip().upper() for s in symbols_csv.split(",") if s.strip()]
    if not symbols:
        raise ValueError("Debes pasar al menos un símbolo separado por coma.")

    # Caso histórico: se pidió una fecha
    if date:
        results: Dict[str, Any] = {}
        for t in symbols:
            hist = _http_get_json(HIST_URL_TMPL.format(ticker=t, date=date),
                                  {"adjusted": "true", "apiKey": api_key})
            results[t] = hist
        return {"mode": "historical", "data": results}

    # Caso normal: snapshot múltiple
    try:
        snap = _http_get_json(SNAPSHOT_URL, {"tickers": ",".join(symbols), "apiKey": api_key})
        return {"mode": "snapshot", "data": snap}
    except PolygonForbidden:
        # Fallback: previous close por cada ticker
        results: Dict[str, Any] = {}
        for t in symbols:
            prev = _http_get_json(PREV_URL_TMPL.format(ticker=t),
                                  {"adjusted": "true", "apiKey": api_key})
            results[t] = prev
        return {"mode": "prev", "data": results}
    
current_investment = {}
def print_quotes_polygon(symbols_csv: str, api_key: str, date: Optional[str] = None) -> None:
    global current_investment
    """
    Imprime un objeto JSON { "Nombre Compañía": precio }.
    - 'snapshot': usa lastTrade.p o day.c (del día en curso).
    - 'prev': usa el cierre del día anterior.
    """
    
    if date:
        resp = fetch_quotes_polygon(symbols_csv, api_key, date)
    else:
        resp = fetch_quotes_polygon(symbols_csv, api_key)
    mode = resp.get("mode")
    data = resp.get("data", {})

    # 1) {ticker: price}
    if mode == "snapshot":
        prices_by_ticker = _extract_prices_from_snapshot(data)
    else:
        prices_by_ticker = _extract_prices_from_prev(data)

    # 2) Resolver nombres y mapear {Nombre: precio}
    name_cache: Dict[str, str] = {}
    mapped: Dict[str, float] = {}
    for t, price in prices_by_ticker.items():
        name = _company_name(t, api_key, name_cache)
        mapped[name] = price

    current_investment = json.dumps(mapped, indent=2, ensure_ascii=False)
    print(current_investment)

if __name__ == "__main__":
    if not API_KEY_POLYGON or not TICKETS:
        print(
            file=sys.stderr
        )
        sys.exit(2)

    try:
        # Usa exactamente las variables que definiste
        print("FECHA: ", {args.date})
        print_quotes_polygon(TICKETS, API_KEY_POLYGON, args.date)
        # Guardar archivo
        book.save(EXCEL_FILE)
        print(f"✅ Archivo actualizado: {EXCEL_FILE}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
        
        
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
DEFAULT_CHAT_ID = os.getenv("DEFAULT_CHAT_ID")
TELEGRAM_API = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
# Límite seguro para texto en Telegram (los bots aceptan ~4096 chars)
TG_LIMIT = 4000

def _read_sheet_values(excel_path: str, sheet_name: str) -> pd.DataFrame:

    wb = load_workbook(excel_path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    rows = list(ws.values)
    if not rows:
        return pd.DataFrame()
    headers = list(rows[0])
    data = rows[1:]
    df = pd.DataFrame(data, columns=headers)
    return df

def _format_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """2 dec para USD/totales, 8 dec para cantidades; resto texto."""
    out = df.copy()
    usd_like = { "Cripto",
        "Cantidad Total",
        "Precio Compra Total (USD)",
        "Cantidad Actual",
        "Ganancia/Perdida (USD)"}
    for c in out.columns:
        if pd.api.types.is_numeric_dtype(out[c]):
            if "Cantidad" in c and "USD" not in c:
                out[c] = out[c].map(lambda x: "" if pd.isna(x) else f"{float(x):,.8f}".replace(",", "X").replace(".", ",").replace("X", "."))
            elif c in usd_like or "USD" in c:
                out[c] = out[c].map(lambda x: "" if pd.isna(x) else f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            else:
                out[c] = out[c].map(lambda x: "" if pd.isna(x) else f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            out[c] = out[c].astype(str)
    return out

def _send_html_pre(chat_id: str, title: str, body: str):
    if not TELEGRAM_TOKEN:
        raise RuntimeError("Falta TELEGRAM_TOKEN")
    payload = {
        "chat_id": chat_id or DEFAULT_CHAT_ID,
        "text": f"{html.escape(title)}\n<pre>{html.escape(body)}</pre>",
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    r = requests.post(TELEGRAM_API, json=payload, timeout=30)
    r.raise_for_status()

def send_telegram_table(
    sheet_name: str,
    *,
    excel_path: str,
    columns: Optional[List[str]] = None,
    top: Optional[int] = None,
    sort_by: Optional[str] = None,
    ascending: bool = False,
    rows_per_msg: int = 12,
    tablefmt: str = "github",
    title: Optional[str] = None,
    chat_id: Optional[str] = None,
) -> None:
    chat_id = chat_id or DEFAULT_CHAT_ID
    if not chat_id:
        raise RuntimeError("Falta DEFAULT_CHAT_ID/TELEGRAM_CHAT_ID")
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No encuentro el Excel en: {excel_path}")

    # 1) Leer hoja
    df = _read_sheet_values(excel_path, sheet_name)

    # 2) Filtrar y ordenar
    if columns:
        cols = [c for c in columns if c in df.columns]
        df = df[cols] if cols else df
    if sort_by and sort_by in df.columns:
        df = df.sort_values(sort_by, ascending=ascending)

    if top is not None:
        df = df.head(top)

    # === NUEVO: calcular totales y diferencia sobre datos numéricos ===
    col_precio = "Precio Compra Total (USD)"
    col_cant   = "Cantidad Actual"

    precio_num = pd.to_numeric(df[col_precio], errors="coerce") if col_precio in df.columns else pd.Series(dtype=float)
    cant_num   = pd.to_numeric(df[col_cant], errors="coerce")   if col_cant   in df.columns else pd.Series(dtype=float)

    total_precio = float(precio_num.sum(skipna=True)) if len(precio_num) else 0.0
    total_cant   = float(cant_num.sum(skipna=True))   if len(cant_num)   else 0.0
    diferencia   = total_cant - total_precio

    # 5) Limpieza final para visual
    df = df.replace({None: pd.NA}).fillna("")

    # 4) Formatear números para la tabla
    df_fmt = _format_numeric_columns(df)
    n = len(df_fmt)
    pages = max(1, math.ceil(n / rows_per_msg)) if rows_per_msg else 1

    for p in range(pages):
        start = p * rows_per_msg
        end = None if not rows_per_msg else start + rows_per_msg
        page_df_fmt = df_fmt.iloc[start:end]  # 👈 usar el slice correcto

        # body = tabulate(
        #     page_df_fmt,
        #     headers="keys",
        #     tablefmt=tablefmt,
        #     numalign="right",
        #     stralign="left",
        #     showindex=False
        # )

        # === NUEVO: agregar totales al final de la ÚLTIMA página ===
        #if p == pages - 1:
#
        #    footer_lines = [
        #        "",
        #        f"Total Precio Compra: {total_precio:,.0f}",
        #        f"Total Cantidad Actual: {total_cant:,.0f}",
        #        f"✅ Ganancia: {diferencia:,.0f}" if (diferencia > 0) else f"❌ Pérdida: {diferencia:,.0f}",
        #    ]
        #    body = body + "\n" + "\n".join(footer_lines) + "\n\n"
        
        
        data = json.loads(current_investment)
        body = "\n".join(f"{name}: {price}" for name, price in data.items()) + "\n\n"
        print(body)
            
        if top10_telegram != "":
            body = body + "\n" + top10_telegram
        if bajas_telegram != "":
            body = body + "\n" + bajas_telegram

        # Recorte ultraseguro si excede límites
        if len(body) > 3900:
            body = "\n".join(line[:120] for line in body.splitlines())

        page_title = f"{title or sheet_name} ({p+1}/{pages})" if pages > 1 else (title or sheet_name)
        _send_html_pre(chat_id, page_title, body)
        
if telegram:
    headers_list = [
        "Cripto",
        "Cantidad Total",
        "Precio Compra Total (USD)",
        "Cantidad Actual",
        "Ganancia/Perdida (USD)"
    ]

    send_telegram_table(
        sheet_name="Totales",
        excel_path=EXCEL_FILE,        # usa tu constante existente
        columns=headers_list,
        top=20,
        sort_by="Total Ganancia",     # opcional: ordena por ganancia
        ascending=False,
        rows_per_msg=12,
        tablefmt="github",
        title="Totales",
        chat_id=DEFAULT_CHAT_ID
    )