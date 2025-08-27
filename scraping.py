import requests
import pandas as pd
from datetime import date, datetime, timedelta
from typing import Optional
from io import StringIO
from typing import List, Optional  # <-- IMPORTANTE

def _is_weekend(d: datetime) -> bool:
    return d.weekday() >= 5  # 5=Saturday, 6=Sunday

def _prev_business_day(d: datetime) -> datetime:
    """Si la fecha cae en fin de semana, retrocede al día hábil previo."""
    while _is_weekend(d):
        d -= timedelta(days=1)
    return d

def stooq_close_on_date(ticker: str, date_str: Optional[str] = None) -> Optional[float]:
    """
    Devuelve el precio de cierre de `ticker` en Stooq para la fecha dada.
    - ticker: ej. 'fig.us', 'mnst.us', '1810.hk'
    - date_str: 'YYYY-MM-DD' o None → usa la fecha de hoy (o último hábil)
    """
    # Si no se pasa fecha, usar la de hoy
    if not date_str:
        date_str = date.today().strftime("%Y-%m-%d")

    # Convertir a datetime
    d = datetime.strptime(date_str, "%Y-%m-%d")
    d = _prev_business_day(d)  # ajustar si es fin de semana

    d_compact = d.strftime("%Y%m%d")
    url = f"https://stooq.com/q/d/l/?s={ticker.lower()}&d1={d_compact}&d2={d_compact}&i=d"

    r = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()

    text = r.text.strip()
    if not text or len(text.splitlines()) <= 1:
        return None

    df = pd.read_csv(StringIO(text))
    if df.empty or "Close" not in df.columns or pd.isna(df.loc[0, "Close"]):
        return None

    return float(df.loc[0, "Close"])

# Ejemplos:
body = "FIG 2025-08-19:" + str(stooq_close_on_date("fig.us","2025-08-18")) + "\n"
body+= "Itau 2025-08-18:" + str(stooq_close_on_date("itub.us", "2025-08-18")) + "\n"
body+= "Monster 2025-08-18:" + str(stooq_close_on_date("mnst.us", "2025-08-18")) + "\n"
body+= "Xiaomi 2025-08-19:" + str(stooq_close_on_date("1810.hk", "2025-08-18")) + "\n"
print(body)
