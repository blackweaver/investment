# Registro de inversiones

Scripts de Python para llevar un registro de acciones y criptomonedas en Excel. Permiten registrar movimientos, consultar precios, calcular resultados y enviar resúmenes a Telegram.

- `actions.py`: registra compras y ventas de acciones en `actions.xlsx`, con cotizaciones de Polygon.
- `crypto.py`: registra movimientos en `crypto.xlsx`, consulta precios en CoinGecko e históricos en Binance.
- `scraping.py`: consulta precios de cierre en Stooq; incluye ejemplos que se ejecutan al correr el archivo.
- `crypto.ipynb`: permite explorar los datos desde Jupyter.

## Instalación

Desde la carpeta del proyecto, creá y activá un entorno virtual:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install pandas openpyxl python-dotenv requests tabulate 'openai>=1,<2'
```

`requirements.txt` está pendiente de actualizar: incluye módulos de la biblioteca estándar, un nombre incorrecto (`datatime`) y OpenAI 0.28. El comando anterior instala las dependencias usadas por los scripts, incluida la versión del SDK necesaria para las llamadas con `client`.

## Configuración

Creá un archivo `.env` junto a los scripts. Guardá las credenciales reales únicamente en ese archivo: está excluido del repositorio mediante `.gitignore`. El README usa valores de ejemplo; los montos, fechas y plataformas de los comandos también son ficticios.

```dotenv
OPENAI_API_KEY=tu_clave_openai

API_KEY_POLYGON=tu_clave_polygon
TICKETS=AAPL,MSFT,ITUB

CRYPTO=BTC,ETH,SOL,BCH,BNB,USDT
CRYPTO_IDS=bitcoin,ethereum,solana,bitcoin-cash,binancecoin,tether

BINANCE_API_KEY=tu_clave_binance
BINANCE_API_SECRET=tu_secreto_binance

TELEGRAM_TOKEN=tu_token_telegram
DEFAULT_CHAT_ID=tu_chat_id
```

`TICKETS` define las acciones disponibles y genera sus argumentos de línea de comandos. `CRYPTO` hace lo mismo para las criptomonedas; `CRYPTO_IDS` debe contener sus identificadores de CoinGecko en el mismo orden y con la misma cantidad de elementos. Escribí los símbolos en mayúsculas, separados por comas y sin espacios.

Las claves de Binance se usan para consultar la cuenta con `--binance`. Las de Telegram se usan con `--telegram`. La clave de OpenAI debe estar configurada al iniciar los scripts; `crypto.py` también hace una consulta inicial aunque no se pidan análisis.

Ambos scripts cargan `.env` desde su propia carpeta. Los archivos Excel se leen y guardan en el directorio desde el que ejecutás el comando, por lo que conviene correrlos desde la carpeta del proyecto.

## Acciones

### Registrar compras

Los montos se indican en USD. La cantidad de acciones se calcula a partir del precio consultado para la fecha de compra:

```bash
python actions.py --aapl 100 --aapl_platform ejemplo --aapl_date 2026-10-23
```

Podés registrar varias compras en una ejecución. Cada símbolo configurado en `TICKETS` tiene los argumentos `--simbolo`, `--simbolo_platform` y `--simbolo_date`, escritos en minúsculas. La fecha es opcional y usa el día de ejecución por defecto.

Las compras se agregan a una hoja por acción. El resumen `Acciones invertidas` muestra cantidades, inversión, valor actual, precio y ganancia o pérdida.

### Actualizar cotizaciones y totales

```bash
python actions.py
```

Para consultar cotizaciones de una fecha determinada en el resumen:

```bash
python actions.py --date 2026-10-23
```

`--date` controla las cotizaciones del resumen; la fecha de cada compra se indica con `--simbolo_date`. Los totales se regeneran en cada ejecución; no hace falta pasar `--totales`.

Las consultas de precios a Polygon incluyen un límite de cinco llamadas por minuto. Si el plan no permite consultar el snapshot, se intenta obtener el cierre anterior. Para compras, si la consulta histórica devuelve un error 403, se usa la cotización disponible como alternativa.

### Registrar una venta total

```bash
python actions.py --sell AAPL --sell_total 125 --sell_date 2026-10-26
```

`--sell_total` es el importe total recibido en USD. El script registra la venta en la hoja `Ventas`, calcula el resultado respecto de toda la inversión acumulada y elimina la hoja de la acción vendida. Este comando cierra la posición completa. Si omitís `--sell_date`, se usa la fecha de ejecución.

### Consultas con OpenAI

- `--top10`: genera una lista de acciones con posible crecimiento.
- `--bajas`: genera una lista de acciones a bajo costo con posible crecimiento.

```bash
python actions.py --top10 --bajas
```

## Criptomonedas

### Registrar movimientos

Los importes se indican en USD; el script calcula las cantidades usando el precio de la fecha indicada:

```bash
python crypto.py --btc 100 --btc_wallet ejemplo --btc_date 2026-10-23 --eth 50 --eth_wallet ejemplo --eth_date 2026-10-23
```

Cada símbolo configurado en `CRYPTO` tiene los argumentos `--simbolo`, `--simbolo_wallet` y `--simbolo_date`. Si omitís la fecha, se usa la del día de ejecución.

Para registrar un retiro, usá un importe negativo:

```bash
python crypto.py --eth -40 --eth_wallet binance --eth_date 2026-10-23
```

Un cambio entre criptomonedas se registra como una salida y una entrada:

```bash
python crypto.py --usdt -100 --usdt_wallet binance --usdt_date 2026-10-23 --bnb 100 --bnb_wallet binance --bnb_date 2026-10-23
```

El argumento `--usdt_tax` permite aplicar un factor al importe de compra registrado; su valor por defecto es `1`. Se aplica a todos los movimientos de esa ejecución:

```bash
python crypto.py --usdt 100 --usdt_tax 1.05 --usdt_wallet binance
```

Cada cripto tiene su propia hoja de movimientos. La hoja `Totales` se regenera con la cantidad acumulada, inversión, valor actual, ganancia o pérdida y precio unitario. Los precios de cripto se registran con hasta cuatro decimales.

### Actualizar precios y consultar información adicional

```bash
python crypto.py
python crypto.py --tendencia --binance --top10 --bajas
```

- `--tendencia`: agrega históricos de los últimos 180 días y gráficos de evolución absoluta y normalizada.
- `--binance`: consulta saldos de la cuenta y los guarda en hojas adicionales.
- `--top10`: consulta criptomonedas que el modelo considera sobrevaloradas para una posible venta. La hoja conserva el nombre `Top 10 Cripto del momento`.
- `--bajas`: consulta criptomonedas que el modelo considera infravaloradas para una posible compra.

Las listas de ambos scripts son respuestas generadas por OpenAI; estas consultas no incorporan una búsqueda de noticias ni cotizaciones al modelo.

## Telegram

Configurá `TELEGRAM_TOKEN` y `DEFAULT_CHAT_ID` en `.env`. Para obtener el identificador del chat, enviá un mensaje al bot y consultá `https://api.telegram.org/bot<TU_TOKEN>/getUpdates`; el valor está en `message.chat.id`.

```bash
python crypto.py --telegram
python actions.py --telegram
```

Los mensajes incluyen tablas con cantidades, inversión, valor actual y resultados. En acciones, el resultado final del mensaje también incorpora las ganancias o pérdidas registradas en `Ventas`.

En `actions.py`, el envío lee el Excel guardado antes de actualizarlo. Para enviar un resumen actualizado, ejecutá primero `python actions.py` y después `python actions.py --telegram`.

## Explorar el Excel con Jupyter

```bash
python -m pip install jupyterlab
jupyter lab crypto.ipynb
```

Para cargar todas las hojas desde una celda:

```python
import pandas as pd

hojas = pd.read_excel("crypto.xlsx", sheet_name=None)
for nombre, datos in hojas.items():
    print(nombre)
    display(datos)
```

## Ejecución automática en macOS

Si usás un LaunchAgent, configurá en el `.plist` la ruta absoluta al Python del entorno virtual, al script y al directorio de trabajo (`WorkingDirectory`).

Para validar y recargar un agente de acciones ya creado:

```bash
plutil -lint ~/Library/LaunchAgents/com.example.actions.plist
launchctl bootout gui/$UID/com.example.actions
launchctl bootstrap gui/$UID ~/Library/LaunchAgents/com.example.actions.plist
launchctl kickstart -k gui/$UID/com.example.actions
tail -f ~/Library/Logs/actions.cron.log
```

El comando `bootout` puede devolver un error si el agente todavía no estaba cargado. Reemplazá `com.example.actions` por el identificador de tu agente; para cripto, por ejemplo, `com.example.crypto`. La ruta del log debe coincidir con la definida en el `.plist`.
