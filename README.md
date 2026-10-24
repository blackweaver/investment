## 📝 Instrucciones:

### Correr para recibir en el móvil

Entorno de python: source ~/.venvs/crypto39/bin/activate
python crypto.py --top10 --tendencia --telegram --bajas
python actions.py --top10 --telegram --bajas

Renombralo como .env (sin “example”).

Abrilo y reemplazá la línea con tu clave real de OpenAI:

OPENAI_API_KEY=sk-tu_clave_aquí
Ahora voy a generarte el script actualizar_registro_cripto.py que:

✅ Carga tu API Key desde .env
✅ Usa openai para consultar las 3 criptos más prometedoras según analistas
✅ Actualiza el Excel agregando una fila por cripto con precio y notas

Voy con eso ahora.

Listo, Ricardo. Acá tenés el script final para automatizar tu registro cripto con integración a CoinGecko y OpenAI:

## 📜 Descargar actualizar_registro_cripto.py

🔑 Descargar .env.example

## 🔧 ¿Cómo usarlo?

Cocoa tu Excel Registro_Cripto_Historico.xlsx en la misma carpeta que el script.

Renombrá y editá el archivo .env.example como .env con tu clave de OpenAI:

OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxx
Instalá dependencias si no las tenés:

pip install openai pandas openpyxl python-dotenv requests
Ejecutá el script:

python actualizar_registro_cripto.py

## ✅ ¿Qué hace este script?

Actualiza precios de BTC, ETH y SOL

Crea (si no existe) o actualiza una hoja por cripto

Agrega una nueva fila con:

Fecha actual

Precio simulado de compra (5% abajo)

Precio actual

Ganancia calculada

Y en la columna "Notas", el resultado del análisis de ChatGPT

¿Querés que más adelante agreguemos:

Análisis de tendencia con TradingView?

O que se conecte con tu billetera real (por ejemplo vía etherscan o Binance API)?

Estoy listo cuando lo necesites.

## Empezar proyecto en python

1. Crear requirements.txt
2. Lista de requerimientos:
   openai pandas openpyxl python-dotenv requests
3. Creo el entorno: python -m venv env
4. Lo activo: source env/bin/activate (source ~/.venvs/crypto39/bin/activate)
5. Instalo las librerías: pip install -r ./requirements.txt (-r significa leer desde...)
6. Actualizar pip: python -m pip install --upgrade pip
7. Mostrar los print en consola: python ./01_tratamiento_datos.py

## Activar entorno por fuera de GoogleDrive

source ~/.venvs/crypto39/bin/activate

## OpenIA

Cómo saber si tengo acceso a GPT-4

`curl https://api.openai.com/v1/models \  -H "Authorization: Bearer TU_API_KEY"`

## Ejemplo de ejecución completa con todos los parámetros desde la línea de comandos

```
python crypto.py --bch 8 --bch_wallet bitcoin.com --bch_date 2023-05-05 --btc 250 --btc_wallet bitcoin.com --btc_date 2022-08-20 --eth 100 --eth_wallet bitcoin.com --eth_date 2024-12-31 --tendencia --binance --top10

python crypto.py --btc 310 --btc_wallet prex --btc_date 2025-07-15 --eth 110 --eth_wallet prex --eth_date 2025-07-15 --sol 80 --sol_wallet prex --sol_date 2025-07-15 --tendencia --binance --top10

python crypto.py --usdt 100 --usdt_tax 1.05 --usdt_wallet binance --usdt_date 2025-07-26 --eth 300 --eth_wallet binance --eth_date 2025-07-26 --sol 80 --sol_wallet binance --sol_date 2025-07-26

python crypto.py --ada 592 --ada_wallet binance --ada_date 2026-03-11 --usdt -155 --usdt_wallet binance --usdt_date 2026-03-11

Retiros

python crypto.py --eth -30 --eth_wallet prex
python crypto.py --eth -40 --eth_wallet binance
python crypto.py --btc -300 --btc_wallet bitcoin.com --btc_date 2025-07-15

Cambio (Retiro + Agregado)
python crypto.py --usdt -100 --usdt_wallet binance --usdt_date 2025-08-10 --bnb 100 --bnb_wallet binance --bnb_date 2025-08-10

Enviando mensaje

python crypto.py --telegram

Para acciones

python actions.py --fig 237 --fig_platform eToro --fig_date 2025-08-14
python actions.py --xiacf 158 --xiacf_platform eToro --xiacf_date 2025-08-14
python actions.py --itub 101 --itub_platform eToro --itub_date 2025-08-18
python actions.py --mnst 101 --mnst_platform eToro --mnst_date 2025-08-18
python actions.py --amzn 234 --amzn_platform eToro --amzn_date 2025-09-16
python actions.py --plug 49 --plug_platform eToro --plug_date 2025-09-16

```

## Jypyter

Jupyter es un entorno interactivo donde podés escribir y ejecutar código (Python y otros lenguajes) celda por celda, y ver los resultados de inmediato, incluyendo tablas, gráficos, texto formateado, imágenes, etc..

### ¿Por qué se asocia tanto con “ver datos”?

Porque en ciencia de datos y análisis, es muy cómodo para cargar un Excel, una base de datos o una API, y mostrar los resultados en tablas (pandas) o gráficos (matplotlib, plotly).

Renderiza automáticamente los DataFrames como tablas con scroll, ordenamiento y estilos (en VS Code y JupyterLab es aún más visual)

Permite combinar código + visualización + notas explicativas en un mismo documento.

### Pero va mucho más allá - Con Jupyter también podés:

Ejecutar y depurar funciones de tu aplicación (no solo mostrar datos).

Probar consultas a APIs.

Automatizar reportes.

Hacer prototipado rápido de ideas y algoritmos.

Documentar paso a paso un flujo de trabajo.

Integrar texto (Markdown) con resultados en el mismo archivo.

````
# Le dice a Jupyter: "cada vez que ejecutes una celda, recargá todos los módulos que importaste".

%load_ext autoreload
%autoreload 2

dfs = pd.read_excel("crypto.xlsx", sheet_name=None)  # None = todas las hojas

# dfs es un diccionario: clave = nombre de hoja, valor = DataFrame
for nombre, df in dfs.items():
    print(f"--- {nombre} ---")
    display(df)  # en Jupyter muestra tabla
```

Ejectura archivo Jupyter:
```
jupyter lab crypto.ipynb
jupyter nbconvert --to script crypto.ipynb
```

## Ejecutar script automático con un plist

Veo si está ok

```
plutil -lint ~/Library/LaunchAgents/com.ricardo.crypto.plist
```

Bajo el anterior

```
launchctl bootout gui/$UID/com.ricardo.crypto 2>/dev/null || true
```

Lo cargo nuevamente

```
launchctl bootstrap gui/$UID ~/Library/LaunchAgents/com.ricardo.actions.plist
launchctl kickstart -k gui/$UID/com.ricardo.actions
```

Veo los logs en tiempo real

```
tail -f /Users/Ricardo_/Library/Logs/actions.cron.log
```

### Resumen

```
➜  ~ launchctl bootout gui/$UID/com.ricardo.crypto 2>/dev/null || true
➜  ~ launchctl bootstrap gui/$UID ~/Library/LaunchAgents/com.ricardo.crypto.plist
➜  ~ launchctl kickstart -k gui/$UID/com.ricardo.crypto
➜  ~ tail -f /Users/Ricardo_/Library/Logs/crypto.cron.log
```

## Bot automático de Telegram

TELEGRAM_TOKEN=xxxxxxxxxxxxxxxxxxxx
DEFAULT_CHAT_ID=xxxxxxxxxx

https://api.telegram.org/botxxxxxxxxxxxxx/getUpdates

Obtengo:

```
{
ok: true,
result: [
{
update_id: 187857952,
message: {
message_id: 3,
from: {
id: 5177802022,
is_bot: false,
first_name: "Ricardo",
language_code: "es"
},
chat: {
id: 5177802022,
first_name: "Ricardo",
type: "private"
},
date: 1755710616,
text: "Hola, cómo estás?"
}
}
]
}
```
````
