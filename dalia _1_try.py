import pandas as pd
from datetime import datetime, timedelta
import requests
import pytz
from io import BytesIO

# -----------------------
# CONFIGURACIÓN
# -----------------------
url_excel = "https://tecmx-my.sharepoint.com/:x:/g/personal/a01412804_tec_mx/Ebyfbml2HTNIqufa6e7aEIMB3F7GSOe_SNjbhvU3jH2a8g?download=1"
hoja_excel = "VENTAS"

token_telegram = "8466924514:AAFECiYb46oWmL5k123ikGaj7_pMZYeKYOw"
chat_id_telegram = "5578633101"

# -----------------------
# DESCARGAR ARCHIVO DESDE ONEDRIVE
# -----------------------
print("Descargando archivo desde OneDrive...")
response = requests.get(url_excel)
if response.status_code != 200:
    print("Error al descargar el archivo:", response.status_code)
    exit()
excel_bytes = BytesIO(response.content)

# -----------------------
# LEER EXCEL
# -----------------------
df = pd.read_excel(excel_bytes, sheet_name=hoja_excel, engine="openpyxl", header=1)
print("Columnas detectadas:", list(df.columns))

# -----------------------
# IDENTIFICAR COLUMNAS
# -----------------------
col_fecha = [c for c in df.columns if 'FECHA' in str(c).upper()][0]
col_cliente = [c for c in df.columns if 'CLIENTE' in str(c).upper()][0]
col_precio = [c for c in df.columns if 'PRECIO TOTAL' in str(c).upper()][0]

col_desc = [c for c in df.columns if 'DESCRIPCIÓN' in str(c).upper()][0]
col_anticipo = [c for c in df.columns if 'ANTICIPO' in str(c).upper()][0]
col_restante = [c for c in df.columns if 'RESTANTE' in str(c).upper()][0]

# -----------------------
# LIMPIAR CLIENTE
# -----------------------
df[col_cliente] = df[col_cliente].astype(str).str.strip()
df[col_cliente] = df[col_cliente].replace(["", "nan", "NaN", "None"], pd.NA)
df[col_cliente] = df[col_cliente].fillna("")

# -----------------------
# CONVERTIR FECHAS
# -----------------------
def parse_fecha(x):
    if pd.isna(x) or x == '':
        return pd.NaT
    if isinstance(x, datetime):
        return x
    if isinstance(x, (int, float)):  # Excel date
        return datetime(1899, 12, 30) + timedelta(days=x)
    try:
        return pd.to_datetime(str(x), dayfirst=True, errors='coerce')
    except:
        return pd.NaT

df[col_fecha] = df[col_fecha].apply(parse_fecha)

# -----------------------
# FILTRAR PEDIDOS DE HOY
# -----------------------
hoy = datetime.now(pytz.timezone("America/Mexico_City")).date()
pedidos_hoy = df[df[col_fecha].dt.date == hoy]

# -----------------------
# FORMAR MENSAJE
# -----------------------
if pedidos_hoy.empty:
    mensaje_final = "El día de hoy no hay pedidos"
else:
    lista_msgs = ["Pedido(s) para HOY:\n"]
    for _, row in pedidos_hoy.iterrows():

        cliente = str(row[col_cliente]).strip()
        precio = row[col_precio]
        descripcion = str(row[col_desc]).strip()

        # Limpiar descripción
        if descripcion.upper() in ["", "NAN", "NONE"]:
            descripcion = ""

        anticipo = row[col_anticipo]
        restante = row[col_restante]

        # Línea principal
        msg = f"{cliente} - ${precio}"

        # Descripción si existe
        if descripcion:
            msg += f"\n{descripcion}"

        # Anticipo / restante si existen
        if not pd.isna(anticipo) or not pd.isna(restante):
            anticipo_txt = f"{anticipo}" if not pd.isna(anticipo) else "0"
            restante_txt = f"{restante}" if not pd.isna(restante) else "0"
            msg += f"\nEl anticipo fue de ${anticipo_txt} y el restante es ${restante_txt}"

        lista_msgs.append(msg + "\n")

    mensaje_final = "\n".join(lista_msgs)

# -----------------------
# ENVIAR A TELEGRAM
# -----------------------
url = f"https://api.telegram.org/bot{token_telegram}/sendMessage"
response = requests.post(url, data={"chat_id": chat_id_telegram, "text": mensaje_final})
print("Mensaje enviado:", response.json())


