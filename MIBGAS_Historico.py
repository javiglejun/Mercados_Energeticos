from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests


# =========================================================
# CONFIGURACION
# =========================================================

ANIO_ACTUAL = datetime.now().year

URL_MIBGAS = (
    f"https://www.mibgas.es/es/file-access/"
    f"MIBGAS_Data_{ANIO_ACTUAL}.xlsx"
    f"?path=AGNO_{ANIO_ACTUAL}/XLS"
)

NOMBRE_HOJA = "Trading Data PVB&VTP"

# La carpeta base será siempre aquella donde está este script.
CARPETA_REPOSITORIO = Path(__file__).resolve().parent

CARPETA_DATOS = CARPETA_REPOSITORIO / "data"

FICHERO_SALIDA = (
    CARPETA_DATOS / f"MIBGAS_{ANIO_ACTUAL}.csv"
)


NOMBRES_COLUMNAS = [
    "Trading day",
    "Product",
    "Place of delivery",
    "Area",
    "First Day Delivery",
    "Last Day Delivery",
    "Reference Price [EUR/MWh]",
    "Auction Price [EUR/MWh]",
    "Last Price [EUR/MWh]",
    "Bid [EUR/MWh]",
    "Ask [EUR/MWh]",
    "Source",
    "Maximum Price [EUR/MWh]",
    "Minimum Price [EUR/MWh]",
    "Price difference between purchases and sales [%]",
    "Auction Volume Traded [MWh]",
    "OTC Volume Registered [MWh]",
    "Volume Traded [MWh]",
]


COLUMNAS_FECHA = [
    "Trading day",
    "First Day Delivery",
    "Last Day Delivery",
]


COLUMNAS_NUMERICAS = [
    "Reference Price [EUR/MWh]",
    "Auction Price [EUR/MWh]",
    "Last Price [EUR/MWh]",
    "Bid [EUR/MWh]",
    "Ask [EUR/MWh]",
    "Maximum Price [EUR/MWh]",
    "Minimum Price [EUR/MWh]",
    "Price difference between purchases and sales [%]",
    "Auction Volume Traded [MWh]",
    "OTC Volume Registered [MWh]",
    "Volume Traded [MWh]",
]


# =========================================================
# DESCARGA
# =========================================================

def descargar_excel():
    print("=" * 60)
    print(f"ACTUALIZACION DE MIBGAS {ANIO_ACTUAL}")
    print("=" * 60)
    print(f"Descargando: {URL_MIBGAS}")

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "*/*",
    }

    respuesta = requests.get(
        URL_MIBGAS,
        headers=headers,
        timeout=180,
    )

    respuesta.raise_for_status()

    if not respuesta.content:
        raise ValueError(
            "El fichero descargado está vacío."
        )

    if not respuesta.content.startswith(b"PK"):
        raise ValueError(
            "El contenido descargado no parece un fichero XLSX válido."
        )

    print(
        f"Descarga completada: "
        f"{len(respuesta.content):,} bytes"
    )

    return BytesIO(respuesta.content)


# =========================================================
# TRANSFORMACION
# =========================================================

def transformar_datos(contenido_excel):
    print()
    print("Leyendo hoja:", NOMBRE_HOJA)
    print("Leyendo las columnas A:R por posición.")

    datos = pd.read_excel(
        contenido_excel,
        sheet_name=NOMBRE_HOJA,
        usecols="A:R",
        header=0,
        engine="openpyxl",
    )

    print(f"Filas leídas inicialmente: {len(datos):,}")
    print(f"Columnas leídas: {datos.shape[1]}")

    if datos.shape[1] != 18:
        raise ValueError(
            "Se esperaban 18 columnas entre A y R, "
            f"pero se han leído {datos.shape[1]}."
        )

    # Sustituir los encabezados problemáticos de MIBGAS.
    datos.columns = NOMBRES_COLUMNAS

    # Eliminar filas completamente vacías.
    datos = datos.dropna(how="all")

    # Convertir fechas.
    for columna in COLUMNAS_FECHA:
        datos[columna] = pd.to_datetime(
            datos[columna],
            errors="coerce",
        )

    # Convertir precios y volúmenes.
    for columna in COLUMNAS_NUMERICAS:
        datos[columna] = pd.to_numeric(
            datos[columna],
            errors="coerce",
        )

    # Limpiar campos de texto sin convertir valores vacíos en texto.
    columnas_texto = [
        "Product",
        "Place of delivery",
        "Area",
        "Source",
    ]

    for columna in columnas_texto:
        datos[columna] = (
            datos[columna]
            .astype("string")
            .str.strip()
        )

    # Mantener únicamente registros con fecha válida.
    datos = datos.dropna(
        subset=["Trading day"]
    )

    # Eliminar duplicados.
    datos = datos.drop_duplicates()

    # Ordenar.
    datos = datos.sort_values(
        by=[
            "Trading day",
            "Product",
            "Place of delivery",
            "First Day Delivery",
            "Last Day Delivery",
        ],
        na_position="last",
    )

    datos = datos.reset_index(drop=True)

    if datos.empty:
        raise ValueError(
            "No quedan registros después de transformar los datos."
        )

    print(f"Registros )
