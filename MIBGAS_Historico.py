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

CARPETA_DATOS = Path("data")

FICHERO_SALIDA = (
    CARPETA_DATOS / f"MIBGAS_{ANIO_ACTUAL}.csv"
)


# Nombres limpios que asignaremos a las columnas A:R.
# No se utilizan los encabezados originales del Excel.

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
# DESCARGA DEL EXCEL
# =========================================================

def descargar_excel(url):
    """
    Descarga el fichero Excel anual publicado por MIBGAS.
    """

    print("=" * 60)
    print("DESCARGA DEL FICHERO MIBGAS")
    print("=" * 60)
    print(f"URL: {url}")

    encabezados = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(X11; Linux x86_64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        ),
        "Accept": (
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet,"
            "application/vnd.ms-excel,"
            "*/*"
        ),
    }

    respuesta = requests.get(
        url,
        headers=encabezados,
        timeout=180,
    )

    respuesta.raise_for_status()

    if not respuesta.content:
        raise ValueError(
            "El fichero descargado esta vacio."
        )

    # Un fichero XLSX es internamente un fichero ZIP y comienza por PK.
    if not respuesta.content.startswith(b"PK"):
        tipo_contenido = respuesta.headers.get(
            "Content-Type",
            "desconocido",
        )

        raise ValueError(
            "La descarga no parece ser un fichero XLSX valido. "
            f"Content-Type recibido: {tipo_contenido}"
        )

    print(
        f"Descarga completada: "
        f"{len(respuesta.content):,} bytes"
    )

    return BytesIO(respuesta.content)


# =========================================================
# LECTURA Y LIMPIEZA
# =========================================================

def limpiar_datos(contenido_excel):
    """
    Lee las columnas A:R de la hoja de MIBGAS.

    Las columnas se leen por posicion y posteriormente se les
    asignan nombres controlados. De esta manera, el script no
    depende de saltos de linea, simbolos especiales o caracteres
    ocultos presentes en los encabezados originales.
    """

    print()
    print("=" * 60)
    print("LECTURA Y LIMPIEZA DE LOS DATOS")
    print("=" * 60)
    print(f"Hoja seleccionada: {NOMBRE_HOJA}")
    print("Columnas seleccionadas: A:R")

    datos = pd.read_excel(
        contenido_excel,
        sheet_name=NOMBRE_HOJA,
        usecols="A:R",
        header=0,
        engine="openpyxl",
    )

    print(
        f"Dimensiones iniciales: "
        f"{datos.shape,} filas y "
        f"{datos.shape[1]} columnas"
    )

    # Comprobar que efectivamente se han obtenido 18 columnas.
    if datos.shape[1] != len(NOMBRES_COLUMNAS):
        raise ValueError(
            "Numero inesperado de columnas. "
            f"Se esperaban {len(NOMBRES_COLUMNAS)} columnas "
            f"y se han encontrado {datos.shape[1]}."
        )

    # Reemplazar todos los encabezados originales.
    datos.columns = NOMBRES_COLUMNAS

    print("Encabezados originales sustituidos correctamente.")

    # Eliminar filas completamente vacias.
    filas_antes = len(datos)

    datos = datos.dropna(how="all")

    print(
        "Filas completamente vacias eliminadas: "
        f"{filas_antes - len(datos):,}"
    )

    # Convertir las columnas de fecha.
    for columna in COLUMNAS_FECHA:
        datos[columna] = pd.to_datetime(
            datos[columna],
            errors="coerce",
        )

    # Convertir las columnas numericas.
    for columna in COLUMNAS_NUMERICAS:
        datos[columna] = pd.to_numeric(
            datos[columna],
            errors="coerce",
        )

    # Convertir las columnas de texto.
    columnas_texto = [
        "Product",
        "Place of delivery",
        "Area",
        "Source",
    ]

    for columna in columnas_texto:
        datos[columna] = datos[columna].astype("string").str.strip()

    # Eliminar filas sin fecha de negociacion valida.
    filas_antes = len(datos)

    datos = datos.dropna(
        subset=["Trading day"]
    )

    print(
        "Filas sin Trading day valido eliminadas: "
        f"{filas_antes - len(datos):,}"
    )

    # Eliminar duplicados completos.
    filas_antes = len(datos)

    datos = datos.drop_duplicates()

    print(
        "Filas duplicadas eliminadas: "
        f"{filas_antes - len(datos):,}"
    )

    # Ordenar los registros.
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
            "El resultado esta vacio despues de limpiar los datos."
        )

    fecha_minima = datos["Trading day"].min()
    fecha_maxima = datos["Trading day"].max()

    print()
    print("Lectura completada correctamente.")
    print(f"Registros validos: {len(datos):,}")
    print(
        "Periodo disponible: "
        f"{fecha_minima.strftime('%Y-%m-%d')} a "
        f"{fecha_maxima.strftime('%Y-%m-%d')}"
    )

    print()
    print("Columnas finales:")

    for numero, columna in enumerate(
        datos.columns,
        start=1,
    ):
        print(f"{numero:02d}. {columna}")

    return datos


# =========================================================
# VALIDACION
# =========================================================

def validar_datos(datos):
    """
    Realiza comprobaciones basicas antes de guardar el CSV.
    """

    print()
    print("=" * 60)
    print("VALIDACION DEL RESULTADO")
    print("=" * 60)

    if datos.empty:
        raise ValueError(
            "No hay datos para guardar.")
      
