from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests


# ---------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------

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


COLUMNAS_SELECCIONADAS = [
    "Trading day",
    "Product",
    "Place of delivery",
    "Area",
    "First Day Delivery",
    "Last Day Delivery",
    "Reference Price\n[EUR/MWh]",
    "Auction Price\n[EUR/MWh]",
    "Last Price\n[EUR/MWh]",
    "Bid\n[EUR/MWh]",
    "Ask\n[EUR/MWh]",
    "Source",
    "Maximum Price\n[EUR/MWh]",
    "Minimum Price\n[EUR/MWh]",
    "Price difference between purchases and sales\n[%]",
    "Auction Volume Traded®[MWh]",
    "OTC Volume Registered [MWh]",
    "Volume Traded\n[MWh]",
]


COLUMNAS_FECHA = [
    "Trading day",
    "First Day Delivery",
    "Last Day Delivery",
]


COLUMNAS_NUMERICAS = [
    "Reference Price\n[EUR/MWh]",
    "Auction Price\n[EUR/MWh]",
    "Last Price\n[EUR/MWh]",
    "Bid\n[EUR/MWh]",
    "Ask\n[EUR/MWh]",
    "Maximum Price\n[EUR/MWh]",
    "Minimum Price\n[EUR/MWh]",
    "Price difference between purchases and sales\n[%]",
    "Auction Volume Traded®[MWh]",
    "OTC Volume Registered [MWh]",
    "Volume Traded\n[MWh]",
]


def descargar_excel(url):
    """Descarga el fichero Excel anual publicado por MIBGAS."""

    print(f"Descargando fichero desde:\n{url}")

    encabezados = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(compatible; MIBGAS-Historico/1.0)"
        )
    }

    respuesta = requests.get(
        url,
        headers=encabezados,
        timeout=120,
    )

    respuesta.raise_for_status()

    if not respuesta.content:
        raise ValueError(
            "El fichero descargado esta vacio."
        )

    print(
        "Descarga completada:",
        f"{len(respuesta.content):,} bytes",
    )

    return BytesIO(respuesta.content)


def limpiar_datos(contenido_excel):
    """Reproduce las transformaciones realizadas en Power Query."""

    datos = pd.read_excel(
        contenido_excel,
        sheet_name=NOMBRE_HOJA,
        engine="openpyxl",
    )

    # Normalizar los encabezados sin eliminar los saltos de linea.
    datos.columns = [
        str(columna).strip()
        for columna in datos.columns
    ]

    columnas_no_encontradas = [
        columna
        for columna in COLUMNAS_SELECCIONADAS
        if columna not in datos.columns
    ]

    if columnas_no_encontradas:
        print("Columnas encontradas en el Excel:")

        for columna in datos.columns:
            print(repr(columna))

        raise ValueError(
            "No se han encontrado las siguientes columnas: "
            + ", ".join(columnas_no_encontradas)
        )

    # Mantener solamente las columnas empleadas en Power BI.
    datos = datos[COLUMNAS_SELECCIONADAS].copy()

    # Eliminar filas completamente vacias.
    datos = datos.dropna(how="all")

    # Transformar las fechas.
    for columna in COLUMNAS_FECHA:
        datos[columna] = pd.to_datetime(
            datos[columna],
            errors="coerce",
        )

    # Transformar precios y volumenes.
    for columna in COLUMNAS_NUMERICAS:
        datos[columna] = pd.to_numeric(
            datos[columna],
            errors="coerce",
        )

    # Eliminar posibles filas sin fecha de negociacion.
    datos = datos.dropna(subset=["Trading day"])

    # Eliminar duplicados completos.
    datos = datos.drop_duplicates()

    # Ordenar los registros.
    datos = datos.sort_values(
        by=[
            "Trading day",
            "Product",
            "Place of delivery",
            "First Day Delivery",
        ],
        na_position="last",
    )

    return datos


def guardar_csv(datos, fichero_salida):
    """Guarda el fichero de salida para Power BI."""

    fichero_salida.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    datos.to_csv(
        fichero_salida,
        index=False,
        encoding="utf-8-sig",
        date_format="%Y-%m-%d",
    )

    print(f"Fichero guardado: {fichero_salida}")
    print(f"Numero de registros: {len(datos):,}")
    print(
        "Periodo:",
        datos["Trading day"].min().date(),
        "-",
        datos["Trading day"].max().date(),
    )


def main():
    contenido_excel = descargar_excel(URL_MIBGAS)
    datos_mibgas = limpiar_datos(contenido_excel)
    guardar_csv(datos_mibgas, FICHERO_SALIDA)


if __name__ == "__main__":
    main()
