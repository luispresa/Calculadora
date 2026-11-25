"""Consolidador de ventas mensuales.

Lee todos los Excel de una carpeta (uno por mes, columnas: fecha, producto,
canal, monto) y arma un solo Excel con:
  - una hoja por mes (datos tal cual, ordenados por fecha)
  - una hoja "Resumen" con el total por mes y por canal
  - una gráfica de barras del total mensual

Uso:
    python3 consolidador.py [carpeta_entrada] [archivo_salida.xlsx]

Por defecto usa datos_mensuales/ como entrada y
salida/consolidado.xlsx como salida.
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

COLUMNAS_ESPERADAS = {"fecha", "producto", "canal", "monto"}


def nombre_hoja(archivo: Path, usados: set[str]) -> str:
    base = archivo.stem[:31] or "hoja"
    nombre = base
    i = 2
    while nombre in usados:
        sufijo = f"_{i}"
        nombre = base[: 31 - len(sufijo)] + sufijo
        i += 1
    usados.add(nombre)
    return nombre


def cargar_archivos(carpeta_entrada: Path) -> dict[str, pd.DataFrame]:
    archivos = sorted(carpeta_entrada.glob("*.xlsx"))
    if not archivos:
        sys.exit(f"No encontré archivos .xlsx en {carpeta_entrada}")

    hojas: dict[str, pd.DataFrame] = {}
    usados: set[str] = set()
    for archivo in archivos:
        df = pd.read_excel(archivo)
        columnas = {c.strip().lower() for c in df.columns}
        if columnas != COLUMNAS_ESPERADAS:
            print(
                f"AVISO: {archivo.name} tiene columnas distintas a las esperadas "
                f"{sorted(COLUMNAS_ESPERADAS)} -> encontré {sorted(columnas)}. "
                "Lo salto; no invento datos."
            )
            continue
        df.columns = [c.strip().lower() for c in df.columns]
        df["fecha"] = pd.to_datetime(df["fecha"])
        df = df.sort_values("fecha")
        hojas[nombre_hoja(archivo, usados)] = df
    if not hojas:
        sys.exit("Ningún archivo tenía las columnas esperadas. No se generó salida.")
    return hojas


def armar_resumen(hojas: dict[str, pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame]:
    completo = pd.concat(hojas.values(), ignore_index=True)
    completo["mes"] = completo["fecha"].dt.to_period("M").astype(str)

    por_mes = completo.groupby("mes")["monto"].sum().reset_index()
    por_mes.columns = ["mes", "total"]

    por_mes_canal = (
        completo.groupby(["mes", "canal"])["monto"]
        .sum()
        .unstack("canal")
        .fillna(0)
        .reset_index()
    )
    return por_mes, por_mes_canal


def escribir_excel(
    hojas: dict[str, pd.DataFrame],
    por_mes: pd.DataFrame,
    por_mes_canal: pd.DataFrame,
    archivo_salida: Path,
) -> None:
    archivo_salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        for nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre, index=False)

        fila_total_por_canal = len(por_mes) + 4
        por_mes.to_excel(writer, sheet_name="Resumen", index=False, startrow=0)
        por_mes_canal.to_excel(
            writer, sheet_name="Resumen", index=False, startrow=fila_total_por_canal
        )

    _agregar_grafica(archivo_salida, len(por_mes))


def _agregar_grafica(archivo_salida: Path, n_meses: int) -> None:
    from openpyxl import load_workbook

    libro = load_workbook(archivo_salida)
    resumen = libro["Resumen"]

    grafica = BarChart()
    grafica.title = "Total mensual"
    grafica.y_axis.title = "Monto"
    grafica.x_axis.title = "Mes"

    datos = Reference(resumen, min_col=2, min_row=1, max_row=n_meses + 1)
    categorias = Reference(resumen, min_col=1, min_row=2, max_row=n_meses + 1)
    grafica.add_data(datos, titles_from_data=True)
    grafica.set_categories(categorias)

    columna_ancla = get_column_letter(5)
    resumen.add_chart(grafica, f"{columna_ancla}2")
    libro.save(archivo_salida)


def main() -> None:
    carpeta_entrada = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("datos_mensuales")
    archivo_salida = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("salida/consolidado.xlsx")

    hojas = cargar_archivos(carpeta_entrada)
    por_mes, por_mes_canal = armar_resumen(hojas)
    escribir_excel(hojas, por_mes, por_mes_canal, archivo_salida)

    print(f"\nConsolidé {len(hojas)} meses en {archivo_salida}")
    print("\nTotal por mes:")
    print(por_mes.to_string(index=False))


if __name__ == "__main__":
    main()
