"""Genera 12 Excel de ejemplo (uno por mes) con columnas fecha, producto, canal, monto.

Son datos ficticios solo para probar el Consolidador. Reemplazalos por tus
archivos reales en datos_mensuales/ (mismo nombre de columnas) y volvé a
correr consolidador.py.
"""
import random
from calendar import monthrange
from datetime import date
from pathlib import Path

import pandas as pd

CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "datos_mensuales"
PRODUCTOS = ["Camisa", "Pantalón", "Chaqueta", "Zapatos", "Gorra"]
CANALES = ["Tienda física", "Online", "Mayorista"]
ANIO = 2026

random.seed(7)

CARPETA_SALIDA.mkdir(exist_ok=True)

for mes in range(1, 13):
    dias_del_mes = monthrange(ANIO, mes)[1]
    n_filas = random.randint(25, 40)
    filas = []
    for _ in range(n_filas):
        dia = random.randint(1, dias_del_mes)
        filas.append(
            {
                "fecha": date(ANIO, mes, dia),
                "producto": random.choice(PRODUCTOS),
                "canal": random.choice(CANALES),
                "monto": round(random.uniform(150, 2500), 2),
            }
        )
    df = pd.DataFrame(filas).sort_values("fecha")
    nombre_archivo = CARPETA_SALIDA / f"ventas_{ANIO}-{mes:02d}.xlsx"
    df.to_excel(nombre_archivo, index=False)
    print(f"Generado {nombre_archivo.name} ({n_filas} filas)")
