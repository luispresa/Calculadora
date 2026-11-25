# Consolidador de ventas mensuales

Herramienta hecha con Claude Code para el Ejercicio 7 (Sesión 3, "Agentes de AI con Claude", UFM).
Junta 12 Excel mensuales (uno por mes, columnas `fecha`, `producto`, `canal`, `monto`) en un solo
archivo con una hoja por mes, un resumen por mes y por canal, y una gráfica del total mensual.

## Carpetas

- `datos_mensuales/`: ahí van tus archivos `.xlsx`, uno por mes. Ahora mismo tiene **datos de
  ejemplo generados al azar** (no son ventas reales) porque no había archivos tuyos disponibles.
  Borralos y poné los tuyos cuando los tengas.
- `salida/consolidado.xlsx`: el resultado.
- `scripts/generar_datos_ejemplo.py`: el script que generó los datos de ejemplo (no hace falta
  volver a correrlo salvo que quieras regenerarlos).
- `consolidador.py`: la herramienta.

## Cómo volver a correrla el mes que viene

1. Agregá el Excel del nuevo mes a la carpeta `datos_mensuales/` (mismas 4 columnas, cualquier
   nombre de archivo).
2. Abrí una terminal en esta carpeta del proyecto.
3. Corré `python3 consolidador.py`. Usa `datos_mensuales/` como entrada y escribe
   `salida/consolidado.xlsx` por defecto (podés pasarle otra carpeta/archivo como argumentos).
4. Revisá el mensaje en pantalla: si algún archivo tiene columnas distintas, te avisa y lo salta
   en vez de inventar datos — corregí ese archivo y volvé a correr.
5. Abrí `salida/consolidado.xlsx`: una hoja por mes, la hoja `Resumen` con el total por mes y por
   canal, y la gráfica de barras del total mensual.

## Verificación hecha

Para enero 2026: suma manual de la columna `monto` en el archivo de origen = **44,583.71**,
igual al total que aparece en la hoja `Resumen` para `2026-01`. También se probó que un archivo
con columnas distintas se salta con un aviso, sin inventar datos.
