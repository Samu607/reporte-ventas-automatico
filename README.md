# Reporte de ventas automático

Convierte los archivos de ventas desordenados de un negocio (CSV y Excel, uno por sucursal, cada uno con su propio formato) en un **reporte de Excel listo para presentar**, en segundos y con un solo comando.

## El problema que resuelve

En muchos negocios alguien pasa horas cada semana o cada mes juntando a mano los archivos de cada sucursal o vendedor. Las columnas se llaman distinto, las fechas vienen en varios formatos, los precios están escritos como texto (`$18.500`, `18,500.00`), hay filas repetidas y errores de digitación. Este script hace todo ese trabajo solo.

## Qué hace

1. **Lee todos los archivos** `.csv` y `.xlsx` de una carpeta. En los CSV detecta solo si el separador es coma o punto y coma.
2. **Reconoce las columnas** aunque tengan nombres distintos, por ejemplo "Fecha" o "FECHA VENTA", y "Asesor" o "Vendedor".
3. **Limpia los datos:**
   - fechas como `14/07/2026`, `2026-07-14` o `14-jul-2026`
   - precios como `$18.500`, `18500`, `18,500.00` o `18.500,50`
   - nombres de producto con mayúsculas o espacios inconsistentes
   - filas vacías y duplicadas
4. **Genera un Excel con 5 hojas:**
   - **Resumen:** ventas totales, transacciones, unidades, ticket promedio, gráfico de ventas por mes y participación por sucursal
   - **Productos** y **Vendedores:** ranking con unidades, ventas y % del total, con su gráfico
   - **Datos limpios:** todos los registros unificados, con filtros
   - **Problemas:** qué filas no se incluyeron, en qué archivo y fila están, y por qué, para que el negocio las corrija

## Cómo usarlo

```bash
pip install -r requirements.txt

# 1. (Opcional) Crear datos de ejemplo desordenados
python generar_datos_ejemplo.py

# 2. Generar el reporte
python reporte.py --entrada datos_ejemplo --salida reporte_ventas.xlsx
```

Salida de ejemplo:

```
✔ Reporte generado: reporte_ventas.xlsx
  2 archivos · 284 filas leídas · 259 válidas
  6 vacías · 13 duplicadas · 6 con errores
  Ventas totales: $26.008.500
```

## Adaptarlo a un cliente nuevo

Casi siempre basta con agregar los nombres de columna del cliente en el diccionario `ALIAS_COLUMNAS`, al inicio de `reporte.py`. Si falta una columna, el script dice exactamente cuál y qué columnas encontró en el archivo.

Extensiones típicas que se pueden ofrecer:
- Envío automático del reporte por correo cada lunes (Programador de tareas de Windows o cron)
- Lectura directa desde Google Sheets o desde la exportación del software de facturación
- Comparación contra el mes anterior y alertas cuando un producto baja de cierto nivel

## Tecnologías

Python 3.10+, pandas y openpyxl (gráficos nativos de Excel, sin depender de imágenes).
