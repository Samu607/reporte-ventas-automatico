# Reporte de ventas automático

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Salida-Excel%20con%20gr%C3%A1ficos-217346?logo=microsoftexcel&logoColor=white)

Convierte los archivos de ventas desordenados de un negocio (CSV y Excel, uno por sucursal, cada uno con su propio formato) en un **reporte de Excel listo para presentar**, en segundos y con un solo comando.

> 📥 **Mira el resultado:** descarga [`reporte_ventas.xlsx`](reporte_ventas.xlsx), un reporte real generado con los datos de ejemplo.

## El problema que resuelve

En muchos negocios alguien pasa horas cada semana o cada mes juntando a mano los archivos de cada sucursal o vendedor. Las columnas se llaman distinto, las fechas vienen en varios formatos, los precios están escritos como texto (`$18.500`, `18,500.00`), hay filas repetidas y errores de digitación. Este script hace todo ese trabajo solo.

## Antes y después

**Antes:** así llegan los datos de una sucursal (fragmento real de [`ventas_sucursal_norte.csv`](datos_ejemplo/ventas_sucursal_norte.csv)):

| Fecha | Producto | Cant. | Precio unitario | Vendedor |
|---|---|---|---|---|
| 15/05/2026 | Café en grano 1kg | 12 | `42,000.00` | Laura Gómez |
| 07/07/2026 | `␣␣Café en grano 1kg␣` | 4 | `42000` | Camila Torres |
| 07/07/2026 | TÉ VERDE X20 | 9 | `$12.900` | Camila Torres |
| 15/05/2026 | AROMÁTICA FRUTOS ROJOS | 1 | `8,400.00` | Julián Mejía |

Y la otra sucursal usa columnas distintas (`FECHA VENTA`, `Articulo`, `Unidades`, `Valor Unit`, `Asesor`) y fechas como `14-jul-2026`.

**Después:** todo queda unificado y limpio:

| Fecha | Sucursal | Producto | Cantidad | Precio | Total |
|---|---|---|---|---|---|
| 15/05/2026 | Norte | Café en grano 1kg | 12 | $42.000 | $504.000 |
| 07/07/2026 | Norte | Café en grano 1kg | 4 | $42.000 | $168.000 |
| 07/07/2026 | Norte | Té verde x20 | 9 | $12.900 | $116.100 |
| 15/05/2026 | Norte | Aromática frutos rojos | 1 | $8.400 | $8.400 |

Además se genera un resumen con indicadores, gráficos y una lista de las filas con errores para corregir.

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

Requisitos: [Python 3.10 o superior](https://www.python.org/downloads/) (en Windows, marca **"Add Python to PATH"** al instalar).

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

Para usarlo con tus propios archivos, ponlos en una carpeta y cambia `--entrada` por esa carpeta.

## Estructura del proyecto

```
reporte.py                 # script principal: lee, limpia y genera el Excel
generar_datos_ejemplo.py   # crea archivos de ventas "desordenados" para probar
datos_ejemplo/             # archivos de ejemplo (CSV de Norte, Excel de Centro)
reporte_ventas.xlsx        # reporte generado con los datos de ejemplo
requirements.txt           # dependencias (pandas, openpyxl)
```

## Adaptarlo a un cliente nuevo

Casi siempre basta con agregar los nombres de columna del cliente en el diccionario `ALIAS_COLUMNAS`, al inicio de `reporte.py`. Si falta una columna, el script dice exactamente cuál y qué columnas encontró en el archivo.

Extensiones típicas que se pueden ofrecer:
- Envío automático del reporte por correo cada lunes (Programador de tareas de Windows o cron)
- Lectura directa desde Google Sheets o desde la exportación del software de facturación
- Comparación contra el mes anterior y alertas cuando un producto baja de cierto nivel

## ¿Necesitas algo así para tu negocio?

Adapto este reporte a los archivos de tu empresa (ventas, inventario, cartera, nómina…) o automatizo otras tareas repetitivas en Excel. Escríbeme por mi perfil de GitHub [@Samu607](https://github.com/Samu607).

## Tecnologías

Python 3.10+, pandas y openpyxl (gráficos nativos de Excel, sin depender de imágenes).
