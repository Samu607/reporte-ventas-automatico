"""
Reporte de ventas automático
============================

Lee todos los archivos de ventas (.csv y .xlsx) de una carpeta, aunque cada
uno tenga columnas, fechas y precios en formatos distintos; los limpia y une,
y genera un reporte en Excel listo para presentar:

  * Resumen con indicadores clave y gráficos por mes y por sucursal
  * Ranking de productos y de vendedores
  * Datos limpios unificados (con filtros)
  * Hoja de "Problemas encontrados": qué filas se descartaron y por qué

Uso:
    python reporte.py --entrada datos_ejemplo --salida reporte_ventas.xlsx
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# 1. Configuración: nombres de columna que se aceptan para cada campo.
#    Para un cliente nuevo normalmente basta con agregar alias aquí.
# --------------------------------------------------------------------------
ALIAS_COLUMNAS = {
    "fecha": ["fecha", "fecha venta", "fecha_venta", "dia", "date"],
    "producto": ["producto", "articulo", "item", "referencia", "descripcion"],
    "cantidad": ["cantidad", "cant", "cant.", "unidades", "qty"],
    "precio": ["precio", "precio unitario", "valor unit", "valor unitario", "precio_unit", "price"],
    "vendedor": ["vendedor", "asesor", "cajero", "empleado", "seller"],
}

MESES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}
NOMBRE_MES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
              "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Estilos del reporte
COLOR_PRINCIPAL = "1F4E78"
COLOR_SUAVE = "DDEBF7"
TITULO = Font(name="Calibri", size=18, bold=True, color=COLOR_PRINCIPAL)
SUBTITULO = Font(name="Calibri", size=11, italic=True, color="595959")
ENCABEZADO = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
RELLENO_ENC = PatternFill("solid", fgColor=COLOR_PRINCIPAL)
RELLENO_KPI = PatternFill("solid", fgColor=COLOR_SUAVE)
BORDE = Border(bottom=Side(style="thin", color="BFBFBF"))
PESOS = '"$"#,##0'
PORCENTAJE = "0.0%"


# --------------------------------------------------------------------------
# 2. Limpieza
# --------------------------------------------------------------------------
def normalizar_texto(texto: str) -> str:
    """minúsculas, sin tildes ni espacios sobrantes: 'Fecha  Venta ' -> 'fecha venta'."""
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto).strip().lower()


def mapear_columnas(df: pd.DataFrame, archivo: str) -> pd.DataFrame:
    mapa = {}
    for columna in df.columns:
        clave = normalizar_texto(columna)
        for campo, alias in ALIAS_COLUMNAS.items():
            if clave in alias:
                mapa[columna] = campo
    df = df.rename(columns=mapa)
    faltan = [c for c in ALIAS_COLUMNAS if c not in df.columns]
    if faltan:
        raise ValueError(
            f"En '{archivo}' no encontré las columnas: {', '.join(faltan)}. "
            f"Columnas del archivo: {', '.join(map(str, df.columns))}. "
            "Agrega el nombre a ALIAS_COLUMNAS en reporte.py."
        )
    return df[list(ALIAS_COLUMNAS)]


def convertir_precio(valor) -> float | None:
    """Entiende $18.500 · 18500 · 18,500.00 · 18.500,50 · 18500.0"""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = re.sub(r"[^\d.,-]", "", str(valor))
    if not texto:
        return None
    if "," in texto and "." in texto:
        # El último separador que aparece es el decimal
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        partes = texto.split(",")
        texto = texto.replace(",", "") if len(partes[-1]) == 3 else texto.replace(",", ".")
    elif "." in texto:
        partes = texto.split(".")
        if len(partes) > 2 or len(partes[-1]) == 3:   # 18.500 o 1.200.000 -> miles
            texto = texto.replace(".", "")
    try:
        return float(texto)
    except ValueError:
        return None


def convertir_fecha(valor) -> pd.Timestamp | None:
    """Entiende 14/07/2026 · 2026-07-14 · 14-jul-2026 · fechas reales de Excel."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (datetime, pd.Timestamp)):
        return pd.Timestamp(valor)
    texto = normalizar_texto(valor)
    if not texto:
        return None
    m = re.fullmatch(r"(\d{1,2})[-/ ]([a-z]{3})[a-z]*[-/ ](\d{4})", texto)
    if m and m.group(2) in MESES:
        texto = f"{m.group(3)}-{MESES[m.group(2)]:02d}-{int(m.group(1)):02d}"
    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"]
    for formato in formatos:
        try:
            return pd.Timestamp(datetime.strptime(texto.split(" ")[0], formato))
        except ValueError:
            continue
    return None


def convertir_cantidad(valor) -> float | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        numero = float(str(valor).replace(",", "."))
    except ValueError:
        return None
    return numero if numero > 0 else None


def nombre_producto(valor) -> str | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    if not texto:
        return None
    # "CAFÉ MOLIDO 500G" -> "Café molido 500g"
    return texto[0].upper() + texto[1:].lower()


def mostrar(valor) -> str:
    """Cómo se ve un valor original en la hoja de Problemas."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)) or str(valor).strip() == "":
        return "vacío"
    return f"'{valor}'"


def sucursal_desde_archivo(ruta: Path) -> str:
    nombre = re.sub(r"(?i)^(ventas?_)?(sucursal_)?", "", ruta.stem)
    return nombre.replace("_", " ").strip().title() or ruta.stem


def leer_archivo(ruta: Path) -> pd.DataFrame:
    if ruta.suffix.lower() == ".csv":
        # Detecta automáticamente si el separador es coma o punto y coma
        return pd.read_csv(ruta, sep=None, engine="python", dtype=object, encoding="utf-8-sig")
    return pd.read_excel(ruta, dtype=object)


def cargar_y_limpiar(carpeta: Path) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    archivos = sorted(p for p in carpeta.iterdir()
                      if p.suffix.lower() in {".csv", ".xlsx", ".xls"} and not p.name.startswith("~$"))
    if not archivos:
        raise FileNotFoundError(f"No hay archivos .csv o .xlsx en {carpeta}")

    limpias, problemas = [], []
    estadisticas = {"archivos": len(archivos), "filas_leidas": 0, "vacias": 0, "duplicadas": 0}

    for ruta in archivos:
        crudo = leer_archivo(ruta)
        estadisticas["filas_leidas"] += len(crudo)
        df = mapear_columnas(crudo, ruta.name)
        df["fila_original"] = df.index + 2          # +2: encabezado y base 1, como en Excel
        df["archivo"] = ruta.name
        df["sucursal"] = sucursal_desde_archivo(ruta)

        vacias = df[list(ALIAS_COLUMNAS)].isna().all(axis=1)
        estadisticas["vacias"] += int(vacias.sum())
        df = df[~vacias]

        for _, fila in df.iterrows():
            registro = {
                "fecha": convertir_fecha(fila["fecha"]),
                "producto": nombre_producto(fila["producto"]),
                "cantidad": convertir_cantidad(fila["cantidad"]),
                "precio": convertir_precio(fila["precio"]),
                "vendedor": nombre_producto(fila["vendedor"]) and str(fila["vendedor"]).strip(),
                "sucursal": fila["sucursal"],
                "archivo": fila["archivo"],
                "fila_original": fila["fila_original"],
            }
            motivos = []
            if registro["fecha"] is None:
                motivos.append(f"fecha no válida ({mostrar(fila['fecha'])})")
            if registro["producto"] is None:
                motivos.append("sin producto")
            if registro["cantidad"] is None:
                motivos.append(f"cantidad no válida ({mostrar(fila['cantidad'])})")
            if registro["precio"] is None:
                motivos.append(f"precio no válido ({mostrar(fila['precio'])})")
            if motivos:
                problemas.append({"archivo": ruta.name, "fila": fila["fila_original"],
                                  "motivo": "; ".join(motivos)})
            else:
                limpias.append(registro)

    datos = pd.DataFrame(limpias)
    clave = ["fecha", "producto", "cantidad", "precio", "vendedor", "sucursal"]
    repetidas = datos.duplicated(subset=clave, keep="first")
    for _, fila in datos[repetidas].iterrows():
        problemas.append({"archivo": fila["archivo"], "fila": fila["fila_original"],
                          "motivo": "fila duplicada (se dejó solo la primera)"})
    estadisticas["duplicadas"] = int(repetidas.sum())
    datos = datos[~repetidas].copy()

    datos["total"] = datos["cantidad"] * datos["precio"]
    datos["mes"] = datos["fecha"].dt.to_period("M")
    datos = datos.sort_values(["fecha", "sucursal"]).reset_index(drop=True)
    return datos, pd.DataFrame(problemas, columns=["archivo", "fila", "motivo"]), estadisticas


# --------------------------------------------------------------------------
# 3. Escritura del Excel
# --------------------------------------------------------------------------
def escribir_tabla(hoja, fila: int, col: int, encabezados: list[str], filas: list[list],
                   formatos: dict[int, str] | None = None) -> tuple[int, int]:
    """Escribe una tabla con encabezado azul. Devuelve (fila_inicio_datos, fila_fin)."""
    formatos = formatos or {}
    for j, titulo in enumerate(encabezados):
        celda = hoja.cell(row=fila, column=col + j, value=titulo)
        celda.font, celda.fill = ENCABEZADO, RELLENO_ENC
        celda.alignment = Alignment(horizontal="center", vertical="center")
    for i, valores in enumerate(filas, start=1):
        for j, valor in enumerate(valores):
            celda = hoja.cell(row=fila + i, column=col + j, value=valor)
            celda.border = BORDE
            if j in formatos:
                celda.number_format = formatos[j]
    return fila + 1, fila + len(filas)


def ajustar_anchos(hoja, minimo: int = 10, maximo: int = 45) -> None:
    for columna in hoja.columns:
        largo = max((len(str(c.value)) for c in columna if c.value is not None), default=0)
        letra = get_column_letter(columna[0].column)
        hoja.column_dimensions[letra].width = max(minimo, min(maximo, largo + 3))


def grafico_barras(titulo: str, hoja, col_cat: int, col_val: int, fila_ini: int, fila_fin: int,
                   eje_y: str = "Ventas ($)") -> BarChart:
    grafico = BarChart()
    grafico.title, grafico.y_axis.title = titulo, eje_y
    grafico.style, grafico.height, grafico.width = 10, 7.5, 16
    grafico.legend = None
    grafico.add_data(Reference(hoja, min_col=col_val, min_row=fila_ini - 1, max_row=fila_fin),
                     titles_from_data=True)
    grafico.set_categories(Reference(hoja, min_col=col_cat, min_row=fila_ini, max_row=fila_fin))
    grafico.y_axis.numFmt = PESOS
    grafico.y_axis.delete = False
    grafico.x_axis.delete = False
    return grafico


def hoja_resumen(libro: Workbook, datos: pd.DataFrame, estadisticas: dict) -> None:
    hoja = libro.active
    hoja.title = "Resumen"
    hoja.sheet_view.showGridLines = False

    inicio, fin = datos["fecha"].min(), datos["fecha"].max()
    hoja["B2"] = "Reporte de ventas"
    hoja["B2"].font = TITULO
    hoja["B3"] = (f"Periodo: {inicio:%d/%m/%Y} al {fin:%d/%m/%Y}  ·  "
                  f"Generado automáticamente el {datetime.now():%d/%m/%Y %H:%M}")
    hoja["B3"].font = SUBTITULO

    total = datos["total"].sum()
    kpis = [
        ("Ventas totales", total, PESOS),
        ("Transacciones", len(datos), "#,##0"),
        ("Unidades vendidas", datos["cantidad"].sum(), "#,##0"),
        ("Ticket promedio", total / len(datos), PESOS),
    ]
    for j, (nombre, valor, formato) in enumerate(kpis):
        col = 2 + j * 2
        etiqueta = hoja.cell(row=5, column=col, value=nombre)
        etiqueta.font = Font(size=10, color="595959")
        numero = hoja.cell(row=6, column=col, value=float(valor))
        numero.font = Font(size=16, bold=True, color=COLOR_PRINCIPAL)
        numero.number_format = formato
        for fila in (5, 6):
            for c in (col, col + 1):
                hoja.cell(row=fila, column=c).fill = RELLENO_KPI
        hoja.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
        hoja.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col + 1)

    # Ventas por mes
    por_mes = datos.groupby("mes")["total"].sum()
    filas_mes = [[f"{NOMBRE_MES[p.month]} {p.year}", float(v)] for p, v in por_mes.items()]
    hoja["B9"] = "Ventas por mes"
    hoja["B9"].font = Font(size=12, bold=True, color=COLOR_PRINCIPAL)
    ini, fin_mes = escribir_tabla(hoja, 10, 2, ["Mes", "Ventas"], filas_mes, {1: PESOS})
    hoja.add_chart(grafico_barras("Ventas por mes", hoja, 2, 3, ini, fin_mes), "F9")

    # Ventas por sucursal
    fila_suc = max(fin_mes + 3, 26)
    por_suc = datos.groupby("sucursal")["total"].sum().sort_values(ascending=False)
    filas_suc = [[s, float(v), float(v / total)] for s, v in por_suc.items()]
    hoja.cell(row=fila_suc, column=2, value="Ventas por sucursal").font = Font(
        size=12, bold=True, color=COLOR_PRINCIPAL)
    ini_s, fin_s = escribir_tabla(hoja, fila_suc + 1, 2, ["Sucursal", "Ventas", "% del total"],
                                  filas_suc, {1: PESOS, 2: PORCENTAJE})
    torta = PieChart()
    torta.title, torta.height, torta.width = "Participación por sucursal", 7.5, 10
    torta.add_data(Reference(hoja, min_col=3, min_row=ini_s - 1, max_row=fin_s), titles_from_data=True)
    torta.set_categories(Reference(hoja, min_col=2, min_row=ini_s, max_row=fin_s))
    hoja.add_chart(torta, f"F{fila_suc}")

    # Calidad de datos
    fila_cal = fila_suc + 17
    hoja.cell(row=fila_cal, column=2, value="Calidad de los datos").font = Font(
        size=12, bold=True, color=COLOR_PRINCIPAL)
    descartadas = (estadisticas["filas_leidas"] - estadisticas["vacias"]
                   - estadisticas["duplicadas"] - len(datos))
    filas_cal = [
        ["Archivos procesados", estadisticas["archivos"]],
        ["Filas leídas", estadisticas["filas_leidas"]],
        ["Filas vacías ignoradas", estadisticas["vacias"]],
        ["Duplicados eliminados", estadisticas["duplicadas"]],
        ["Filas con errores (ver hoja Problemas)", descartadas],
        ["Filas válidas en el reporte", len(datos)],
    ]
    escribir_tabla(hoja, fila_cal + 1, 2, ["Concepto", "Cantidad"], filas_cal, {1: "#,##0"})

    hoja.column_dimensions["A"].width = 3
    hoja.column_dimensions["B"].width = 38
    for letra in "CDEFGHI":
        hoja.column_dimensions[letra].width = 15


def hoja_ranking(libro: Workbook, datos: pd.DataFrame, campo: str, titulo: str) -> None:
    hoja = libro.create_sheet(titulo)
    hoja.sheet_view.showGridLines = False
    total = datos["total"].sum()
    grupo = (datos.groupby(campo)
             .agg(unidades=("cantidad", "sum"), ventas=("total", "sum"), transacciones=("total", "size"))
             .sort_values("ventas", ascending=False))
    filas = [[i, nombre, float(f.unidades), int(f.transacciones), float(f.ventas), float(f.ventas / total)]
             for i, (nombre, f) in enumerate(grupo.iterrows(), start=1)]
    hoja["B2"] = f"Ranking de {titulo.lower()}"
    hoja["B2"].font = TITULO
    ini, fin = escribir_tabla(
        hoja, 4, 2, ["#", campo.capitalize(), "Unidades", "Transacciones", "Ventas", "% del total"],
        filas, {2: "#,##0", 3: "#,##0", 4: PESOS, 5: PORCENTAJE})
    hoja.add_chart(grafico_barras(f"Ventas por {campo}", hoja, 3, 6, ini, fin), "I4")
    ajustar_anchos(hoja)
    hoja.column_dimensions["A"].width = 3
    hoja.column_dimensions["B"].width = 5


def hoja_datos(libro: Workbook, datos: pd.DataFrame) -> None:
    hoja = libro.create_sheet("Datos limpios")
    columnas = ["fecha", "sucursal", "producto", "vendedor", "cantidad", "precio", "total", "archivo"]
    filas = [[f.fecha.to_pydatetime(), f.sucursal, f.producto, f.vendedor, float(f.cantidad),
              float(f.precio), float(f.total), f.archivo] for f in datos.itertuples()]
    _, fin = escribir_tabla(hoja, 1, 1, [c.capitalize() for c in columnas], filas,
                            {0: "dd/mm/yyyy", 4: "#,##0", 5: PESOS, 6: PESOS})
    hoja.auto_filter.ref = f"A1:{get_column_letter(len(columnas))}{fin}"
    hoja.freeze_panes = "A2"
    ajustar_anchos(hoja)


def hoja_problemas(libro: Workbook, problemas: pd.DataFrame) -> None:
    hoja = libro.create_sheet("Problemas")
    hoja["A1"] = "Filas que no se incluyeron en el reporte"
    hoja["A1"].font = Font(size=14, bold=True, color="C00000")
    hoja["A2"] = "Corrige estas filas en el archivo original y vuelve a ejecutar el reporte."
    hoja["A2"].font = SUBTITULO
    filas = problemas.sort_values(["archivo", "fila"]).values.tolist() if len(problemas) else [
        ["—", "—", "No se encontraron problemas"]]
    escribir_tabla(hoja, 4, 1, ["Archivo", "Fila en el archivo", "Motivo"], filas)
    ajustar_anchos(hoja, maximo=70)


def generar_reporte(entrada: Path, salida: Path) -> dict:
    datos, problemas, estadisticas = cargar_y_limpiar(entrada)
    libro = Workbook()
    hoja_resumen(libro, datos, estadisticas)
    hoja_ranking(libro, datos, "producto", "Productos")
    hoja_ranking(libro, datos, "vendedor", "Vendedores")
    hoja_datos(libro, datos)
    hoja_problemas(libro, problemas)
    libro.save(salida)
    estadisticas.update(validas=len(datos), problemas=len(problemas), total=float(datos["total"].sum()))
    return estadisticas


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera un reporte de ventas en Excel a partir de "
                                                 "archivos .csv/.xlsx desordenados.")
    parser.add_argument("--entrada", default="datos_ejemplo", help="Carpeta con los archivos de ventas")
    parser.add_argument("--salida", default="reporte_ventas.xlsx", help="Archivo Excel a generar")
    args = parser.parse_args()

    r = generar_reporte(Path(args.entrada), Path(args.salida))
    print(f"✔ Reporte generado: {args.salida}")
    print(f"  {r['archivos']} archivos · {r['filas_leidas']} filas leídas · {r['validas']} válidas")
    print(f"  {r['vacias']} vacías · {r['duplicadas']} duplicadas · "
          f"{r['problemas'] - r['duplicadas']} con errores")
    print(f"  Ventas totales: ${r['total']:,.0f}".replace(",", "."))


if __name__ == "__main__":
    main()
