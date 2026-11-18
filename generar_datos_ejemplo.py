"""
Genera archivos de ventas de ejemplo "desordenados", como los que suele
tener un negocio real: cada sucursal exporta con columnas distintas,
fechas en varios formatos, precios escritos como texto, filas repetidas
y filas vacías.

Uso:
    python generar_datos_ejemplo.py
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)

CARPETA = Path(__file__).parent / "datos_ejemplo"

PRODUCTOS = {
    "Café molido 500g": 18_500,
    "Café en grano 1kg": 42_000,
    "Chocolate de mesa": 9_800,
    "Panela orgánica": 6_500,
    "Galletas de avena": 7_200,
    "Té verde x20": 12_900,
    "Aromática frutos rojos": 8_400,
}
VENDEDORES = ["Laura Gómez", "Andrés Ríos", "Camila Torres", "Julián Mejía", "Sara Henao"]
MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def fecha_aleatoria() -> date:
    inicio = date(2026, 5, 1)
    return inicio + timedelta(days=random.randint(0, 183))


def precio_como_texto(valor: int):
    """Escribe el precio de formas distintas, como pasa en la vida real."""
    estilo = random.choice(["pesos", "plano", "coma", "numero"])
    if estilo == "pesos":
        return "$" + f"{valor:,}".replace(",", ".")      # $18.500
    if estilo == "plano":
        return str(valor)                                 # 18500
    if estilo == "coma":
        return f"{valor:,}.00"                            # 18,500.00
    return valor                                          # número real


def ventas(n: int) -> pd.DataFrame:
    filas = []
    for _ in range(n):
        producto = random.choice(list(PRODUCTOS))
        filas.append({
            "fecha": fecha_aleatoria(),
            "producto": random.choice([producto, producto.upper(), producto.lower(), f"  {producto} "]),
            "cantidad": random.randint(1, 12),
            "precio": precio_como_texto(PRODUCTOS[producto]),
            "vendedor": random.choice(VENDEDORES),
        })
    return pd.DataFrame(filas).astype(object)


def ensuciar(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega duplicados, filas vacías y algunos errores de digitación."""
    duplicados = df.sample(6, random_state=1)
    vacias = pd.DataFrame([{c: None for c in df.columns}] * 3)
    errores = df.sample(3, random_state=2).copy()
    errores.iloc[0, errores.columns.get_loc("cantidad")] = "doce"      # cantidad en letras
    errores.iloc[1, errores.columns.get_loc("fecha")] = "31/02/2026"   # fecha que no existe
    errores.iloc[2, errores.columns.get_loc("precio")] = ""            # precio vacío
    sucio = pd.concat([df, duplicados, vacias, errores], ignore_index=True)
    return sucio.sample(frac=1, random_state=3).reset_index(drop=True)


def formatear_fechas(df: pd.DataFrame, formato) -> pd.DataFrame:
    df = df.copy()
    df["fecha"] = df["fecha"].map(lambda d: formato(d) if isinstance(d, date) else d)
    return df


def main() -> None:
    CARPETA.mkdir(exist_ok=True)

    # Sucursal Norte: CSV, fechas dd/mm/aaaa
    norte = formatear_fechas(ensuciar(ventas(140)), lambda d: d.strftime("%d/%m/%Y"))
    norte = norte.rename(columns={
        "fecha": "Fecha", "producto": "Producto", "cantidad": "Cant.",
        "precio": "Precio unitario", "vendedor": "Vendedor",
    })

    # Sucursal Centro: Excel, fechas mezcladas (2026-07-14 y 14-jul-2026)
    centro = formatear_fechas(
        ensuciar(ventas(120)),
        lambda d: random.choice([d.isoformat(), f"{d.day}-{MESES_ES[d.month - 1]}-{d.year}"]),
    )
    centro = centro.rename(columns={
        "fecha": "FECHA VENTA", "producto": "Articulo", "cantidad": "Unidades",
        "precio": "Valor Unit", "vendedor": "Asesor",
    })

    norte.to_csv(CARPETA / "ventas_sucursal_norte.csv", index=False, encoding="utf-8-sig")
    centro.to_excel(CARPETA / "ventas_sucursal_centro.xlsx", index=False)

    print(f"Listo: {len(norte)} filas (Norte) y {len(centro)} filas (Centro) en {CARPETA}")


if __name__ == "__main__":
    main()
