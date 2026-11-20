"""
Subir a marketplaces — Kbas Office
Lógica sin Streamlit (probada aparte). De momento: Faire.
Basado en faire_generador.py (versión con interruptor KBAS / VOLUM).
"""
from __future__ import annotations

import io
import re

import pandas as pd
from openpyxl import load_workbook

# ═══════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN POR DEFECTO (cámbiala aquí para que quede guardada)
# ═══════════════════════════════════════════════════════════════════
MARKUP = 1.20        # +20 % al precio venta → mayorista Faire
RETAIL_MULT = 2.5    # PVP = mayorista × 2,5
MAX_NAME_LEN = 60    # límite de Faire
IMG_COLS = ["Web imagen 1", "Web imagen 2", "Web imagen 3", "Web imagen 4"]
WEIGHT_UNIT = "kg"

PAIS_VOLUM = "CN"    # Volum: todo China

# KBAS: país (código ISO) según los 3 dígitos del proveedor (posiciones 3-4-5 de la Ref.)
# Ej: KB3882601 → "388" → CN.  ⚠️ En verano cambian proveedores.
PAIS_POR_PROVEEDOR = {
    "388": "CN",
    "366": "CN",
    "340": "CN",
    "351": "CN",
    "387": "CN",
    "386": "CN",
    "001": "ES",
    "348": "CN",   # ⚠️ en el script de Ankorstore ponía "India" en el comentario: revisar
    "159": "ES",
}

# Faire quiere el nombre del país; Ankorstore, el código ISO
ISO_A_NOMBRE = {
    "CN": "China", "ES": "España", "IN": "India", "PT": "Portugal", "IT": "Italia",
    "FR": "Francia", "TR": "Turquía", "MA": "Marruecos", "VN": "Vietnam", "BD": "Bangladés",
    "PK": "Pakistán", "ID": "Indonesia", "TH": "Tailandia", "MX": "México",
}

# Peso (kg) por las 2 primeras letras de la Ref.
PESOS_POR_PREFIJO = {
    "VB": 0.5, "VK": 0.3, "VF": 0.4,                       # Volum
    "KB": 0.5, "KC": 0.5, "KS": 0.2, "KT": 0.4, "KD": 0.5,  # Kbas
}

FAIRE_SHEET = "Productos|Español"
FAIRE_START_ROW = 4

ANKOR_SHEET = "Tus productos"
ANKOR_START_ROW = 2
IVA = 21             # Ankorstore columna O
PACK_UNITS = 1       # Ankorstore columna Q
STOCK = 100          # Ankorstore columna R (los agotados se quitan a mano)
DESC_MIN = 30        # Ankorstore exige descripción ≥ 30 caracteres


# ───────────────────────────── utilidades ─────────────────────────────

def _vacio(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip())


def parse_medidas(s):
    if _vacio(s):
        return (None, None, None)
    nums = re.findall(r"\d+(?:[.,]\d+)?", str(s).lower())
    if len(nums) < 3:
        return (None, None, None)
    vals = [float(n.replace(",", ".")) for n in nums[:3]]
    return vals[0], vals[1], vals[2]


def limitar(s, n=MAX_NAME_LEN):
    return None if _vacio(s) else str(s).strip()[:n]


def clean_url(u):
    if _vacio(u):
        return None
    u = str(u).strip()
    return u if u.startswith("http") else None


def imagenes_fila(row, img_cols=IMG_COLS, solo_principal=False):
    urls = []
    for col in img_cols:
        u = clean_url(row.get(col))
        if u and u not in urls:
            urls.append(u)
        if solo_principal and urls:
            break
    return " ".join(urls) if urls else None   # Faire: separadas por ESPACIO


def clean_gtin(v):
    if _vacio(v):
        return None
    try:
        return str(int(float(v)))
    except (ValueError, TypeError):
        return str(v).strip()


def peso_por_ref(sku, pesos):
    if _vacio(sku):
        return None
    return pesos.get(str(sku).strip().upper()[:2])


def pais_por_ref(sku, marca, pais_volum, paises):
    if marca == "VOLUM":
        return pais_volum, None
    if _vacio(sku):
        return None, None
    codigo = str(sku).strip().upper()[2:5]
    return paises.get(codigo), codigo


def leer_erp(fichero) -> pd.DataFrame:
    df = pd.read_excel(fichero)
    # Algunas exportaciones del ERP traen la "ó" rota
    df.columns = [str(c).replace("Descripci�n", "Descripción").strip() for c in df.columns]
    return df.reset_index(drop=True)


def quitar_duplicados(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Ankorstore va sin variantes: una fila por Ref. (se queda la primera)."""
    refs = df["Ref."].astype(str).str.strip()
    dup = refs[refs.duplicated() & refs.ne("") & refs.ne("nan")].unique().tolist()
    return df[~(refs.duplicated() & refs.ne("") & refs.ne("nan"))].reset_index(drop=True), dup


def comprobar_columnas(df: pd.DataFrame) -> list[str]:
    necesarias = ["Ref.", "Descripción", "Precio venta"]
    return [c for c in necesarias if c not in df.columns]


# ───────────────────────────── Faire ─────────────────────────────

def generar_faire(
    erp: pd.DataFrame,
    plantilla,                       # ruta o bytes de faire-products-template.xlsx
    marca: str = "KBAS",
    markup: float = MARKUP,
    retail_mult: float = RETAIL_MULT,
    paises: dict | None = None,
    pesos: dict | None = None,
    pais_volum: str = PAIS_VOLUM,
    solo_principal: bool = False,
) -> tuple[bytes, dict]:
    marca = marca.upper()
    paises = PAIS_POR_PROVEEDOR if paises is None else paises
    pesos = PESOS_POR_PREFIJO if pesos is None else pesos

    if isinstance(plantilla, (bytes, bytearray)):
        plantilla = io.BytesIO(plantilla)
    wb = load_workbook(plantilla)
    if FAIRE_SHEET not in wb.sheetnames:
        raise ValueError(f"La plantilla no tiene la hoja '{FAIRE_SHEET}'. ¿Es la plantilla de Faire en español?")
    ws = wb[FAIRE_SHEET]

    avisos = {"sin_imagen": [], "sin_peso": [], "sin_gtin": [], "sin_pais": [], "sin_precio": [], "sin_medidas": []}
    filas = []

    fila = FAIRE_START_ROW - 1
    for _, r in erp.iterrows():
        sku = None if _vacio(r.get("Ref.")) else str(r.get("Ref.")).strip()
        nombre = None if _vacio(r.get("Descripción")) else str(r.get("Descripción")).strip()
        if not sku and not nombre:
            continue
        fila += 1   # sin huecos aunque el ERP traiga filas vacías

        largo, ancho, alto = parse_medidas(r.get("Web medidas"))
        if largo is None:
            avisos["sin_medidas"].append(sku)

        precio = r.get("Precio venta")
        wh = rt = None
        if not _vacio(precio):
            try:
                wh = round(float(str(precio).replace(",", ".")) * markup, 2)
                rt = round(wh * retail_mult, 2)
            except ValueError:
                pass
        if wh is None:
            avisos["sin_precio"].append(sku)

        imgs = imagenes_fila(r, solo_principal=solo_principal)
        if imgs is None:
            avisos["sin_imagen"].append(sku)

        peso = peso_por_ref(sku, pesos)
        if peso is None:
            avisos["sin_peso"].append(sku)

        gtin = clean_gtin(r.get("Barcode"))
        if gtin is None:
            avisos["sin_gtin"].append(sku)

        iso, cod = pais_por_ref(sku, marca, pais_volum, paises)
        pais = ISO_A_NOMBRE.get(iso, iso) if iso else None
        if pais is None:
            avisos["sin_pais"].append(f"{sku} (prov. {cod})")

        nombre_unico = limitar(f"{sku} - {nombre}" if (nombre and sku) else (nombre or sku))

        ws[f"A{fila}"] = nombre_unico
        ws[f"B{fila}"] = "Publicado"
        ws[f"E{fila}"] = nombre
        ws[f"F{fila}"] = "Por artículo"
        ws[f"H{fila}"] = 1
        if peso is not None:
            ws[f"I{fila}"] = peso
            ws[f"J{fila}"] = WEIGHT_UNIT
        ws[f"K{fila}"] = largo
        ws[f"L{fila}"] = ancho
        ws[f"M{fila}"] = alto
        ws[f"N{fila}"] = "cm"
        ws[f"V{fila}"] = sku
        if gtin is not None:
            ws[f"W{fila}"] = gtin
            ws[f"W{fila}"].number_format = "@"
        ws[f"AE{fila}"] = wh
        ws[f"AF{fila}"] = rt
        ws[f"AU{fila}"] = imgs
        if pais is not None:
            ws[f"AV{fila}"] = pais

        filas.append({
            "Ref.": sku, "Nombre Faire": nombre_unico, "Mayorista €": wh, "PVP €": rt,
            "Peso kg": peso, "Medidas": f"{largo:g}×{ancho:g}×{alto:g}" if largo is not None else None,
            "EAN": gtin, "País": pais, "Nº fotos": len(imgs.split(" ")) if imgs else 0,
        })

    out = io.BytesIO()
    wb.save(out)
    avisos["vista"] = pd.DataFrame(filas)
    avisos["total"] = len(filas)
    return out.getvalue(), avisos


# ───────────────────────────── Ankorstore ─────────────────────────────

def descripcion_min(nombre, sku, minimo=DESC_MIN):
    base = f"{nombre or ''} - B2B - SKU {sku or ''}".strip()
    if len(base) < minimo:
        base += " | Venta mayorista B2B"
    while len(base) < minimo:
        base += "."
    return base


def generar_ankorstore(
    erp: pd.DataFrame,
    plantilla,
    marca: str = "KBAS",
    markup: float = MARKUP,
    retail_mult: float = RETAIL_MULT,
    paises: dict | None = None,
    pesos: dict | None = None,
    pais_volum: str = PAIS_VOLUM,
    iva: float = IVA,
    stock: int = STOCK,
) -> tuple[bytes, dict]:
    marca = marca.upper()
    paises = PAIS_POR_PROVEEDOR if paises is None else paises
    pesos = PESOS_POR_PREFIJO if pesos is None else pesos

    if isinstance(plantilla, (bytes, bytearray)):
        plantilla = io.BytesIO(plantilla)
    wb = load_workbook(plantilla)
    ws = wb[ANKOR_SHEET] if ANKOR_SHEET in wb.sheetnames else wb.active

    # limpiar datos de ejemplo de la plantilla
    cols = ["A", "B", "C", "D", "E", "F", "H", "I", "J", "K", "L", "M", "N", "O", "Q", "R", "S",
            "U", "V", "W", "X", "Y", "Z", "AA"]
    for f in range(ANKOR_START_ROW, ws.max_row + 1):
        for c in cols:
            ws[f"{c}{f}"].value = None

    erp, duplicadas = quitar_duplicados(erp)
    avisos = {"sin_imagen": [], "sin_peso": [], "sin_gtin": [], "sin_pais": [], "sin_precio": [],
              "sin_medidas": [], "duplicadas": duplicadas}
    filas = []
    factor_iva = 1 + iva / 100
    fila = ANKOR_START_ROW - 1

    for _, r in erp.iterrows():
        sku = None if _vacio(r.get("Ref.")) else str(r.get("Ref.")).strip()
        nombre = None if _vacio(r.get("Descripción")) else str(r.get("Descripción")).strip()
        if not sku and not nombre:
            continue
        fila += 1

        medidas_txt = None if _vacio(r.get("Web medidas")) else str(r.get("Web medidas")).strip()
        largo, ancho, alto = parse_medidas(medidas_txt)
        if largo is None:
            avisos["sin_medidas"].append(sku)

        wh = pvp = None
        precio = r.get("Precio venta")
        if not _vacio(precio):
            try:
                wh = round(float(str(precio).replace(",", ".")) * markup, 2)
                pvp = round(round(wh * retail_mult, 2) * factor_iva, 2)   # retail CON IVA
            except ValueError:
                pass
        if wh is None:
            avisos["sin_precio"].append(sku)

        imgs = []
        for col in IMG_COLS:
            u = clean_url(r.get(col))
            if u and u not in imgs:
                imgs.append(u)
        if not imgs:
            avisos["sin_imagen"].append(sku)

        peso = peso_por_ref(sku, pesos)
        if peso is None:
            avisos["sin_peso"].append(sku)

        iso, cod = pais_por_ref(sku, marca, pais_volum, paises)
        if iso is None:
            avisos["sin_pais"].append(f"{sku} (prov. {cod})")

        gtin = clean_gtin(r.get("Barcode"))
        if gtin is None:
            avisos["sin_gtin"].append(sku)

        color = None if _vacio(r.get("Colores")) else str(r.get("Colores")).strip()

        ws[f"A{fila}"] = sku
        ws[f"B{fila}"] = nombre
        ws[f"C{fila}"] = descripcion_min(nombre, sku)
        ws[f"D{fila}"] = medidas_txt
        ws[f"E{fila}"] = color
        for idx, c in enumerate(["H", "I", "J", "K", "L"]):
            ws[f"{c}{fila}"] = imgs[idx] if idx < len(imgs) else None
        ws[f"M{fila}"] = wh
        ws[f"N{fila}"] = pvp
        ws[f"O{fila}"] = iva
        ws[f"Q{fila}"] = PACK_UNITS
        ws[f"R{fila}"] = stock
        if iso:
            ws[f"S{fila}"] = iso
        if gtin:
            ws[f"U{fila}"] = gtin
            ws[f"U{fila}"].number_format = "@"
        if largo is not None:
            ws[f"V{fila}"] = "cm"
            ws[f"W{fila}"] = largo
            ws[f"X{fila}"] = ancho
            ws[f"Y{fila}"] = alto
        if peso is not None:
            ws[f"Z{fila}"] = WEIGHT_UNIT
            ws[f"AA{fila}"] = peso

        filas.append({
            "Ref.": sku, "Nombre": nombre, "Mayorista €": wh, "PVP € (IVA incl.)": pvp,
            "Peso kg": peso, "Medidas": f"{largo:g}×{ancho:g}×{alto:g}" if largo is not None else None,
            "EAN": gtin, "País": iso, "Nº fotos": len(imgs),
        })

    out = io.BytesIO()
    wb.save(out)
    avisos["vista"] = pd.DataFrame(filas)
    avisos["total"] = len(filas)
    return out.getvalue(), avisos
