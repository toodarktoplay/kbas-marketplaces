"""
Subir a marketplaces — Kbas Office
Excel de la temporada (ERP) → archivo listo para subir a Faire o Ankorstore, para Kbas o Volum.
"""
from pathlib import Path

import pandas as pd
import streamlit as st

import marketplaces_core as core

st.set_page_config(page_title="Subir a marketplaces · Kbas Office", page_icon="🛒", layout="wide")

# Paleta oficial Kbas (la misma del portal)
MAUVE, SAND, CHOCOLATE, INK, BLUE = "#A88EA1", "#E1C9AF", "#522003", "#1A1A1A", "#7E9BC9"
CREMA, SUAVE = "#FDF9F3", "#7A6A5A"

st.markdown(
    f"""
    <style>
      .stApp {{ background: {CREMA}; }}
      h1, h2, h3 {{ color: {CHOCOLATE}; }}
      div[data-testid="stMetric"] {{
          background: white; border-radius: 12px; padding: 12px 16px;
          border-top: 4px solid {MAUVE}; box-shadow: 0 1px 4px rgba(0,0,0,.06);
      }}
      .nota {{ background: {SAND}; border-radius: 12px; padding: 12px 16px; color: {INK}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🛒 Subir a marketplaces")
st.caption("Sube el Excel de la temporada del ERP y descarga el archivo listo para Faire o Ankorstore.")

# ─────────────────────────── 1. Marca y marketplace ───────────────────────────
c1, c2 = st.columns(2)
with c1:
    marca = st.radio("Marca", ["KBAS", "VOLUM"], horizontal=True,
                     format_func=lambda m: "Kbas" if m == "KBAS" else "Volum")
with c2:
    marketplace = st.radio("Marketplace", ["Faire", "Ankorstore"], horizontal=True)

es_faire = marketplace == "Faire"
nombre_plantilla = "faire-products-template.xlsx" if es_faire else "ankorstore_plantilla.xlsx"

# ─────────────────────────── 2. Archivos ───────────────────────────
plantilla_repo = Path(__file__).parent / nombre_plantilla
c1, c2 = st.columns(2)
with c1:
    erp_file = st.file_uploader(f"Excel de la temporada {marca.title()} (exportación del ERP)", type=["xlsx", "xls"])
with c2:
    if plantilla_repo.exists():
        st.success(f"Plantilla de {marketplace} cargada ✓")
        plantilla_subida = st.file_uploader(f"¿Plantilla nueva de {marketplace}? (opcional)", type=["xlsx"],
                                            key=f"tpl_{marketplace}")
    else:
        plantilla_subida = st.file_uploader(f"Plantilla de {marketplace} ({nombre_plantilla})", type=["xlsx"],
                                            key=f"tpl_{marketplace}")

# ─────────────────────────── 3. Ajustes ───────────────────────────
with st.expander("⚙️ Precios, países y pesos"):
    p1, p2, p3 = st.columns(3)
    markup = p1.number_input("Mayorista = precio venta ×", value=core.MARKUP, step=0.05, format="%.2f")
    retail = p2.number_input("PVP = mayorista ×" + ("" if es_faire else " (luego + IVA)"),
                             value=core.RETAIL_MULT, step=0.1, format="%.2f")
    if es_faire:
        solo_principal = p3.checkbox("Solo la foto principal", value=False,
                                     help="Por defecto se envían hasta 4 fotos, separadas por espacio.")
    else:
        iva = p3.number_input("IVA %", value=core.IVA, step=1)
        stock = p3.number_input("Stock por referencia", value=core.STOCK, step=10,
                                help="Ankorstore lo pide obligatorio. Los agotados se quitan a mano.")

    t1, t2 = st.columns(2)
    with t1:
        if marca == "VOLUM":
            pais_volum = st.text_input("País de procedencia de toda Volum (código ISO: CN, IN, ES…)",
                                       value=core.PAIS_VOLUM).strip().upper()
            paises = core.PAIS_POR_PROVEEDOR
        else:
            pais_volum = core.PAIS_VOLUM
            st.markdown("**País por proveedor** (dígitos 3-4-5 de la Ref.) · código ISO: CN, IN, ES…")
            df_p = st.data_editor(
                pd.DataFrame(list(core.PAIS_POR_PROVEEDOR.items()), columns=["Proveedor", "País"]),
                num_rows="dynamic", hide_index=True, key="paises",
            )
            paises = {str(p).strip().zfill(3): str(c).strip().upper()
                      for p, c in df_p.values if pd.notna(p) and pd.notna(c) and str(p).strip() and str(c).strip()}
    with t2:
        st.markdown("**Peso por prefijo** (2 primeras letras, en kg)")
        prefijos = [p for p in core.PESOS_POR_PREFIJO if p.startswith("V" if marca == "VOLUM" else "K")]
        df_w = st.data_editor(
            pd.DataFrame([(p, core.PESOS_POR_PREFIJO[p]) for p in prefijos], columns=["Prefijo", "Peso kg"]),
            num_rows="dynamic", hide_index=True, key=f"pesos_{marca}",
        )
        pesos = {str(p).strip().upper(): float(w)
                 for p, w in df_w.values if pd.notna(p) and str(p).strip() and pd.notna(w)}
    st.caption("Los cambios aquí valen para esta sesión. Para dejarlos fijos, edítalos en marketplaces_core.py.")

# ─────────────────────────── 4. Generar ───────────────────────────
if not erp_file:
    st.markdown('<div class="nota">Sube el Excel del ERP con las columnas <b>Ref.</b>, <b>Descripción</b>, '
                "<b>Precio venta</b>, <b>Web medidas</b>, <b>Barcode</b> y <b>Web imagen 1–4</b>.</div>",
                unsafe_allow_html=True)
    st.stop()

plantilla = plantilla_subida.getvalue() if plantilla_subida else (plantilla_repo if plantilla_repo.exists() else None)
if plantilla is None:
    st.warning(f"Falta la plantilla de {marketplace}.")
    st.stop()

erp = core.leer_erp(erp_file)
faltan = core.comprobar_columnas(erp)
if faltan:
    st.error(f"Al Excel le faltan estas columnas: {', '.join(faltan)}")
    st.stop()

try:
    if es_faire:
        salida, av = core.generar_faire(erp, plantilla, marca, markup, retail, paises, pesos, pais_volum,
                                        solo_principal)
    else:
        salida, av = core.generar_ankorstore(erp, plantilla, marca, markup, retail, paises, pesos, pais_volum,
                                             iva, int(stock))
except ValueError as e:
    st.error(str(e))
    st.stop()

st.subheader("Resultado")
m = st.columns(5)
m[0].metric("Productos", av["total"])
m[1].metric("Sin foto", len(av["sin_imagen"]))
m[2].metric("Sin EAN", len(av["sin_gtin"]))
m[3].metric("Sin peso", len(av["sin_peso"]))
m[4].metric("Sin país", len(av["sin_pais"]))

etiquetas = {
    "sin_imagen": "📷 Sin foto", "sin_gtin": "🏷️ Sin EAN", "sin_peso": "⚖️ Sin peso (añade el prefijo en ⚙️)",
    "sin_pais": "🌍 Sin país (añade el proveedor en ⚙️)", "sin_precio": "💶 Sin precio", "sin_medidas": "📐 Sin medidas",
    "duplicadas": "♻️ Referencias repetidas (Ankorstore va sin variantes; se ha dejado la primera)",
}
hay_avisos = False
for clave, texto in etiquetas.items():
    if av.get(clave):
        hay_avisos = True
        st.warning(f"**{texto}** ({len(av[clave])}): {', '.join(map(str, av[clave]))}")
if not hay_avisos:
    st.success("Todo completo: fotos, EAN, peso, país, precio y medidas ✓")

st.download_button(
    f"⬇️ Descargar {marca}_{marketplace.lower()}.xlsx",
    salida,
    file_name=f"{marca}_{marketplace.lower()}.xlsx",
    type="primary",
)

with st.expander("Ver vista previa"):
    st.dataframe(av["vista"], hide_index=True, use_container_width=True)
