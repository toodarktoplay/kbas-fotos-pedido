# -*- coding: utf-8 -*-
"""
Fotos de pedido — Kbas Office
Sube el Excel del pedido (del ERP, refs con color) + las fotos de la temporada.
Devuelve un ZIP con todas las fotos de esas referencias: la principal
y todas sus variantes (_D, _LT, _H...), listas para enviar al cliente.
"""

import io
import os
import zipfile
import tempfile

import re

import pandas as pd
import pdfplumber
import streamlit as st

# ---------- Configuración ----------
REF_CANDIDATAS = [
    "REF.", "REF", "REFERENCIA", "REFERÈNCIA",
    "COD", "CODIGO", "CÓDIGO", "COD_ART", "COD.ART",
    "COD_ARTICULO", "CÓD. ARTÍCULO", "ARTICULO", "ARTÍCULO"
]
EXTS = (".jpg", ".jpeg", ".png", ".webp")


def norm_header(s):
    s = str(s).strip().upper()
    return (s.replace("Á", "A").replace("É", "E").replace("Í", "I")
            .replace("Ó", "O").replace("Ú", "U").replace("Ü", "U"))


def detectar_col_ref(df):
    mapa = {norm_header(c): c for c in df.columns}
    for cand in REF_CANDIDATAS:
        c = norm_header(cand)
        if c in mapa:
            return mapa[c]
    return None


def normalizar_ref(ref):
    if ref is None:
        return None
    s = str(ref).strip().upper()
    if not s or s == "NAN":
        return None
    return s.replace("/", "_").replace("-", "_").replace(" ", "")


def volcar_fotos(subidas, destino):
    for f in subidas:
        if f.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(f) as z:
                    z.extractall(destino)
            except zipfile.BadZipFile:
                st.error(f"El archivo {f.name} no es un ZIP válido.")
                st.stop()
        else:
            with open(os.path.join(destino, f.name), "wb") as d:
                d.write(f.getbuffer())


PATRON_REF = re.compile(r"\b[A-Z]{2}\d{4,8}_\d{1,3}\b")


def extraer_refs(archivo_bytes, nombre_archivo):
    """Saca las referencias del pedido, venga en PDF o en Excel."""
    if nombre_archivo.lower().endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(archivo_bytes)) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
        refs = list(dict.fromkeys(PATRON_REF.findall(texto.upper())))
        if not refs:
            raise ValueError("No encuentro referencias en el PDF. ¿Es el pedido del ERP?")
        return refs, "PDF"

    df = pd.read_excel(io.BytesIO(archivo_bytes))
    col_ref = detectar_col_ref(df)
    if not col_ref:
        raise ValueError(
            "No encuentro la columna de referencia. "
            f"Columnas del Excel: {', '.join(str(c) for c in df.columns)}"
        )
    refs = []
    for v in df[col_ref]:
        r = normalizar_ref(v)
        if r and r not in refs:
            refs.append(r)
    return refs, col_ref


def recopilar_fotos(archivo_bytes, nombre_archivo, carpeta):
    """Devuelve (zip_bytes, n_refs, n_fotos, sin_foto, origen)."""
    refs, origen = extraer_refs(archivo_bytes, nombre_archivo)

    # Índice de TODAS las fotos (aquí las variantes también cuentan)
    fotos = {}  # nombre_sin_ext_upper -> ruta
    for root, _, files in os.walk(carpeta):
        for f in files:
            if f.lower().endswith(EXTS):
                nombre = os.path.splitext(f)[0].strip().upper()
                fotos.setdefault(nombre, os.path.join(root, f))

    salida = io.BytesIO()
    n_fotos, sin_foto, ya_metidas = 0, [], set()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        for ref in refs:
            encontradas = [
                (nombre, ruta) for nombre, ruta in fotos.items()
                if nombre == ref or nombre.startswith(ref + "_")
            ]
            if not encontradas:
                sin_foto.append(ref)
                continue
            for nombre, ruta in sorted(encontradas):
                base = os.path.basename(ruta)
                if base.upper() in ya_metidas:
                    continue
                z.write(ruta, base)
                ya_metidas.add(base.upper())
                n_fotos += 1

    salida.seek(0)
    return salida, len(refs), n_fotos, sin_foto, origen


# ---------- Interfaz ----------
st.set_page_config(page_title="Fotos de pedido · Kbas Office", page_icon="📦")

st.title("Fotos de pedido")
st.caption("Sube el pedido (PDF del ERP o Excel) y las fotos de la temporada. Te devuelve un ZIP con todas las fotos de esas referencias (principal + detalles), listo para enviar al cliente.")

excel_subido = st.file_uploader(
    "1 · Pedido del ERP: el PDF directamente, o un Excel con columna de referencia",
    type=["pdf", "xlsx", "xlsm"],
)

fotos_subidas = st.file_uploader(
    "2 · Fotos de la temporada: ZIP con la carpeta, varios ZIP, fotos sueltas, o todo mezclado",
    type=["zip", "jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

if excel_subido and fotos_subidas:
    if st.button("Preparar fotos del pedido", type="primary"):
        with tempfile.TemporaryDirectory() as tmp:
            volcar_fotos(fotos_subidas, tmp)
            with st.spinner("Buscando y empaquetando fotos…"):
                try:
                    salida, n_refs, n_fotos, sin_foto, origen = recopilar_fotos(
                        excel_subido.getvalue(), excel_subido.name, tmp
                    )
                except ValueError as e:
                    st.error(str(e))
                    st.stop()

        st.success(f"Listo: {n_fotos} fotos de {n_refs} referencias del pedido ({origen}).")
        if sin_foto:
            with st.expander(f"⚠ {len(sin_foto)} referencias sin ninguna foto"):
                st.text("\n".join(sin_foto))

        nombre = os.path.splitext(excel_subido.name)[0] + "_fotos.zip"
        st.download_button(
            "⬇ Descargar " + nombre,
            data=salida,
            file_name=nombre,
            mime="application/zip",
        )
else:
    st.info("Sube el pedido y al menos una foto o ZIP para empezar.")
