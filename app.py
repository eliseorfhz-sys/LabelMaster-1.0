import io
import json
import os
import re
import sqlite3
import textwrap
from datetime import datetime
from typing import Any, Optional

import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode.code128 import Code128 as Code128RL
from barcode import Code128 as Code128Img
from barcode.writer import ImageWriter

COLUMNAS_ESPERADAS = [
    "Código",
    "Producto",
    "P. Costo",
    "P. Venta",
    "P. Mayoreo",
    "Departamento",
    "Existencia",
    "Inv. Mínimo",
    "Inv. Máximo",
    "Tipo de Venta",
    "Proveedor",
]

COLS_ETIQUETA = 4
ROWS_ETIQUETA = 8
# Factor <1 reduce solo la altura visual del Code128 (1 − 0.15 = 0.85)
ESCALA_ALTURA_CODIGO_BARRAS = 0.85
# Máximo de líneas para el nombre en cada etiqueta (PDF)
MAX_LINEAS_NOMBRE_ETIQUETA = 4

DIR_APP = os.path.dirname(os.path.abspath(__file__))
RUTA_BD = os.path.join(DIR_APP, "catalogo.db")

MAP_COLUMNAS_A_BD = {
    "Código": "codigo",
    "Producto": "producto",
    "P. Costo": "p_costo",
    "P. Venta": "p_venta",
    "P. Mayoreo": "p_mayoreo",
    "Departamento": "departamento",
    "Existencia": "existencia",
    "Inv. Mínimo": "inv_minimo",
    "Inv. Máximo": "inv_maximo",
    "Tipo de Venta": "tipo_venta",
    "Proveedor": "proveedor",
}
MAP_COLUMNAS_DESDE_BD = {v: k for k, v in MAP_COLUMNAS_A_BD.items()}

COLUMNAS_NUMERICAS = [
    "P. Costo",
    "P. Venta",
    "P. Mayoreo",
    "Existencia",
    "Inv. Mínimo",
    "Inv. Máximo",
]

COLUMNAS_CUPONES_ESPERADAS = [
    "Código",
    "Validez",
    "Tipo de Cuenta",
    "Validación",
    "Precio",
    "SSID",
]

CODIGO_BARRAS_CUPON_1H = "1H0525042026"
CODIGO_BARRAS_CUPON_4H = "4H1025042026"
COLS_CUPON = 3
ROWS_CUPON = 7
TAM_GRUPO_CUPONES = 42


def _registrar_fuente_unicode() -> tuple[str, str]:
    """Helvetica no cubre bien acentos; usa Arial en Windows si existe."""
    arial = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf")
    arial_b = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arialbd.ttf")
    if os.path.isfile(arial):
        if "ArialEtiquetas" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("ArialEtiquetas", arial))
        if os.path.isfile(arial_b) and "ArialEtiquetas-Bold" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("ArialEtiquetas-Bold", arial_b))
        return "ArialEtiquetas", "ArialEtiquetas-Bold"
    return "Helvetica", "Helvetica-Bold"


def normalizar_nombres_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def cargar_excel(archivo) -> pd.DataFrame:
    ext = archivo.name.lower().split(".")[-1]
    buf = io.BytesIO(archivo.getvalue())
    if ext == "xlsx":
        return pd.read_excel(buf, engine="openpyxl")
    if ext == "xls":
        return pd.read_excel(buf, engine="xlrd")
    return pd.read_excel(buf)


def precio_a_float(valor: Any) -> Optional[float]:
    """Interpreta precios desde Excel/BD: números, '$12.50', '1.234,56', etc."""
    if valor is None:
        return None
    if isinstance(valor, (int, np.integer)):
        return float(int(valor))
    if isinstance(valor, (float, np.floating)):
        if pd.isna(valor) or (isinstance(valor, float) and np.isnan(valor)):
            return None
        return float(valor)
    s = str(valor).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    s = re.sub(r"[\s$€]", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(mxn|usd)\b", "", s, flags=re.IGNORECASE).strip()
    tiene_coma = "," in s
    tiene_punto = "." in s
    if tiene_coma and tiene_punto:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif tiene_coma:
        partes = s.split(",")
        if len(partes) == 2 and len(partes[1]) <= 2 and partes[1].isdigit():
            s = partes[0].replace(".", "") + "." + partes[1]
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def codigo_para_etiqueta(valor: Any) -> str:
    """Código legible y estable para barras / búsqueda (quita .0 de Excel/SQLite)."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return ""
    if pd.isna(valor):
        return ""
    if isinstance(valor, (int, np.integer)):
        return str(int(valor))
    if isinstance(valor, (float, np.floating)):
        f = float(valor)
        if np.isnan(f):
            return ""
        if f.is_integer():
            return str(int(round(f)))
        s = ("%s" % f).rstrip("0").rstrip(".")
        return s
    s = str(valor).strip()
    if not s or s.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s


def coercer_numericas_catalogo(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in COLUMNAS_NUMERICAS:
        if col not in df.columns:
            continue

        def celda(x: Any):
            if x is None:
                return np.nan
            if isinstance(x, float) and np.isnan(x):
                return np.nan
            if pd.isna(x):
                return np.nan
            p = precio_a_float(x)
            return np.nan if p is None else p

        df[col] = df[col].map(celda)
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def preparar_dataframe_catalogo(df: pd.DataFrame) -> pd.DataFrame:
    """Deja solo las columnas esperadas (rellena faltantes)."""
    df = normalizar_nombres_columnas(df)
    out = pd.DataFrame()
    for col in COLUMNAS_ESPERADAS:
        if col in df.columns:
            out[col] = df[col]
        else:
            out[col] = pd.NA
    out = coercer_numericas_catalogo(out)
    if "Código" in out.columns:
        out = out.drop_duplicates(subset=["Código"], keep="last")
    return out.reset_index(drop=True)


def guardar_catalogo_en_bd(df: pd.DataFrame) -> int:
    """Reemplaza la tabla de productos en SQLite. Devuelve filas guardadas."""
    datos = preparar_dataframe_catalogo(df)
    bd = datos.rename(columns=MAP_COLUMNAS_A_BD)
    conn = sqlite3.connect(RUTA_BD)
    try:
        bd.to_sql("productos", conn, if_exists="replace", index=False)
    finally:
        conn.close()
    return len(datos)


def cargar_catalogo_desde_bd() -> pd.DataFrame:
    """Lee el catálogo desde SQLite; si no hay BD o tabla, DataFrame vacío con columnas."""
    if not os.path.isfile(RUTA_BD):
        return pd.DataFrame(columns=COLUMNAS_ESPERADAS)
    conn = sqlite3.connect(RUTA_BD)
    try:
        cur = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='productos' LIMIT 1"
        )
        if cur.fetchone() is None:
            return pd.DataFrame(columns=COLUMNAS_ESPERADAS)
        raw = pd.read_sql("SELECT * FROM productos", conn)
    finally:
        conn.close()
    raw = raw.rename(columns=MAP_COLUMNAS_DESDE_BD)
    for col in COLUMNAS_ESPERADAS:
        if col not in raw.columns:
            raw[col] = pd.NA
    raw = raw[COLUMNAS_ESPERADAS]
    raw = coercer_numericas_catalogo(raw)
    return raw.reset_index(drop=True)


def inicializar_bd_cupones() -> None:
    conn = sqlite3.connect(RUTA_BD)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS cupones_grupos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre_grupo TEXT NOT NULL,
                tipo_cupon TEXT NOT NULL,
                codigo_barras TEXT NOT NULL,
                total_items INTEGER NOT NULL,
                creado_en TEXT NOT NULL,
                datos_json TEXT NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def preparar_dataframe_cupones(df: pd.DataFrame) -> pd.DataFrame:
    df = normalizar_nombres_columnas(df)
    out = pd.DataFrame()
    for col in COLUMNAS_CUPONES_ESPERADAS:
        if col in df.columns:
            out[col] = df[col]
        else:
            out[col] = pd.NA

    if "Código" in out.columns:
        out["Código"] = out["Código"].map(codigo_para_etiqueta)
    if "Precio" in out.columns:
        out["Precio"] = out["Precio"].map(formato_precio)
    for col in ["Validez", "Tipo de Cuenta", "Validación", "SSID"]:
        out[col] = out[col].map(lambda x: "" if pd.isna(x) else str(x).strip())

    out = out.replace({np.nan: "", pd.NA: ""})
    return out.reset_index(drop=True)


def guardar_grupo_cupones_en_bd(
    *,
    nombre_grupo: str,
    tipo_cupon: str,
    codigo_barras: str,
    items: list[dict[str, str]],
) -> int:
    conn = sqlite3.connect(RUTA_BD)
    try:
        cur = conn.execute(
            """
            INSERT INTO cupones_grupos
            (nombre_grupo, tipo_cupon, codigo_barras, total_items, creado_en, datos_json)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                nombre_grupo.strip() or f"Grupo {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                tipo_cupon,
                codigo_barras,
                len(items),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                json.dumps(items, ensure_ascii=False),
            ),
        )
        conn.commit()
        return int(cur.lastrowid)
    finally:
        conn.close()


def cargar_grupos_cupones_desde_bd() -> pd.DataFrame:
    if not os.path.isfile(RUTA_BD):
        return pd.DataFrame(
            columns=["id", "nombre_grupo", "tipo_cupon", "codigo_barras", "total_items", "creado_en"]
        )
    conn = sqlite3.connect(RUTA_BD)
    try:
        cur = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='cupones_grupos' LIMIT 1"
        )
        if cur.fetchone() is None:
            return pd.DataFrame(
                columns=["id", "nombre_grupo", "tipo_cupon", "codigo_barras", "total_items", "creado_en"]
            )
        return pd.read_sql(
            """
            SELECT id, nombre_grupo, tipo_cupon, codigo_barras, total_items, creado_en
            FROM cupones_grupos
            ORDER BY id DESC
            """,
            conn,
        )
    finally:
        conn.close()


def cargar_items_de_grupo(grupo_id: int) -> list[dict[str, str]]:
    conn = sqlite3.connect(RUTA_BD)
    try:
        cur = conn.execute(
            "SELECT datos_json FROM cupones_grupos WHERE id = ? LIMIT 1",
            (int(grupo_id),),
        )
        row = cur.fetchone()
        if row is None:
            return []
        try:
            raw = json.loads(row[0] or "[]")
        except Exception:
            return []
        if not isinstance(raw, list):
            return []
        out: list[dict[str, str]] = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            out.append(
                {
                    "Código": str(item.get("Código", "") or "").strip(),
                    "Validez": str(item.get("Validez", "") or "").strip(),
                    "Tipo de Cuenta": str(item.get("Tipo de Cuenta", "") or "").strip(),
                    "Validación": str(item.get("Validación", "") or "").strip(),
                    "Precio": str(item.get("Precio", "") or "").strip(),
                    "SSID": str(item.get("SSID", "") or "").strip(),
                }
            )
        return out
    finally:
        conn.close()


def generar_pdf_cupones_internet(
    items: list[dict[str, str]],
    *,
    codigo_barras_grupo: str,
    margen_in: float = 0.22,
    separacion_in: float = 0.05,
) -> bytes:
    font, font_bold = _registrar_fuente_unicode()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    page_w, page_h = letter
    margin = margen_in * inch
    gap = separacion_in * inch
    usable_w = page_w - 2 * margin
    usable_h = page_h - 2 * margin
    label_w = (usable_w - gap * (COLS_CUPON - 1)) / COLS_CUPON
    label_h = (usable_h - gap * (ROWS_CUPON - 1)) / ROWS_CUPON
    por_pagina = COLS_CUPON * ROWS_CUPON
    pad = 5

    def dibujar_un_cupon(x_left: float, y_bottom: float, item: dict[str, str]) -> None:
        c.setStrokeGray(0.32)
        c.setLineWidth(0.35)
        c.roundRect(x_left, y_bottom, label_w, label_h, 2.5, stroke=1, fill=0)

        x = x_left + pad
        # Baja un poco más el bloque de texto para aprovechar espacio antes del código de barras.
        y_top = y_bottom + label_h - pad - 17
        ancho = label_w - 2 * pad

        codigo_txt = str(item.get("Código", "") or "").strip() or "—"
        validez_txt = str(item.get("Validez", "") or "").strip()
        tipo_txt = str(item.get("Tipo de Cuenta", "") or "").strip()
        valida_txt = str(item.get("Validación", "") or "").strip()
        precio_txt = str(item.get("Precio", "") or "").strip() or "$0.00"
        ssid_txt = str(item.get("SSID", "") or "").strip()

        x_centro = x_left + (label_w / 2)

        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_bold, 15.0)
        linea_codigo = _texto_ajustar_ancho(f"Código: {codigo_txt}", font_bold, 15.0, ancho)
        c.drawCentredString(x_centro, y_top, linea_codigo)

        c.setFont(font, 9.2)
        y = y_top - 14.0
        campos = [
            ("Validez", validez_txt),
            ("Tipo", tipo_txt),
            ("Validación", valida_txt),
            ("Precio", precio_txt),
            ("SSID", ssid_txt),
        ]
        for nombre, valor in campos:
            linea = f"{nombre}: {valor}" if valor else f"{nombre}: —"
            linea_fit = _texto_ajustar_ancho(linea, font, 9.2, ancho)
            c.drawCentredString(x_centro, y, linea_fit)
            y -= 10.5

        by = y_bottom + 4
        png = codigo128_a_png(
            codigo_barras_grupo,
            angosto=True,
            module_width_mult=1.30,
            module_height_mult=1.14,
        )
        if png is not None:
            dibujar_imagen_barras(c, png, x, by, ancho, 15.8)

    for i, item in enumerate(items):
        if i > 0 and i % por_pagina == 0:
            c.showPage()
        slot = i % por_pagina
        row = slot // COLS_CUPON
        col = slot % COLS_CUPON
        x_left = margin + col * (label_w + gap)
        y_bottom = page_h - margin - (row + 1) * label_h - row * gap
        dibujar_un_cupon(x_left, y_bottom, item)

    c.save()
    buf.seek(0)
    return buf.getvalue()


def _texto_ajustar_ancho(texto: str, nombre_fuente: str, tam_pt: float, ancho_max_pt: float) -> str:
    """Acorta con … si el texto en esa fuente supera el ancho (evita desbordes al centrar)."""
    if not texto or ancho_max_pt <= 0:
        return texto or ""
    if pdfmetrics.stringWidth(texto, nombre_fuente, tam_pt) <= ancho_max_pt:
        return texto
    suf = "…"
    t = texto
    while len(t) > 0 and pdfmetrics.stringWidth(t + suf, nombre_fuente, tam_pt) > ancho_max_pt:
        t = t[:-1]
    return (t + suf) if t else suf


def formato_precio(valor: Any) -> str:
    n = precio_a_float(valor)
    if n is None:
        return "$0.00"
    return f"${n:,.2f}"


def codigo128_a_png(
    codigo: str,
    *,
    angosto: bool,
    module_width_mult: float = 1.0,
    module_height_mult: float = 1.0,
) -> Optional[io.BytesIO]:
    """Genera PNG Code128 (más fiable en PDF que el widget nativo de ReportLab)."""
    cod = str(codigo).strip()
    if not cod:
        return None
    base_h = 6.0 if angosto else 7.5
    mh = base_h * 6.0 * 0.4 * ESCALA_ALTURA_CODIGO_BARRAS * max(0.5, module_height_mult)
    if COLS_ETIQUETA >= 4:
        # Celda ~1/4 del ancho carta: módulos un poco más finos para que al escalar no quede todo por ancho
        base_w = 0.082
        mw = base_w * 3.0 * 3.2
    else:
        base_w = 0.09 if angosto else 0.11
        mw = base_w * 3.0 * 4.0
    mw = mw * max(0.5, module_width_mult)
    opts = {
        "module_width": mw,
        "module_height": mh,
        "quiet_zone": 2.0,
        "font_size": 0,
        "text_distance": 0,
    }
    try:
        buf = io.BytesIO()
        Code128Img(cod, writer=ImageWriter()).write(buf, options=opts)
        buf.seek(0)
        return buf
    except Exception:
        return None


def dibujar_imagen_barras(
    c: canvas.Canvas,
    png_buf: io.BytesIO,
    x_left: float,
    y_bottom: float,
    ancho_max: float,
    alto_max: float,
) -> float:
    """Coloca la imagen centrada; devuelve altura usada en puntos."""
    img = ImageReader(png_buf)
    iw, ih = img.getSize()
    if iw <= 0 or ih <= 0:
        return 0.0
    escala = min(ancho_max / iw, alto_max / ih)
    w = iw * escala
    h = ih * escala
    x = x_left + max(0.0, (ancho_max - w) / 2)
    c.drawImage(img, x, y_bottom, width=w, height=h, mask="auto")
    return h


def fila_desde_bd_para_etiqueta(row: pd.Series) -> tuple[str, str, str]:
    """Nombre, precio formateado y código para Code128 leyendo la fila del catálogo/BD."""
    nombre = str(row.get("Producto", "") or "").strip() or "—"
    precio_txt = formato_precio(row.get("P. Venta"))
    codigo = codigo_para_etiqueta(row.get("Código"))
    return nombre, precio_txt, codigo


def generar_pdf_etiquetas(
    items: list[tuple[str, str, str]],
    *,
    margen_in: float = 0.22,
    separacion_in: float = 0.038,
) -> bytes:
    """
    items: lista de (nombre_producto, precio_texto, codigo_barras)
    Hoja carta, varias etiquetas estilo góndola (nombre, precio, Code128).
    """
    font, font_bold = _registrar_fuente_unicode()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    page_w, page_h = letter
    margin = margen_in * inch
    gap = separacion_in * inch
    usable_w = page_w - 2 * margin
    usable_h = page_h - 2 * margin
    label_w = (usable_w - gap * (COLS_ETIQUETA - 1)) / COLS_ETIQUETA
    label_h = (usable_h - gap * (ROWS_ETIQUETA - 1)) / ROWS_ETIQUETA
    por_pagina = COLS_ETIQUETA * ROWS_ETIQUETA
    pad = 3 if COLS_ETIQUETA >= 4 else 4
    angosto = COLS_ETIQUETA >= 3
    cuatro_cols = COLS_ETIQUETA >= 4

    def dibujar_una(
        x_left: float,
        y_bottom: float,
        nombre: str,
        precio_txt: str,
        codigo: str,
    ) -> None:
        c.setStrokeGray(0.32)
        c.setLineWidth(0.35)
        r_radio = 2.5 if cuatro_cols else 3
        c.roundRect(x_left, y_bottom, label_w, label_h, r_radio, stroke=1, fill=0)
        c.saveState()
        clip = c.beginPath()
        clip.rect(x_left + 0.5, y_bottom + 0.5, label_w - 1.0, label_h - 1.0)
        c.clipPath(clip, stroke=0, fill=0)

        y_top = y_bottom + label_h
        nombre_limpio = str(nombre).strip() or "—"
        cod = str(codigo).strip()
        ancho_disp = max(8.0, label_w - 2 * pad)

        # Abajo: barras (PNG); 4 columnas = celda más baja en proporción al texto
        by = y_bottom + 2 if cuatro_cols else y_bottom + 3
        if cuatro_cols:
            reserva_superior = 19.0
        else:
            reserva_superior = 22.0 if angosto else 26.0
        _alto = max(40.0, label_h - reserva_superior - (by - y_bottom))
        _alto = min(_alto, label_h * 0.82)
        factor_alto_bc = 0.30 if cuatro_cols else (0.34 if angosto else 0.38)
        alto_bc_max = max(11.0, _alto * factor_alto_bc) * ESCALA_ALTURA_CODIGO_BARRAS
        bloque_bc = 4.0
        png = codigo128_a_png(cod, angosto=angosto) if cod else None
        if png is not None:
            h_usada = dibujar_imagen_barras(c, png, x_left + pad, by, ancho_disp, alto_bc_max)
            bloque_bc = max(bloque_bc, h_usada + 2.0)
        elif cod:
            try:
                bc_h = min(alto_bc_max * 0.88, alto_bc_max)
                bw_lim = 0.045 if cuatro_cols else 0.055
                bw_obj = max(0.016, min(bw_lim, (ancho_disp / max(len(cod) * 9, 1)) * 2.6))
                bc = Code128RL(cod, barHeight=bc_h, barWidth=bw_obj, humanReadable=0)
                bw = bc.width
                bx = x_left + pad + max(0.0, (ancho_disp - bw) / 2)
                c.setFillColorRGB(0, 0, 0)
                c.setStrokeColorRGB(0, 0, 0)
                bc.drawOn(c, bx, by)
                bloque_bc = bc_h + 4.0
            except Exception:
                c.setFont(font, 5 if cuatro_cols else 6)
                c.setFillColorRGB(0, 0, 0)
                c.drawCentredString(x_left + label_w / 2, by + 1, f"Cód: {cod[:20]}")

        y_precio = by + bloque_bc + (1.5 if cuatro_cols else 2.0)
        y_nombre_ceiling = y_top - pad
        if cuatro_cols:
            tam_precio, tam_nombre, lh = 21.06, 9.8, 11.6
        elif angosto:
            tam_precio, tam_nombre, lh = 24.3, 11.2, 13.2
        else:
            tam_precio, tam_nombre, lh = 27.54, 12.5, 14.8
        sep_nom_precio = max(16.0, tam_precio * 0.92 + 7.0)
        ancho_ut = max(10.0, label_w - 2 * pad)
        ancho_chars = max(8, min(36, int(ancho_ut / (tam_nombre * 0.65))))
        y_primera = y_nombre_ceiling - int(tam_nombre * 0.88)
        y_suelo = y_precio + tam_precio + 3
        hueco_lineas = y_primera - y_suelo
        if hueco_lineas <= 0:
            max_lineas = 1
        else:
            max_lineas = int(hueco_lineas // lh) + 1
            max_lineas = max(1, min(MAX_LINEAS_NOMBRE_ETIQUETA, max_lineas))

        x_centro = x_left + label_w / 2
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_bold, tam_precio)
        precio_dib = _texto_ajustar_ancho(precio_txt, font_bold, tam_precio, ancho_ut)
        c.drawCentredString(x_centro, y_precio, precio_dib)

        c.setFont(font, tam_nombre)
        y_linea = y_nombre_ceiling - int(tam_nombre * 0.92)
        margen_sobre_precio = tam_precio + 4
        for linea in textwrap.wrap(nombre_limpio, width=ancho_chars)[:max_lineas]:
            if y_linea < y_precio + margen_sobre_precio:
                break
            linea_fit = _texto_ajustar_ancho(linea, font, tam_nombre, ancho_ut)
            c.drawCentredString(x_centro, y_linea, linea_fit)
            y_linea -= lh

        c.restoreState()

    for i, (nombre, precio_txt, codigo) in enumerate(items):
        if i > 0 and i % por_pagina == 0:
            c.showPage()
        slot = i % por_pagina
        row = slot // COLS_ETIQUETA
        col = slot % COLS_ETIQUETA
        x_left = margin + col * (label_w + gap)
        y_bottom = page_h - margin - (row + 1) * label_h - row * gap
        dibujar_una(x_left, y_bottom, nombre, precio_txt, codigo)

    c.save()
    buf.seek(0)
    return buf.getvalue()


st.set_page_config(page_title="LabelMaster-1.0", layout="wide")
# Oculta controles de Streamlit en la esquina superior derecha (Deploy y botones vecinos).
st.markdown(
    """
    <style>
    /* Mantener visible el menú/hamburguesa lateral */
    #MainMenu {visibility: visible !important;}
    /* Ocultar solo acciones de deploy/compartir del header */
    button[title*="Deploy"],
    button[aria-label*="Deploy"],
    [data-testid="stAppDeployButton"] {
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("LabelMaster-1.0")

msg = st.session_state.pop("_import_flash", None)
if msg:
    st.sidebar.success(msg)

st.sidebar.subheader("Catálogo")
st.sidebar.caption(
    "El catálogo se guarda en una base de datos local (catalogo.db). "
    "Importa un Excel para crear o reemplazar todo el catálogo. Columnas: Código, Producto, "
    "P. Costo, P. Venta, P. Mayoreo, Departamento, Existencia, Inv. Mínimo, Inv. Máximo, "
    "Tipo de Venta, Proveedor."
)
archivo = st.sidebar.file_uploader(
    "Importar / actualizar catálogo desde Excel (.xlsx o .xls)",
    type=["xlsx", "xls"],
    help="Sustituye el contenido de la base de datos por las filas del archivo (un código por producto; duplicados: gana la última fila).",
)

if archivo is not None:
    clave = (archivo.name, archivo.size)
    if st.session_state.get("_archivo_clave") != clave:
        try:
            df_import = normalizar_nombres_columnas(cargar_excel(archivo))
            n = guardar_catalogo_en_bd(df_import)
            st.session_state["_archivo_clave"] = clave
            st.session_state.pop("pdf_etiquetas", None)
            st.session_state["_import_flash"] = (
                f"Catálogo guardado en la base de datos: {n} producto(s) desde «{archivo.name}»."
            )
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo leer o guardar el archivo: {e}")

df = cargar_catalogo_desde_bd()
inicializar_bd_cupones()

st.sidebar.caption("Base de datos (SQLite)")
st.sidebar.text(RUTA_BD)
st.sidebar.metric("Productos guardados", len(df))
if len(df) == 0:
    st.info(
        "No hay productos en la base de datos. Sube un archivo Excel para importar el catálogo."
    )

faltantes = [c for c in COLUMNAS_ESPERADAS if c not in df.columns]
if faltantes:
    st.warning(
        "Faltan columnas esperadas (revisa mayúsculas y tildes): "
        + ", ".join(faltantes)
    )

col_categoria = "Departamento"
if col_categoria not in df.columns:
    st.error(
        f"No existe la columna «{col_categoria}». "
        "Columnas encontradas: " + ", ".join(df.columns.astype(str).tolist())
    )
    st.stop()

df_trabajo = df.reset_index(drop=True)

tab_catalogo, tab_etiquetas, tab_cupones = st.tabs(["Catálogo", "Etiquetas", "Cupones Internet"])

with tab_catalogo:
    df_trabajo[col_categoria] = (
        df_trabajo[col_categoria].astype(str).replace("nan", "").str.strip()
    )
    categorias = sorted(
        df_trabajo[col_categoria].replace("", pd.NA).dropna().unique().tolist()
    )
    categorias_con_todas = ["Todas"] + categorias

    st.caption("Búsqueda por código de barras (lector USB): al escanear y enviar Enter, busca automáticamente.")
    with st.form("form_busqueda_catalogo_scan", clear_on_submit=True):
        codigo_scan_catalogo = st.text_input(
            "Escanear o escribir código para buscar",
            placeholder="Enfoca aquí y escanea…",
            key="catalogo_scan_input",
        )
        enviado_scan_catalogo = st.form_submit_button("Buscar (si escribes a mano)")
    if enviado_scan_catalogo:
        cod_norm = codigo_para_etiqueta(codigo_scan_catalogo) or codigo_scan_catalogo.strip()
        st.session_state["_catalogo_codigo_busqueda"] = cod_norm
        st.rerun()

    categoria_sel = st.selectbox(
        "Filtrar por categoría (Departamento)",
        options=categorias_con_todas,
        index=0,
        key="filtro_depto",
    )

    if categoria_sel == "Todas":
        filtrado = df_trabajo
    else:
        filtrado = df_trabajo[df_trabajo[col_categoria] == categoria_sel]

    codigo_busq = (st.session_state.get("_catalogo_codigo_busqueda", "") or "").strip()
    if codigo_busq:
        mask_cod = df_trabajo["Código"].map(codigo_para_etiqueta) == codigo_busq
        por_codigo = df_trabajo[mask_cod]
        if por_codigo.empty:
            por_codigo = df_trabajo[
                df_trabajo["Código"].astype(str).str.strip() == codigo_busq
            ]
        if por_codigo.empty:
            st.warning(f"No se encontró producto con código: {codigo_busq}")
        else:
            idxs = set(por_codigo.index.tolist())
            filtrado = filtrado[filtrado.index.isin(idxs)]
            st.info(f"Búsqueda por código activa: {codigo_busq}")
        if st.button("Limpiar búsqueda por código", key="btn_limpiar_busq_catalogo"):
            st.session_state.pop("_catalogo_codigo_busqueda", None)
            st.rerun()

    st.metric("Productos mostrados", len(filtrado))
    st.dataframe(filtrado, use_container_width=True, hide_index=True)

with tab_etiquetas:
    req = ["Código", "Producto", "P. Venta"]
    if not all(c in df_trabajo.columns for c in req):
        st.error("Para etiquetas se requieren las columnas: Código, Producto, P. Venta.")
        st.stop()

    opciones_map: dict[str, int] = {}
    for idx, row in df_trabajo.iterrows():
        cod = codigo_para_etiqueta(row["Código"]) or "—"
        nom = str(row["Producto"]).strip()[:80]
        etiqueta = f"{cod}  —  {nom}"
        if etiqueta in opciones_map:
            etiqueta = f"{etiqueta}  (fila {idx})"
        opciones_map[etiqueta] = int(idx)

    def _procesar_codigo_escaneo(cod_bus: str) -> None:
        cod_bus = (cod_bus or "").strip()
        if not cod_bus:
            st.session_state["_scan_etiqueta_msg"] = "warning:Escribe un código."
            return
        norm = codigo_para_etiqueta(cod_bus) or cod_bus.strip()
        mask = df_trabajo["Código"].map(codigo_para_etiqueta) == norm
        coinciden = df_trabajo[mask]
        if coinciden.empty:
            coinciden = df_trabajo[
                df_trabajo["Código"].astype(str).str.strip() == cod_bus.strip()
            ]
        if coinciden.empty:
            st.session_state["_scan_etiqueta_msg"] = "error:No hay producto con ese código."
            return
        idx = int(coinciden.index[0])
        lbl = next((k for k, v in opciones_map.items() if v == idx), None)
        if lbl is None:
            st.session_state["_scan_etiqueta_msg"] = (
                "error:No se pudo resolver la etiqueta del producto."
            )
            return
        st.session_state["_etiqueta_agregar_tras_escaneo"] = lbl

    msg_scan = st.session_state.pop("_scan_etiqueta_msg", None)
    if msg_scan:
        if msg_scan.startswith("error:"):
            st.error(msg_scan[6:])
        elif msg_scan.startswith("warning:"):
            st.warning(msg_scan[8:])

    # Selección persistente independiente del widget para evitar reemplazos entre escaneos.
    if "seleccion_etiquetas" not in st.session_state:
        st.session_state["seleccion_etiquetas"] = list(
            st.session_state.get("multiselect_etiquetas", [])
        )

    lbl_pendiente = st.session_state.pop("_etiqueta_agregar_tras_escaneo", None)
    if lbl_pendiente is not None:
        if lbl_pendiente in opciones_map:
            actuales = list(
                st.session_state.get(
                    "seleccion_etiquetas",
                    st.session_state.get("multiselect_etiquetas", []),
                )
            )
            if lbl_pendiente not in actuales:
                actuales.append(lbl_pendiente)
            # Filtrar por opciones vigentes para evitar valores huérfanos.
            actuales = [x for x in actuales if x in opciones_map]
            st.session_state["seleccion_etiquetas"] = actuales
            st.session_state["multiselect_etiquetas"] = actuales
        else:
            st.error("No se pudo resolver la etiqueta del producto.")

    seleccion_vigente = [
        x for x in st.session_state.get("seleccion_etiquetas", []) if x in opciones_map
    ]
    st.session_state["seleccion_etiquetas"] = seleccion_vigente
    st.session_state["multiselect_etiquetas"] = seleccion_vigente

    st.caption("Con **lector USB**: enfoca el campo y escanea; el Enter del lector envía y agrega solo.")
    with st.form("form_escaneo_codigo", clear_on_submit=True):
        escaneo = st.text_input(
            "Escanear o escribir código",
            placeholder="Enfoca aquí y escanea…",
        )
        enviado = st.form_submit_button("Agregar (si escribes a mano)")
    if enviado:
        _procesar_codigo_escaneo(escaneo)
        st.rerun()

    elegidos = st.multiselect(
        "Productos agregados (escaneados)",
        options=seleccion_vigente,
        key="multiselect_etiquetas",
        help="Aquí solo se listan los productos ya agregados por código. Puedes quitar alguno desmarcándolo.",
    )
    st.session_state["seleccion_etiquetas"] = list(elegidos)

    copias = st.number_input("Copias por producto (cada uno se repite en la hoja)", min_value=1, max_value=50, value=1)

    indices_sel = [opciones_map[k] for k in elegidos if k in opciones_map]

    if st.button("Generar PDF de etiquetas", type="primary", key="btn_pdf"):
        if not indices_sel:
            st.warning("Selecciona al menos un producto.")
        else:
            items: list[tuple[str, str, str]] = []
            for ix in indices_sel:
                r = df_trabajo.loc[ix]
                nombre, precio, codigo = fila_desde_bd_para_etiqueta(r)
                for _ in range(int(copias)):
                    items.append((nombre, precio, codigo))
            try:
                st.session_state["pdf_etiquetas"] = generar_pdf_etiquetas(items)
                st.session_state["pdf_etiquetas_meta"] = len(items)
                st.success(
                    f"Listo: {len(items)} etiqueta(s) en hojas carta "
                    f"({COLS_ETIQUETA}×{ROWS_ETIQUETA} por hoja)."
                )
            except Exception as e:
                st.error(f"No se pudo generar el PDF: {e}")

    pdf_guardado = st.session_state.get("pdf_etiquetas")
    if pdf_guardado:
        n = st.session_state.get("pdf_etiquetas_meta", 0)
        st.download_button(
            label=f"Descargar etiquetas.pdf ({n} etiqueta(s))",
            data=pdf_guardado,
            file_name="etiquetas.pdf",
            mime="application/pdf",
            key="dl_etiquetas",
        )

    st.caption(
        f"Diseño: hoja carta con {COLS_ETIQUETA} columnas × {ROWS_ETIQUETA} filas. "
        "Los datos salen de la base de datos (columnas Producto, P. Venta, Código). "
        "Si antes veías $0.00, vuelve a importar el Excel: ahora se interpretan precios con símbolo $ o formato con comas. "
        "En impresión del PDF usa escala 100 % o «tamaño real»."
    )
    st.markdown("**Vista previa — productos en la selección de etiquetas**")
    if not indices_sel:
        st.info(
            "Aquí verás el listado al elegir productos en la lista o al escanear códigos. "
            "Aún no hay ninguno seleccionado."
        )
    else:
        df_prev = df_trabajo.loc[indices_sel, ["Código", "Producto", "P. Venta"]].copy()
        df_prev["P. Venta"] = df_prev["P. Venta"].map(formato_precio)
        df_prev.insert(0, "#", range(1, len(df_prev) + 1))
        st.dataframe(
            df_prev,
            use_container_width=True,
            hide_index=True,
            column_config={
                "#": st.column_config.NumberColumn("Orden", width="small"),
                "Código": st.column_config.TextColumn("Código"),
                "Producto": st.column_config.TextColumn("Producto"),
                "P. Venta": st.column_config.TextColumn("P. Venta"),
            },
        )
        st.caption(
            f"Total: **{len(indices_sel)}** producto(s) distintos · "
            f"**{len(indices_sel) * int(copias)}** etiqueta(s) físicas con copias actuales."
        )

with tab_cupones:
    st.subheader("Fichas de Internet por grupos")
    st.caption(
        "Carga un Excel con 42 elementos por grupo y genera fichas en plantilla carta "
        f"{ROWS_CUPON}×{COLS_CUPON}. Columnas requeridas: "
        + ", ".join(COLUMNAS_CUPONES_ESPERADAS)
        + "."
    )

    tipo_cupon = st.radio(
        "Tipo de cupón",
        options=["1 hora", "4 horas"],
        horizontal=True,
    )
    codigo_barras_grupo = (
        CODIGO_BARRAS_CUPON_1H if tipo_cupon == "1 hora" else CODIGO_BARRAS_CUPON_4H
    )
    st.info(f"Código de barras fijo para este tipo: **{codigo_barras_grupo}**")

    nombre_grupo = st.text_input(
        "Nombre del grupo",
        placeholder="Ejemplo: Turno mañana 25/04",
        key="cupones_nombre_grupo",
    )
    archivo_cupones = st.file_uploader(
        "Cargar Excel de cupones (.xlsx o .xls)",
        type=["xlsx", "xls"],
        key="cupones_excel_uploader",
    )

    if st.button("Validar y guardar grupo de 42", type="primary", key="btn_guardar_grupo_cupones"):
        if archivo_cupones is None:
            st.warning("Primero carga el archivo Excel del grupo.")
        else:
            try:
                df_cupon = preparar_dataframe_cupones(cargar_excel(archivo_cupones))
                if len(df_cupon) != TAM_GRUPO_CUPONES:
                    st.error(
                        f"El archivo debe traer exactamente {TAM_GRUPO_CUPONES} filas. "
                        f"Recibido: {len(df_cupon)}."
                    )
                else:
                    faltan = [c for c in COLUMNAS_CUPONES_ESPERADAS if c not in df_cupon.columns]
                    if faltan:
                        st.error("Faltan columnas requeridas: " + ", ".join(faltan))
                    else:
                        items = df_cupon[COLUMNAS_CUPONES_ESPERADAS].to_dict(orient="records")
                        nuevo_id = guardar_grupo_cupones_en_bd(
                            nombre_grupo=nombre_grupo,
                            tipo_cupon=tipo_cupon,
                            codigo_barras=codigo_barras_grupo,
                            items=items,
                        )
                        st.success(
                            f"Grupo guardado correctamente (ID {nuevo_id}) con {len(items)} cupones."
                        )
            except Exception as e:
                st.error(f"No se pudo procesar el Excel de cupones: {e}")

    st.markdown("### Grupos guardados")
    grupos = cargar_grupos_cupones_desde_bd()
    if grupos.empty:
        st.info("Aún no hay grupos de cupones guardados.")
    else:
        opciones_grupo: dict[str, int] = {}
        for _, g in grupos.iterrows():
            lbl = (
                f"ID {int(g['id'])} · {g['nombre_grupo']} · "
                f"{g['tipo_cupon']} · {int(g['total_items'])} items · {g['creado_en']}"
            )
            opciones_grupo[lbl] = int(g["id"])
        sel_lbl = st.selectbox(
            "Seleccionar grupo para consultar / generar PDF",
            options=list(opciones_grupo.keys()),
            key="sel_grupo_cupones",
        )
        grupo_id = opciones_grupo[sel_lbl]
        items_grupo = cargar_items_de_grupo(grupo_id)
        meta = grupos[grupos["id"] == grupo_id].iloc[0]

        c1, c2, c3 = st.columns(3)
        c1.metric("ID", int(meta["id"]))
        c2.metric("Tipo", str(meta["tipo_cupon"]))
        c3.metric("Elementos", int(meta["total_items"]))
        st.caption(
            f"Código de barras del grupo: **{meta['codigo_barras']}** · "
            f"Creado: {meta['creado_en']}"
        )

        if items_grupo:
            df_prev_cupon = pd.DataFrame(items_grupo)
            df_prev_cupon.insert(0, "#", range(1, len(df_prev_cupon) + 1))
            st.dataframe(df_prev_cupon, use_container_width=True, hide_index=True)

        if st.button("Generar PDF de este grupo", key="btn_pdf_cupones"):
            try:
                pdf_cupones = generar_pdf_cupones_internet(
                    items_grupo,
                    codigo_barras_grupo=str(meta["codigo_barras"]),
                )
                st.session_state["pdf_cupones"] = pdf_cupones
                st.session_state["pdf_cupones_nombre"] = f"cupones_grupo_{grupo_id}.pdf"
                st.success(
                    f"PDF listo con {len(items_grupo)} fichas en formato "
                    f"{COLS_CUPON}×{ROWS_CUPON} por hoja."
                )
            except Exception as e:
                st.error(f"No se pudo generar el PDF de cupones: {e}")

    pdf_cupones_guardado = st.session_state.get("pdf_cupones")
    if pdf_cupones_guardado:
        st.download_button(
            label="Descargar cupones_internet.pdf",
            data=pdf_cupones_guardado,
            file_name=st.session_state.get("pdf_cupones_nombre", "cupones_internet.pdf"),
            mime="application/pdf",
            key="dl_cupones",
        )
