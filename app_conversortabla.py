import streamlit as st
import pandas as pd
import json
import re
from html import escape as html_escape
import streamlit.components.v1 as components

st.set_page_config(page_title="Generador de Tabla HTML igual a Canvas LMS", page_icon="🧱", layout="centered")
st.title("Generador de Tabla HTML igual a Canvas LMS")

# =========================
# Estado / Sesión
# =========================
if "texto_html" not in st.session_state:
    st.session_state["texto_html"] = ""
if "loaded_template" not in st.session_state:
    st.session_state["loaded_template"] = False

def _append_snippet(snippet: str):
    st.session_state["texto_html"] = (st.session_state.get("texto_html", "") or "") + snippet

# =========================
# Constantes / Utilidades
# =========================
LOGICAL_FIELDS = [
    "Unidad Didáctica",
    "Tema del encuentro",
    "Duración",
    "Fecha de realización",
    "Enlace de Conexión",
    "Enlace de Grabación",
]

DEFAULT_EXPECTED = {
    "Unidad Didáctica": "Unidad Didáctica",
    "Tema del encuentro": "Tema del encuentro",
    "Duración": "Duración ",
    "Fecha de realización": "Fecha  de realización",
    "Enlace de Conexión": "Enlace de Conexión",
    "Enlace de Grabación": "Enlace de Grabación",
}

DEFAULT_HEADERS_LABELS = {
    "Unidad Didáctica": "Unidad Didáctica",
    "Tema del encuentro": "Tema del encuentro",
    "Duración": "Duración",
    "Fecha de realización": "Fecha de realización",
    "Enlace de Conexión": "Enlace de Conexión",
    "Enlace de Grabación": "Enlace de Grabación",
}

def labelize(x) -> str:
    try:
        s = str(x)
    except Exception:
        s = repr(x)
    return s

def safe_index(seq, value):
    try:
        return seq.index(value)
    except ValueError:
        return 0

def best_default(colnames, target):
    """Mejor coincidencia de nombre de columna. Soporta cabeceras no-string."""
    def norm_one(x):
        s = str(x)
        return " ".join(s.strip().lower().split())

    aliases = {"duración ": "duración", "fecha  de realización": "fecha de realización"}
    target_norm = norm_one(aliases.get(target, target))

    for orig in colnames:
        if norm_one(orig) == target_norm:
            return orig

    hints = {
        "Unidad Didáctica": ["unidad"],
        "Tema del encuentro": ["tema"],
        "Duración": ["duración", "duracion"],
        "Fecha de realización": ["fecha"],
        "Enlace de Conexión": ["conexión", "conexion", "enlace"],
        "Enlace de Grabación": ["grabación", "grabacion"],
    }
    for orig in colnames:
        n = norm_one(orig)
        for h in hints.get(target, []):
            if h in n:
                return orig
    return None

def ensure_unique_order(order_list):
    if sorted(order_list) != sorted(LOGICAL_FIELDS):
        return LOGICAL_FIELDS[:]
    if len(set(order_list)) != len(order_list):
        return LOGICAL_FIELDS[:]
    return order_list

# ====== Estilos (color + fuente + compacto + alineación + bordes) ======
def make_style(primary="#ba372a", compact=False, font_family="Arial, Helvetica, sans-serif",
               tema_left=False, show_th_borders=True, show_td_borders=True):
    th_border = f"1px solid {primary}" if show_th_borders else "0"
    td_border = f"1px solid {primary}" if show_td_borders else "0"
    return {
        "primary": primary,
        "font": font_family,
        "title_size": "14pt" if compact else "16pt",
        "header_font": "11pt" if compact else "12pt",
        "cell_font": "10pt" if compact else "11pt",
        "padding": "4px" if compact else "8px",
        "tema_main": "11pt" if compact else "11.5pt",
        "tema_sub": "9.5pt" if compact else "10pt",
        "tema_align": "left" if tema_left else "center",
        "cell_align_default": "center",
        "th_border": th_border,
        "td_border": td_border,
    }

# ====== Protección de tabla: sanear y encapsular bloque de usuario ======
SAFE_TAG_WHITELIST = {"strong","em","b","i","u","a","p","br","hr","ul","ol","li","h1","h2","h3","h4","h5","h6"}

def sanitize_user_html(raw: str) -> str:
    """Escapa etiquetas peligrosas (table, tr, td, style, script, etc.) y permite solo etiquetas de texto básicas."""
    if not raw:
        return ""
    raw = re.sub(r"(?is)<\s*(script|style)\b.*?</\s*\1\s*>", "", raw)

    def _repl_open(m):
        tag_full = m.group(0)
        tag_name = (m.group(1) or "").split()[0].lower().strip("/ ")
        return tag_full if tag_name in SAFE_TAG_WHITELIST else tag_full.replace("<", "&lt;").replace(">", "&gt;")

    def _repl_close(m):
        tag_full = m.group(0)
        tag_name = (m.group(1) or "").lower().strip()
        return tag_full if tag_name in SAFE_TAG_WHITELIST else tag_full.replace("<", "&lt;").replace(">", "&gt;")

    raw = re.sub(r"<\s*([a-zA-Z0-9]+)([^>]*)>", _repl_open, raw)
    raw = re.sub(r"</\s*([a-zA-Z0-9]+)\s*>", _repl_close, raw)
    return raw

def wrap_user_block(html_text: str, font_family: str) -> str:
    """Aísla el bloque del usuario en un contenedor propio para no afectar la tabla."""
    if not html_text:
        return ""
    return (
        f"<div style=\\"margin: 0 0 12px 0; font-family:{font_family}; color:#404e5c; font-size:11pt;\\">"
        f"{html_text}"
        f"</div>"
    )

def render_cell_html(lfield, row, style):
    import pandas as _pd
    if lfield == "Tema del encuentro":
        tema_val = "" if _pd.isna(row.get(lfield, "")) else str(row.get(lfield, ""))
        partes = tema_val.split("\\n", 1)
        if len(partes) == 2:
            tema_principal = partes[0].strip()
            tema_sub = partes[1].strip()
            return (
                f'<span style="font-weight: bold; color: #404e5c; font-size: {style["tema_main"]};">{tema_principal}</span><br>'
                f'<span style="color: #7b858d; font-size: {style["tema_sub"]};">{tema_sub}</span>'
            )
        else:
            return f'<span style="font-weight: bold; color: #404e5c; font-size: {style["tema_main"]};">{tema_val}</span>'

    if lfield == "Enlace de Conexión":
        val = str(row.get(lfield, "") or "").strip()
        if val:
            icono_flecha = '&#8594;'
            return (
                f'<a href="{val}" target="_blank" '
                f'style="color: {style["primary"]}; font-weight: bold; text-decoration: underline; '
                f'font-family: {style["font"]}; font-size: {style["cell_font"]};">'
                f'ENLACE <span style="font-size:10pt;">{icono_flecha}</span></a>'
            )
        return ""

    if lfield == "Enlace de Grabación":
        return (f'<span style="color: #404e5c; font-weight: bold; '
                f'font-family: {style["font"]}; font-size: {style["cell_font"]};">GRABACIÓN</span>')

    return str(row.get(lfield, "") if pd.notna(row.get(lfield, "")) else "")

def generar_tabla_html(df_disp, titulo, headers_by_logical, order_list, hidden_set, texto_extra, style, proteger_tabla):
    """Devuelve (html_completo, html_solo_tabla)."""
    # --- bloque del usuario, saneado/encapsulado ---
    user_block = sanitize_user_html(texto_extra) if proteger_tabla else (texto_extra or "")
    user_block = wrap_user_block(user_block, style["font"]) if user_block else ""

    # --- título ---
    title_html = f"""
    <p style="text-align: left;">
        <span style="font-family: {style["font"]}; color: #000000; font-size: {style["title_size"]}; font-weight: bold;">
            {titulo}
        </span>
    </p>
    """

    # --- tabla (SOLO la tabla) ---
    outer_border = f"1px solid {style['primary']}"
    table_html = f'<table style="border-collapse: collapse; width: 100%; border: {outer_border};">\\n<tr>\\n'
    for lfield in order_list:
        if lfield in hidden_set:
            continue
        label = headers_by_logical.get(lfield, lfield)
        table_html += (
            f'<th style="background-color: {style["primary"]}; color: #fff; text-align: center; '
            f'font-family: {style["font"]}; font-size: {style["header_font"]}; font-weight: bold; '
            f'padding: {style["padding"]}; border: {style["th_border"]};">{label}</th>\\n'
        )
    table_html += "</tr>\\n"

    for _, row in df_disp.iterrows():
        table_html += "<tr>\\n"
        for lfield in order_list:
            if lfield in hidden_set:
                continue
            cell_html = render_cell_html(lfield, row, style)
            align = style["tema_align"] if lfield == "Tema del encuentro" else style["cell_align_default"]
            table_html += (
                f'<td style="text-align: {align}; font-family: {style["font"]}; '
                f'font-size: {style["cell_font"]}; color: #404e5c; border: {style["td_border"]}; '
                f'padding: {style["padding"]};">{cell_html}</td>\\n'
            )
        table_html += "</tr>\\n"
    table_html += "</table>"

    # --- html completo ---
    full_html = (user_block or "") + title_html + table_html
    return full_html, table_html

def build_config_dict(map_cols, header_labels_by_logical, display_order, hidden_columns, titulo_principal,
                      texto_html, primary, compact, font_family, tema_left, proteger_tabla,
                      show_th_borders, show_td_borders):
    return {
        "map_cols": map_cols,
        "header_labels": header_labels_by_logical,
        "display_order": display_order,
        "hidden_columns": list(hidden_columns),
        "titulo_principal": titulo_principal,
        "texto_html": texto_html,
        "primary_color": primary,
        "compact_mode": compact,
        "font_family": font_family,
        "tema_align_left": tema_left,
        "protect_table": proteger_tabla,
        "show_th_borders": show_th_borders,
        "show_td_borders": show_td_borders,
    }

def apply_loaded_template(conf):
    st.session_state["tpl_map_cols"] = conf.get("map_cols", {})
    st.session_state["tpl_header_labels"] = conf.get("header_labels", DEFAULT_HEADERS_LABELS)
    st.session_state["tpl_display_order"] = conf.get("display_order", LOGICAL_FIELDS)
    st.session_state["tpl_hidden_columns"] = set(conf.get("hidden_columns", []))
    st.session_state["tpl_titulo_principal"] = conf.get("titulo_principal", "Programación de encuentros sincrónicos")
    st.session_state["texto_html"] = conf.get("texto_html", "")
    st.session_state["tpl_primary_color"] = conf.get("primary_color", "#ba372a")
    st.session_state["tpl_compact_mode"] = bool(conf.get("compact_mode", False))
    st.session_state["tpl_font_family"] = conf.get("font_family", "Arial, Helvetica, sans-serif")
    st.session_state["tpl_tema_left"] = bool(conf.get("tema_align_left", False))
    st.session_state["tpl_protect_table"] = bool(conf.get("protect_table", True))
    st.session_state["tpl_show_th_borders"] = bool(conf.get("show_th_borders", True))
    st.session_state["tpl_show_td_borders"] = bool(conf.get("show_td_borders", True))
    st.session_state["loaded_template"] = True

# =========================
# Cargar / Guardar plantilla
# =========================
with st.expander("📦 Plantillas de configuración (opcional)"):
    st.caption("Puedes cargar una plantilla JSON previamente guardada, o exportar la configuración actual.")
    # clave única
    tpl_upl = st.file_uploader("Cargar plantilla (.json)", type=["json"], key="tpl_json_uploader_v3")
    if tpl_upl is not None:
        try:
            conf = json.load(tpl_upl)
            apply_loaded_template(conf)
            st.success("Plantilla cargada. Los controles se han ajustado con la configuración.")
        except Exception as e:
            st.error(f"No se pudo leer la plantilla: {e}")

# =========================
# Apariencia: color, fuente, compacto, alineación y bordes
# =========================
st.markdown("### Apariencia")
default_primary = st.session_state.get("tpl_primary_color", "#ba372a")
primary_color = st.color_picker("Color institucional (encabezados, bordes y enlaces)", value=default_primary)

font_options = {
    "Arial (segura)": "Arial, Helvetica, sans-serif",
    "Roboto (moderna)": "Roboto, Arial, sans-serif",
    "Georgia (serif)": "Georgia, serif",
    "Times New Roman (serif)": "'Times New Roman', Times, serif",
}
default_font_label = next((k for k, v in font_options.items()
                           if v == st.session_state.get("tpl_font_family", "Arial, Helvetica, sans-serif")), "Arial (segura)")
font_label = st.selectbox("Tipografía", options=list(font_options.keys()),
                          index=list(font_options.keys()).index(default_font_label))
font_family = font_options[font_label]

default_compact = bool(st.session_state.get("tpl_compact_mode", False))
compact_mode = st.checkbox("Modo compacto (tipografía y celdas más pequeñas)", value=default_compact)

tema_left_default = bool(st.session_state.get("tpl_tema_left", False))
tema_left = st.checkbox("Alinear a la izquierda solo la columna “Tema del encuentro”", value=tema_left_default)

protect_table_default = bool(st.session_state.get("tpl_protect_table", True))
proteger_tabla = st.checkbox("Proteger tabla (sanear HTML conflictivo del bloque superior)", value=protect_table_default)

# toggles bordes separados
show_th_borders_default = bool(st.session_state.get("tpl_show_th_borders", True))
show_td_borders_default = bool(st.session_state.get("tpl_show_td_borders", True))
col_b1, col_b2 = st.columns(2)
with col_b1:
    show_th_borders = st.checkbox("Líneas internas en cabecera (th)", value=show_th_borders_default)
with col_b2:
    show_td_borders = st.checkbox("Líneas internas en cuerpo (td)", value=show_td_borders_default)

style = make_style(
    primary=primary_color,
    compact=compact_mode,
    font_family=font_family,
    tema_left=tema_left,
    show_th_borders=show_th_borders,
    show_td_borders=show_td_borders
)

# =========================
# Carga de archivo
# =========================
excel_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"], key="excel_xlsx_uploader_v3")

if excel_file:
    df_raw = pd.read_excel(excel_file)
    colnames = list(df_raw.columns)

    labels = [labelize(c) for c in colnames]
    label_to_orig = {labelize(c): c for c in colnames}

    # =========================
    # Paso 1: Mapeo de columnas
    # =========================
    st.markdown("### Paso 1: Mapea las columnas de tu Excel a los campos lógicos")
    st.caption("Si tu Excel tiene encabezados distintos, asígnalos aquí. Si coinciden, quedarán preseleccionados.")

    tpl_map = st.session_state.get("tpl_map_cols", {})

    def default_label_for(key):
        cand = best_default(colnames, DEFAULT_EXPECTED[key])
        return labelize(cand) if cand is not None else None

    def prefill_label_for(key):
        tpl_val = tpl_map.get(key)
        if tpl_val in colnames:
            return labelize(tpl_val)
        return default_label_for(key)

    options_labels = ["(ninguna)"] + labels

    colA, colB = st.columns(2)
    with colA:
        ud_pref = prefill_label_for("Unidad Didáctica")
        map_ud_lbl = st.selectbox("Columna para: Unidad Didáctica", options=options_labels,
                                  index=safe_index(options_labels, ud_pref) if ud_pref else 0, key="map_ud_lbl")

        tema_pref = prefill_label_for("Tema del encuentro")
        map_tema_lbl = st.selectbox("Columna para: Tema del encuentro", options=options_labels,
                                    index=safe_index(options_labels, tema_pref) if tema_pref else 0, key="map_tema_lbl")

        dur_pref = prefill_label_for("Duración")
        map_dur_lbl = st.selectbox("Columna para: Duración", options=options_labels,
                                   index=safe_index(options_labels, dur_pref) if dur_pref else 0, key="map_dur_lbl")

    with colB:
        fecha_pref = prefill_label_for("Fecha de realización")
        map_fecha_lbl = st.selectbox("Columna para: Fecha de realización", options=options_labels,
                                     index=safe_index(options_labels, fecha_pref) if fecha_pref else 0, key="map_fecha_lbl")

        enlace_pref = prefill_label_for("Enlace de Conexión")
        map_enlace_lbl = st.selectbox("Columna para: Enlace de Conexión (opcional, puedes sobrescribir abajo)", options=options_labels,
                                      index=safe_index(options_labels, enlace_pref) if enlace_pref else 0, key="map_enlace_lbl")

        grab_pref = prefill_label_for("Enlace de Grabación")
        map_grab_lbl = st.selectbox("Columna para: Enlace de Grabación (opcional)", options=options_labels,
                                    index=safe_index(options_labels, grab_pref) if grab_pref else 0, key="map_grab_lbl")

    map_cols = {
        "Unidad Didáctica": label_to_orig.get(map_ud_lbl) if map_ud_lbl != "(ninguna)" else "(ninguna)",
        "Tema del encuentro": label_to_orig.get(map_tema_lbl) if map_tema_lbl != "(ninguna)" else "(ninguna)",
        "Duración": label_to_orig.get(map_dur_lbl) if map_dur_lbl != "(ninguna)" else "(ninguna)",
        "Fecha de realización": label_to_orig.get(map_fecha_lbl) if map_fecha_lbl != "(ninguna)" else "(ninguna)",
        "Enlace de Conexión": label_to_orig.get(map_enlace_lbl) if map_enlace_lbl != "(ninguna)" else "(ninguna)",
        "Enlace de Grabación": label_to_orig.get(map_grab_lbl) if map_grab_lbl != "(ninguna)" else "(ninguna)",
    }

    required_min = ["Unidad Didáctica", "Tema del encuentro", "Duración", "Fecha de realización"]
    missing_map = [lf for lf in required_min if map_cols.get(lf) in (None, "(ninguna)")]
    if missing_map:
        st.error(f"Faltan asignaciones para: {', '.join(missing_map)}. Asigna esas columnas para continuar.")
        st.stop()

    rename_dict = {src: lf for lf, src in map_cols.items() if src and src != "(ninguna)"}
    df = df_raw.rename(columns=rename_dict)
    keep_cols = [c for c in LOGICAL_FIELDS if c in df.columns]
    df = df[keep_cols].copy()

    # =========================
    # Paso 2: Filtro por fecha y enlace global
    # =========================
    st.markdown("### Paso 2: Filtra por fecha y (opcional) sobreescribe enlace de conexión")
    fecha_col = "Fecha de realización"
    if fecha_col not in df.columns:
        st.error("No se encontró la columna lógica 'Fecha de realización' luego del mapeo.")
        st.stop()

    try:
        df[fecha_col] = pd.to_datetime(df[fecha_col]).dt.date
    except Exception:
        pass

    fechas_disponibles = df[fecha_col].dropna().unique().tolist()
    fechas_disponibles = sorted(fechas_disponibles, key=lambda x: str(x))
    fechas_seleccionadas = st.multiselect(
        "Selecciona las fechas de realización para mostrar en la tabla:",
        options=fechas_disponibles,
        default=fechas_disponibles
    )
    df_filtrado = df[df[fecha_col].isin(fechas_seleccionadas)].copy()

    enlace_conexion_global = st.text_input("Enlace de conexión global (opcional; sobrescribe la columna):", value="")
    if enlace_conexion_global.strip():
        df_filtrado["Enlace de Conexión"] = enlace_conexion_global.strip()
    else:
        if "Enlace de Conexión" not in df_filtrado.columns:
            df_filtrado["Enlace de Conexión"] = ""
    if "Enlace de Grabación" not in df_filtrado.columns:
        df_filtrado["Enlace de Grabación"] = ""

    st.write("Vista previa de la tabla filtrada:")
    st.dataframe(df_filtrado, use_container_width=True)

    # =========================
    # Paso 3: Encabezados visibles
    # =========================
    st.markdown("### Paso 3: Personaliza los encabezados visibles (opcional)")
    default_title = st.session_state.get("tpl_titulo_principal", "Programación de encuentros sincrónicos")
    titulo_principal = st.text_input("Título principal encima de la tabla", value=default_title, key="titulo_principal")

    tpl_hdrs = st.session_state.get("tpl_header_labels", DEFAULT_HEADERS_LABELS)
    col1, col2 = st.columns(2)
    with col1:
        h_ud = st.text_input("Encabezado: Unidad Didáctica", value=tpl_hdrs.get("Unidad Didáctica", "Unidad Didáctica"))
        h_tema = st.text_input("Encabezado: Tema del encuentro", value=tpl_hdrs.get("Tema del encuentro", "Tema del encuentro"))
        h_dur = st.text_input("Encabezado: Duración", value=tpl_hdrs.get("Duración", "Duración"))
    with col2:
        h_fecha = st.text_input("Encabezado: Fecha de realización", value=tpl_hdrs.get("Fecha de realización", "Fecha de realización"))
        h_enlace = st.text_input("Encabezado: Enlace de Conexión", value=tpl_hdrs.get("Enlace de Conexión", "Enlace de Conexión"))
        h_grab = st.text_input("Encabezado: Enlace de Grabación", value=tpl_hdrs.get("Enlace de Grabación", "Enlace de Grabación"))

    header_labels_by_logical = {
        "Unidad Didáctica": h_ud,
        "Tema del encuentro": h_tema,
        "Duración": h_dur,
        "Fecha de realización": h_fecha,
        "Enlace de Conexión": h_enlace,
        "Enlace de Grabación": h_grab,
    }

    # =========================
    # Paso 4: Orden y visibilidad
    # =========================
    st.markdown("### Paso 4: Orden de columnas y visibilidad")
    st.caption("Elige el orden 1→6 y qué columnas ocultar. El ocultamiento no modifica tu Excel, solo la salida HTML.")

    tpl_order = st.session_state.get("tpl_display_order", LOGICAL_FIELDS)
    def ord_idx_for(field_name, fallback_index):
        try:
            return tpl_order.index(field_name)
        except ValueError:
            return fallback_index

    ocol1, ocol2, ocol3 = st.columns(3)
    with ocol1:
        ord1 = st.selectbox("Posición 1", LOGICAL_FIELDS, index=ord_idx_for("Unidad Didáctica", 0), key="ord1")
        ord2 = st.selectbox("Posición 2", LOGICAL_FIELDS, index=ord_idx_for("Tema del encuentro", 1), key="ord2")
    with ocol2:
        ord3 = st.selectbox("Posición 3", LOGICAL_FIELDS, index=ord_idx_for("Duración", 2), key="ord3")
        ord4 = st.selectbox("Posición 4", LOGICAL_FIELDS, index=ord_idx_for("Fecha de realización", 3), key="ord4")
    with ocol3:
        ord5 = st.selectbox("Posición 5", LOGICAL_FIELDS, index=ord_idx_for("Enlace de Conexión", 4), key="ord5")
        ord6 = st.selectbox("Posición 6", LOGICAL_FIELDS, index=ord_idx_for("Enlace de Grabación", 5), key="ord6")

    display_order = ensure_unique_order([ord1, ord2, ord3, ord4, ord5, ord6])
    if display_order != [ord1, ord2, ord3, ord4, ord5, ord6]:
        st.warning("Había duplicados o faltantes en el orden. Se restauró el orden lógico por defecto.")
    st.write("**Orden actual:** ", " → ".join(display_order))

    st.markdown("#### Ocultar columnas (opcional)")
    tpl_hidden = st.session_state.get("tpl_hidden_columns", set())
    hidden_cols = set()
    c1, c2, c3 = st.columns(3)
    with c1:
        hide_ud = st.checkbox("Ocultar: Unidad Didáctica", value=("Unidad Didáctica" in tpl_hidden))
        hide_tema = st.checkbox("Ocultar: Tema del encuentro", value=("Tema del encuentro" in tpl_hidden))
    with c2:
        hide_dur = st.checkbox("Ocultar: Duración", value=("Duración" in tpl_hidden))
        hide_fecha = st.checkbox("Ocultar: Fecha de realización", value=("Fecha de realización" in tpl_hidden))
    with c3:
        hide_enlace = st.checkbox("Ocultar: Enlace de Conexión", value=("Enlace de Conexión" in tpl_hidden))
        hide_grab = st.checkbox("Ocultar: Enlace de Grabación", value=("Enlace de Grabación" in tpl_hidden))

    if hide_ud: hidden_cols.add("Unidad Didáctica")
    if hide_tema: hidden_cols.add("Tema del encuentro")
    if hide_dur: hidden_cols.add("Duración")
    if hide_fecha: hidden_cols.add("Fecha de realización")
    if hide_enlace: hidden_cols.add("Enlace de Conexión")
    if hide_grab: hidden_cols.add("Enlace de Grabación")

    # =========================
    # Paso 5: Editor de texto con barra (opcional)
    # =========================
    st.markdown("### Paso 5: Texto (opcional) para insertar arriba de la tabla")
    st.caption("Este texto se insertará en el HTML final. Puedes escribir texto plano o HTML simple.")
    with st.expander("Barra de herramientas de texto (opcional)"):
        tcol1, tcol2, tcol3, tcol4, tcol5, tcol6 = st.columns(6)
        with tcol1:
            if st.button("H1"):
                _append_snippet("<h1 style='font-family: Arial, Helvetica, sans-serif;'>Título H1</h1>\n")
        with tcol2:
            if st.button("H2"):
                _append_snippet("<h2 style='font-family: Arial, Helvetica, sans-serif;'>Título H2</h2>\n")
        with tcol3:
            if st.button("H3"):
                _append_snippet("<h3 style='font-family: Arial, Helvetica, sans-serif;'>Título H3</h3>\n")
        with tcol4:
            if st.button("Negrita"):
                _append_snippet("<strong>texto en negrita</strong> ")
        with tcol5:
            if st.button("Cursiva"):
                _append_snippet("<em>texto en cursiva</em> ")
        with tcol6:
            if st.button("Lista"):
                _append_snippet("<ul><li>Elemento 1</li><li>Elemento 2</li></ul>\n")

        tcol7, tcol8, tcol9 = st.columns(3)
        with tcol7:
            if st.button("Párrafo"):
                _append_snippet("<p style='font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color:#404e5c;'>Tu párrafo aquí.</p>\n")
        with tcol8:
            if st.button("Enlace"):
                _append_snippet("<a href='https://ejemplo.com' target='_blank'>Un enlace</a> ")
        with tcol9:
            if st.button("Separador"):
                _append_snippet("<hr/>\n")

    st.text_area("Editor (puedes escribir o pegar HTML; la barra de herramientas inserta fragmentos):",
                 key="texto_html", height=180)

    # =========================
    # Generación del HTML
    # =========================
    headers_by_logical = header_labels_by_logical
    full_html, table_only_html = generar_tabla_html(
        df_disp=df_filtrado,
        titulo=titulo_principal,
        headers_by_logical=headers_by_logical,
        order_list=display_order,
        hidden_set=hidden_cols,
        texto_extra=st.session_state.get("texto_html", ""),
        style=style,
        proteger_tabla=proteger_tabla
    )

    st.markdown("### Copia este HTML para Canvas (COMPLETO):")
    st.code(full_html, language="html")

    st.markdown("### Copia solo la TABLA (sin bloque superior ni título):")
    st.code(table_only_html, language="html")

    # ====== Botones de copiado ======
    components.html(
        f"""
        <div style="margin: 8px 0 6px 0;">
            <button id="copyFull" style="background:#0f62fe; color:white; border:none; padding:8px 12px; border-radius:6px; cursor:pointer;">
                Copiar HTML completo
            </button>
            <button id="copyTable" style="background:#12a150; color:white; border:none; padding:8px 12px; border-radius:6px; cursor:pointer; margin-left:8px;">
                Copiar solo la tabla
            </button>
            <span id="copyStatus" style="margin-left:8px; color:#555;"></span>
        </div>
        <textarea id="htmlFull" style="position:absolute; left:-10000px; top:-10000px;">{html_escape(full_html)}</textarea>
        <textarea id="htmlTable" style="position:absolute; left:-12000px; top:-12000px;">{html_escape(table_only_html)}</textarea>
        <script>
        const status = document.getElementById('copyStatus');
        async function copyFrom(id) {{
            const area = document.getElementById(id);
            try {{
                await navigator.clipboard.writeText(area.value);
                status.textContent = '✔ Copiado';
                setTimeout(()=>{{status.textContent='';}}, 2000);
            }} catch(e) {{
                area.select();
                document.execCommand('copy');
                status.textContent = '✔ Copiado (fallback)';
                setTimeout(()=>{{status.textContent='';}}, 2000);
            }}
        }}
        document.getElementById('copyFull').addEventListener('click', () => copyFrom('htmlFull'));
        document.getElementById('copyTable').addEventListener('click', () => copyFrom('htmlTable'));
        </script>
        """,
        height=70,
    )

    st.markdown("### Vista previa (¡así se vería en Canvas!):", unsafe_allow_html=True)
    st.markdown(full_html, unsafe_allow_html=True)

    # =========================
    # Guardar / Descargar
    # =========================
    st.markdown("### Guardar esta configuración como plantilla")
    current_conf = build_config_dict(
        map_cols=map_cols,
        header_labels_by_logical=headers_by_logical,
        display_order=display_order,
        hidden_columns=hidden_cols,
        titulo_principal=titulo_principal,
        texto_html=st.session_state.get("texto_html", ""),
        primary=primary_color,
        compact=compact_mode,
        font_family=font_family,
        tema_left=tema_left,
        proteger_tabla=proteger_tabla,
        show_th_borders=show_th_borders,
        show_td_borders=show_td_borders
    )
    st.download_button(
        "⬇️ Descargar plantilla (.json)",
        data=json.dumps(current_conf, ensure_ascii=False, indent=2),
        file_name="plantilla_tabla_canvas.json",
        mime="application/json"
    )

    # HTML mínimo que envuelve solo la tabla (para abrir directo en navegador)
    table_only_page = f"<!doctype html><html><head><meta charset='utf-8'><title>Tabla Canvas</title></head><body>{table_only_html}</body></html>"

    # Descargas
    col_dl1, col_dl2, col_dl3 = st.columns(3)
    with col_dl1:
        st.download_button(
            "⬇️ Descargar HTML completo (.html)",
            data=full_html,
            file_name="tabla_canvas_completo.html",
            mime="text/html"
        )
    with col_dl2:
        st.download_button(
            "⬇️ Descargar solo la tabla (.txt)",
            data=table_only_html,
            file_name="tabla_canvas_solo_tabla.txt",
            mime="text/plain"
        )
    with col_dl3:
        st.download_button(
            "⬇️ Descargar solo la tabla (.html)",
            data=table_only_page,
            file_name="tabla_canvas_solo_tabla.html",
            mime="text/html"
        )

else:
    st.info("Sube un archivo Excel para comenzar.")