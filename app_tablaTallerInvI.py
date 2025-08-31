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

# ... resto del código (omitido aquí por brevedad) ...

def wrap_user_block(html_text: str, font_family: str) -> str:
    """Aísla el bloque del usuario en un contenedor propio para no afectar la tabla."""
    if not html_text:
        return ""
    return (
        f"<div style='margin: 0 0 12px 0; font-family:{font_family}; color:#404e5c; font-size:11pt;'>"
        f"{html_text}"
        f"</div>"
    )

# ... resto del código ...