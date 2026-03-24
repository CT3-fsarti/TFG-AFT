import streamlit as st
import pandas as pd

st.set_page_config(page_title="Debug Streamlit", layout="centered")

st.title("🛠️ Panel de Diagnóstico de Streamlit")
st.write("Este script aísla el problema del formato de las tablas para saber qué está fallando en la nube.")

st.markdown("---")

# 1. Comprobación de versiones
st.header("1️⃣ Versiones del Entorno")
st.info(f"**Versión de Streamlit instalada:** {st.__version__}")
st.info(f"**Versión de Pandas instalada:** {pd.__version__}")

if st.__version__ < "1.32.0":
    st.warning("⚠️ Tienes una versión de Streamlit anterior a la 1.32.0. El centrado nativo no funcionará.")

# Datos de prueba muy simples
df_prueba = pd.DataFrame({
    "ID Nodo": [101, 102, 103],
    "Tipo": ["Financiador", "Célula", "Mula"],
    "Valor": [5000, 0, 150]
})

st.markdown("---")

# 2. Prueba con Pandas Styler (El método clásico CSS)
st.header("2️⃣ Prueba: Centrado con Pandas Styler")
st.write("Si el motor moderno de Streamlit (Glide Data Grid) es muy estricto, ignorará esta orden y verás los números a la derecha y los textos a la izquierda.")

styler = df_prueba.style.set_properties(**{'text-align': 'center'})
styler = styler.set_table_styles([
    dict(selector='th', props=[('text-align', 'center !important')]),
    dict(selector='td', props=[('text-align', 'center !important')])
])

st.dataframe(styler, use_container_width=True)

st.markdown("---")

# 3. Prueba con Streamlit Nativo (El método moderno)
st.header("3️⃣ Prueba: Centrado Nativo (st.column_config)")
st.write("Intenta forzar el centrado usando la configuración oficial de las últimas versiones de Streamlit.")

try:
    # Intentamos aplicar el alignment="center" que nos dio error antes
    config_centrada = {
        "ID Nodo": st.column_config.Column(alignment="center"),
        "Tipo": st.column_config.Column(alignment="center"),
        "Valor": st.column_config.Column(alignment="center")
    }
    st.dataframe(df_prueba, column_config=config_centrada, use_container_width=True)
    st.success("✅ ¡ÉXITO! La orden de centrado nativo funciona correctamente en esta versión.")
    
except TypeError as e:
    st.error(f"❌ FALLO CONFIRMADO: Tu versión de Streamlit no soporta el parámetro 'alignment'.")
    st.code(f"Error exacto: {e}")
except Exception as e:
    st.error(f"❌ Error desconocido.")
    st.code(f"Error exacto: {e}")

st.markdown("---")
st.write("💡 **Siguiente paso:** Revisa qué versión de Streamlit aparece en el punto 1.")