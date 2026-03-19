import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import networkx as nx
from pyvis.network import Network
from openpyxl import load_workbook
import os
import base64

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA (BRANDING)
# ==========================================
st.set_page_config(page_title="Marina Sarti Pineda - Inteligencia FT/AML", layout="wide")

st.markdown("""
<style>
    :root {
        --uc3m-blue: #005691;
    }
    .big-font { font-size:16px !important; color: #31333F; }
    .header-style { background-color: #F8F9FB; padding: 25px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #E6E9EF; }
    .centered-text { text-align: center; }
    .language-block { display: flex; align-items: center; margin-bottom: 15px; }
    .flag-img { width: 35px; height: 25px; object-fit: contain; margin-right: 15px; }
    .title-text { font-size: 18px; color: #111; line-height: 1.4; }
    .profile-photo {
        width: 160px; height: 160px; border-radius: 50% !important; 
        object-fit: cover !important; display: block;
        margin-left: auto; margin-right: auto; border: 3px solid var(--uc3m-blue); 
    }
</style>
""", unsafe_allow_html=True)

def get_base64_of_bin_file(bin_file):
    try:
        with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()
    except FileNotFoundError: return None

assets_status = {
    "Combined Logo B64": get_base64_of_bin_file("Logo_uc3m_PSL.png"),
    "Flag ES B64": get_base64_of_bin_file("logo_spa.jpg"),
    "Flag GB B64": get_base64_of_bin_file("logo_GB.png"),
    "Flag FR B64": get_base64_of_bin_file("logo_FR.jpg"),
    "Marina Photo B64": get_base64_of_bin_file("marina_circular.png")
}

# ==========================================
# 2. ESTRUCTURA DE DISEÑO PRINCIPAL (COLUMNAS)
# ==========================================
col_main, col_profile = st.columns([3, 1])

with col_profile:
    st.markdown("""<div style="margin-top: 30px;"></div>""", unsafe_allow_html=True)
    with st.container():
        st.markdown("### 👤 Perfil de la Autora")
        if assets_status["Marina Photo B64"]:
             st.markdown(f"""<div style="margin-bottom: 15px;"><img src='data:image/png;base64,{assets_status["Marina Photo B64"]}' class='profile-photo' /></div>""", unsafe_allow_html=True)
        st.markdown(f"<p class='centered-text' style='font-size: 18px;'><strong>Marina Sarti Pineda</strong></p>", unsafe_allow_html=True)
        st.markdown("""Graduada en Economía por la Universidad Carlos III de Madrid y París Dauphine, Marina combina el rigor analítico macroeconómico con una vocación por la seguridad global y la inteligencia financiera. Su TFG pionero en Modelado Estructurado de Financiación del Terrorismo se desarrolló en el marco de las tipologías del GAFI.""")
        st.write("") 
        _, clcol, _ = st.columns([1, 2, 1])
        with clcol: st.markdown('[🔗 Perfil de LinkedIn](https://www.linkedin.com/in/marina-sarti-pineda-27211b29b/?originalSubdomain=es)', unsafe_allow_html=True)

with col_main:
    col_logo, col_titles = st.columns([1, 2.5])
    with col_logo:
        if assets_status["Combined Logo B64"]:
             st.markdown(f"""<img src='data:image/png;base64,{assets_status["Combined Logo B64"]}' style='width: 100%; height: auto;' />""", unsafe_allow_html=True)
    with col_titles:
        st.markdown("""
        <div style="text-align: center; padding-top: 15px;">
            <h1 style="color: #005691; margin-bottom: 5px; font-weight: bold; font-size: 32px;">Doble Titulación Internacional en Economía</h1>
            <h4 style="color: #9AA6B8; font-style: italic; font-weight: normal; font-size: 20px;">Marina Sarti Pineda</h4>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("---")
    
    # Banderas
    if assets_status["Flag ES B64"]: st.markdown(f"""<div class="language-block"><img src="data:image/jpeg;base64,{assets_status['Flag ES B64']}" class="flag-img" /><div class="title-text"><strong>Redes de Financiación del Terrorismo:<br>Simulación y análisis mediante Economía de Redes y Teoría de Juegos</strong></div></div>""", unsafe_allow_html=True)
    if assets_status["Flag FR B64"]: st.markdown(f"""<div class="language-block"><img src="data:image/jpeg;base64,{assets_status['Flag FR B64']}" class="flag-img" /><div class="title-text">Réseaux de Financement du Terrorisme:<br>Simulation et analyse à travers l'Économie des Réseaux et la Théorie des Jeux</div></div>""", unsafe_allow_html=True)
    if assets_status["Flag GB B64"]: st.markdown(f"""<div class="language-block"><img src="data:image/png;base64,{assets_status['Flag GB B64']}" class="flag-img" /><div class="title-text">Terrorism Financing Networks:<br>Simulation and Analysis through Network Economics and Game Theory</div></div>""", unsafe_allow_html=True)
    st.markdown("---")

    # Abstract
    st.markdown("## 🎓 Sinopsis del Proyecto")
    with st.expander("🇪🇸 🇬🇧 🇫🇷 Leer Sinopsis / Abstract Completo", expanded=True):
        tab_es, tab_en, tab_fr = st.tabs(["🇪🇸 Español", "🇬🇧 English", "🇫🇷 Français"])
        with tab_es: st.markdown("""La Financiación del Terrorismo (FT) consiste en la captación de fondos, lo cual abarca el proceso de solicitud, recaudación, provisión y puesta a disposición de dinero o activos con el fin de facilitar o potenciar la capacidad de cualquier persona u organización para llevar a cabo actividades relacionadas con el terrorismo. En el caso concreto de España, la Ley 10/2010 establece un marco riguroso frente al suministro, depósito o distribución de fondos.\n\nTanto los grandes grupos organizados como las pequeñas células e incluso los actores individuales necesitan dinero para el desarrollo de la actividad terrorista. La literatura académica y los informes institucionales coinciden en que la falta de fondos limita drásticamente su capacidad operativa, siendo la FT un elemento vertebrador del terrorismo global.\n\nEste trabajo fundamenta su análisis en información reciente, utilizando ejemplos observados en los últimos años que resultan representativos de las dinámicas contemporáneas. El objetivo principal consiste en construir una **simulación de red de financiación del terrorismo** basada en la evidencia recogida en la literatura especializada (tipologías del GAFI, EBA, etc.). \n\nSobre dicha simulación se realiza un análisis estructural mediante herramientas propias del análisis de redes (Economía de Redes) y la Teoría de Juegos, incluyendo el estudio de métricas de centralidad, la importancia relativa de los nodos clave (chokepoints) y la resiliencia del sistema ante intervenciones policiales. El modelo resultante constituye una representation realista y fundamentada empíricamente, derivando en un modelo analítico altamente útil para la inteligencia financiera y el diseño de políticas de seguridad.""")
        with tab_en: st.markdown("""Terrorist Financing (TF) involves the raising of funds, which encompasses the process of soliciting, collecting, providing, and making available money or assets to facilitate or enhance the capacity of any individual or organization to carry out terrorist activities. In Spain, Law 10/2010 establishes a rigorous framework against the supply, deposit, or distribution of funds.\n\nLarge organized groups, small cells, and lone actors require money to carry out terrorist activities. Academic literature and institutional reports agree that a lack of funds drastically limits their operational capacity, making TF a structural backbone of global terrorism.\n\nThis paper bases its analysis on recent information, using examples observed in recent years that are representative of contemporary dynamics. The main objective is to build a **simulation of a terrorist financing network** based on evidence gathered from specialized literature (FATF typologies, EBA, etc.).\n\nA structural analysis is performed on this simulation using Network Economics and Game Theory, including the study of centrality metrics, the relative importance of key nodes (chokepoints), and the system's resilience to law enforcement interventions. The resulting model constitutes a realistic and empirically grounded representation, resulting in an analytical model highly useful for financial intelligence and security policy design.""")
        with tab_fr: st.markdown("""Le Financement du Terrorisme (FT) consiste en la collecte de fonds, ce qui englobe le processus de sollicitation, de rassemblement, de fourniture et de mise à disposition d'argent ou d'actifs dans le but de faciliter la capacité à mener des activités terroristes. En Espagne, la loi 10/2010 établit un cadre rigoureux contre la fourniture ou la distribution de fonds.\n\nLes grands groupes organisés, les petites cellules et les acteurs individuels ont besoin d'argent pour développer leurs activités terroristes. La littérature académique s'accorde à dire que le manque de fonds limite considérablement leur capacité opérationnelle, le FT siendo un élément structurant du terrorisme mondial.\n\nCe travail fonde son analyse sur des informations récentes, en utilisant des exemples représentatifs des dynamiques contemporaines. L'objectif principal est de construire une **simulation d'un réseau de financement du terrorisme** basée sur les preuves recueillies dans la littérature spécialisée (typologies du GAFI, ABE, etc.).\n\nUne analyse structurelle est réalisée sur cette simulation à l'aide de l'Économie des Réseaux et de la Théorie des Jeux, incluant l'étude des métriques de centralité, l'importance des nœuds clés (points d'étranglement) et la résilience du système face aux interventions policières. Le modèle qui en résulte constitue une représentation réaliste et empiriquement fondée, aboutissant à un outil analytique très utile pour le renseignement financier.""")

# ==========================================
# 3. SIMULADOR Y TABLAS
# ==========================================
st.markdown("---")
st.markdown("## 🔬 Simulador de Inteligencia Operativa")
st.markdown("Herramienta interactiva para la simulación de rutas críticas. Sube el modelo base y evalúa la resiliencia de la red frente a la presión policial.")

archivo_subido = st.file_uploader("Sube tu archivo Excel del modelo", type=["xlsx"])

def leer_tabla_excel(wb, nombre_tabla_buscada):
    for hoja in wb.worksheets:
        if nombre_tabla_buscada in hoja.tables:
            rango = hoja.tables[nombre_tabla_buscada].ref
            datos_celdas = hoja[rango]
            filas = [[celda.value for celda in fila] for fila in datos_celdas]
            return pd.DataFrame(filas[1:], columns=filas[0])
    raise ValueError(f"No encuentro la tabla '{nombre_tabla_buscada}'.")

def aplicar_estilos(df):
    styler = df.style.set_properties(**{'text-align': 'center'})
    styler = styler.set_table_styles([dict(selector='th', props=[('text-align', 'center')])])
    
    # 1. Sombreado de filas alternas (Zebra striping)
    def zebra_stripe(data):
        df_styles = pd.DataFrame('', index=data.index, columns=data.columns)
        for i in range(len(data)):
            if i % 2 == 0:
                # Fondo gris muy claro para las filas pares
                df_styles.iloc[i] = 'background-color: #F4F6F9;' 
        return df_styles
        
    styler = styler.apply(zebra_stripe, axis=None)
    
    # 2. Color condicional para la columna 'Activo'
    if 'Activo' in df.columns:
        def color_activo(val):
            try:
                v = int(val)
                if v == 1: return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
                elif v == 0: return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;'
            except: pass
            return ''
        if hasattr(styler, 'map'): styler = styler.map(color_activo, subset=['Activo'])
        else: styler = styler.applymap(color_activo, subset=['Activo'])
            
    return styler

if archivo_subido is not None:
    wb = load_workbook(archivo_subido, data_only=True)
    df_tipos = leer_tabla_excel(wb, "tblTiposDeNodo")
    df_nodos_original = leer_tabla_excel(wb, "tblNodos")
    df_enlaces_original = leer_tabla_excel(wb, "tblEnlaces")

    st.markdown("---")
    st.subheader("🗄️ Base de Datos del Modelo (Simulador interactivo)")
    st.markdown("Modifica los parámetros y el grafo inferior reaccionará al instante.")

    tab_sim1, tab_sim2, tab_sim3 = st.tabs(["🏷️ Tipos de Nodo", "🟢 Nodos (Actores)", "🔗 Enlaces (Rutas)"])

    with tab_sim1: 
        st.markdown("**Diccionario de capas estructurales (Solo lectura)**")
        st.dataframe(aplicar_estilos(df_tipos), use_container_width=True)
    with tab_sim2:
        st.info("💡 Simula el arresto de un actor cambiando su columna **Activo** a 0.")
        cols_nodos = [c for c in ['Activo', 'NodoID', 'Nombre', 'Tipo', 'Descripción'] if c in df_nodos_original.columns]
        if not cols_nodos: cols_nodos = df_nodos_original.columns.tolist()
        df_nodos_editado = st.data_editor(aplicar_estilos(df_nodos_original[cols_nodos]), use_container_width=True, num_rows="dynamic", key="editor_nodos")
    with tab_sim3:
        st.info("💡 Edita la columna **Exposición** o desactiva rutas (1 -> 0).")
        cols_enlaces = [c for c in ['Activo', 'Nodo Origen', 'Nodo Destino', 'Tipo de Enlace', 'Exposición', 'Coste', 'Capacidad', 'Eficiencia'] if c in df_enlaces_original.columns]
        if not cols_enlaces: cols_enlaces = df_enlaces_original.columns.tolist()
        df_enlaces_editado = st.data_editor(aplicar_estilos(df_enlaces_original[cols_enlaces]), use_container_width=True, num_rows="dynamic", key="editor_enlaces")

    df_tipos = df_tipos.dropna(how='all')
    df_nodos = df_nodos_editado.dropna(how='all')
    df_enlaces = df_enlaces_editado.dropna(subset=['Nodo Origen', 'Nodo Destino'])

    df_nodos['Activo'] = pd.to_numeric(df_nodos['Activo'], errors='coerce')
    df_enlaces['Activo'] = pd.to_numeric(df_enlaces['Activo'], errors='coerce')

    df_nodos = pd.merge(df_nodos, df_tipos[['Tipo', 'Capa']], on='Tipo', how='left')
    df_nodos['Capa'] = pd.to_numeric(df_nodos['Capa'], errors='coerce').fillna(0)

    nodos_activos = df_nodos[df_nodos['Activo'] == 1]
    enlaces_activos = df_enlaces[df_enlaces['Activo'] == 1]

    G = nx.DiGraph()
    colores_capa = {'O': '#FF9999', 'I': '#99CCFF', 'G': '#FFCC99', 'D': '#99FF99'}

    for _, row in nodos_activos.iterrows():
        tipo_nodo = str(row['Tipo']).strip()
        nombre_nodo = str(row.get('Nombre', ''))
        G.add_node(str(row['NodoID']).strip(), label=str(row['NodoID']).strip() + (" - " + nombre_nodo if nombre_nodo else ""), color=colores_capa.get(tipo_nodo, '#CCCCCC'), level=int(row['Capa']), title=f"Tipo: {str(row.get('Descripción', tipo_nodo))}")

    for _, row in enlaces_activos.iterrows():
        origen = str(row['Nodo Origen']).strip()
        destino = str(row['Nodo Destino']).strip()
        if origen in G.nodes and destino in G.nodes:
            exposicion = str(row.get('Exposición', '')).strip().lower()
            if 'medio-alto' in exposicion: color_flecha = '#90EE90'
            elif 'medio-bajo' in exposicion: color_flecha = '#FFB84D'
            elif 'alto' in exposicion: color_flecha = '#008000'
            elif 'medio' in exposicion: color_flecha = '#FFD700'
            elif 'bajo' in exposicion: color_flecha = '#FF4444'
            else: color_flecha = '#999999'
            G.add_edge(origen, destino, color=color_flecha, title=f"Exposición: {row.get('Exposición', 'N/A')} | Coste: {row.get('Coste', 'N/A')}")

    # =========================================================
    # CÁLCULO MATEMÁTICO BASADO EN PANTALLA 1920x1080 (NATIVO)
    # =========================================================
    
    if not nodos_activos.empty:
        num_capas = nodos_activos['Capa'].nunique()
        max_nodos_por_capa = nodos_activos['Capa'].value_counts().max()
    else:
        num_capas = 1
        max_nodos_por_capa = 1
        
    ANCHO_CONTENEDOR = 1440
    ALTO_CONTENEDOR = 650
    
    huecos_horizontales = num_capas - 1 if num_capas > 1 else 1
    
    # Restamos un poco más de margen por si los textos crecen
    sep_horizontal = int((ANCHO_CONTENEDOR - 200) / huecos_horizontales)
    
    sep_vertical = int((ALTO_CONTENEDOR - 120) / max_nodos_por_capa)

    sep_horizontal = max(sep_horizontal, 250)
    sep_vertical = max(sep_vertical, 60) # Un poco más de aire vertical para letras grandes

    # 4. Inyección en PyVis nativo con TAMAÑO DE LETRA AUMENTADO (16)
    net = Network(height="650px", width="100%", directed=True, bgcolor="#ffffff", font_color="black")
    net.from_nx(G)
    
    net.set_options(f"""
    {{
      "layout": {{
        "hierarchical": {{
          "enabled": true,
          "direction": "LR",
          "sortMethod": "directed",
          "levelSeparation": {sep_horizontal},
          "nodeDistance": {sep_vertical},
          "treeSpacing": {sep_vertical},
          "blockShifting": false,
          "edgeMinimization": false,
          "parentCentralization": true
        }}
      }},
      "physics": {{
        "enabled": false
      }},
      "edges": {{
        "smooth": {{
          "type": "cubicBezier",
          "forceDirection": "horizontal",
          "roundness": 0.4
        }},
        "arrows": {{
          "to": {{ "enabled": true, "scaleFactor": 0.5 }}
        }}
      }},
      "nodes": {{
        "font": {{
            "size": 16,
            "face": "Arial"
        }}
      }},
      "interaction": {{
        "zoomView": true,
        "dragNodes": true,
        "hover": true
      }}
    }}
    """)

    ruta_html = "mapa_interactivo.html"
    net.save_graph(ruta_html)
    with open(ruta_html, 'r', encoding='utf-8') as f:
        codigo_html = f.read()

    st.markdown("---")
    gcol1, gcol2 = st.columns([4, 1])
    with gcol1:
        st.subheader("🌐 Topología Interactiva de la Red")
        components.html(codigo_html, height=670)

    with gcol2:
        st.subheader("📊 Análisis")
        st.metric("Total Nodos Activos", len(G.nodes))
        st.metric("Total Rutas Activas", len(G.edges))
        st.markdown("### Mapa de Calor (Exposición)")
        st.markdown("🔴 **Bajo:** Canal opaco\n🟠 **Medio-Bajo:** Riesgo latente\n🟡 **Medio:** Neutro\n🟢 **Medio-Alto / Alto:** Canal expuesto (Fricción)")
        st.info("💡 Modifica las tablas superiores para recalcular.")
