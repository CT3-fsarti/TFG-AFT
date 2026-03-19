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
    
    .stMarkdown h2 {
        font-size: 24px !important;
        color: #005691 !important;
        padding-bottom: 5px !important;
        font-weight: 600 !important;
        border-bottom: 2px solid #E6E9EF;
        margin-top: 20px;
    }
    
    .stMarkdown h3 {
        font-size: 18px !important;
        color: #333 !important;
        font-weight: 600 !important;
    }
    
    .centered-text { text-align: center; }
    .language-block { display: flex; align-items: flex-start; margin-bottom: 18px; }
    
    .flag-img { 
        width: 40px !important; 
        height: 26px !important; 
        object-fit: cover !important; 
        margin-right: 15px; 
        margin-top: 2px;
        border: 1px solid #ddd;
    }
    
    .flag-inline {
        height: 14px; 
        width: 22px; 
        object-fit: cover; 
        vertical-align: middle; 
        margin-right: 6px; 
        margin-bottom: 3px;
        border: 1px solid #ccc; 
        border-radius: 2px;
    }
    
    .title-text { font-size: 15px; color: #222; line-height: 1.5; }
    
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
    "Flag ES B64": get_base64_of_bin_file("logo_ES.png"),
    "Flag FR B64": get_base64_of_bin_file("logo_FR.png"),
    "Flag GB B64": get_base64_of_bin_file("logo_GB.png"),
    "Marina Photo B64": get_base64_of_bin_file("marina_circular.png")
}

html_flag_es = f"<img src='data:image/png;base64,{assets_status['Flag ES B64']}' class='flag-inline'>" if assets_status['Flag ES B64'] else ""
html_flag_fr = f"<img src='data:image/png;base64,{assets_status['Flag FR B64']}' class='flag-inline'>" if assets_status['Flag FR B64'] else ""
html_flag_gb = f"<img src='data:image/png;base64,{assets_status['Flag GB B64']}' class='flag-inline'>" if assets_status['Flag GB B64'] else ""

# ==========================================
# 2. ESTRUCTURA DE DISEÑO PRINCIPAL (COLUMNAS)
# ==========================================
col_main, col_profile = st.columns([3, 1])

with col_profile:
    st.markdown("""<div style="margin-top: 30px;"></div>""", unsafe_allow_html=True)
    with st.container():
        st.markdown("<h3 class='centered-text'>👤 Perfil de la Autora</h3>", unsafe_allow_html=True)
        
        if assets_status["Marina Photo B64"]:
             st.markdown(f"""<div style="margin-bottom: 15px;"><img src='data:image/png;base64,{assets_status["Marina Photo B64"]}' class='profile-photo' /></div>""", unsafe_allow_html=True)
        
        st.markdown(f"<p class='centered-text' style='font-size: 18px; margin-bottom: 5px;'><strong>Marina Sarti Pineda</strong></p>", unsafe_allow_html=True)
        
        st.markdown("""
        <div style="text-align: center; margin-bottom: 15px;">
            <a href="https://www.linkedin.com/in/marina-sarti-pineda-27211b29b/?originalSubdomain=es" target="_blank" style="text-decoration: none; font-size: 14px; color: #005691;">🔗 Perfil de LinkedIn</a>
        </div>
        """, unsafe_allow_html=True)
        
        # STACK TECNOLÓGICO
        st.markdown("""
        <hr style="margin: 10px 0 15px 0; border-top: 1px solid #E6E9EF;">
        <p class='centered-text' style='font-size: 15px; color: #333; margin-bottom: 10px;'><strong>🛠️ Stack Tecnológico</strong></p>
        <div style="display: flex; flex-wrap: wrap; gap: 6px; justify-content: center; margin-bottom: 20px;">
            <a href="https://www.python.org/" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">Python</a>
            <a href="https://networkx.org/" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">NetworkX</a>
            <a href="https://pandas.pydata.org/" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">Pandas (Excel)</a>
            <a href="https://streamlit.io/" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">Streamlit</a>
            <a href="https://pyvis.readthedocs.io/" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">PyVis</a>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style="font-size: 14px; line-height: 1.5; color: #333; text-align: justify;">
            <p style="margin-bottom: 8px;">{html_flag_es} Graduada en Economía por la <strong>Universidad Carlos III de Madrid y París Dauphine (DTI)</strong>, Marina combina el rigor analítico macroeconómico con una vocación por la seguridad global, la inteligencia financiera y la tecnología.<br><br>Su TFG pionero en Modelado Estructurado de Financiación del Terrorismo se desarrolló en el marco de las tipologías del GAFI.</p>
            <hr style="margin: 12px 0; border-top: 1px solid #eee;">
            <p style="margin-bottom: 8px;">{html_flag_fr} Diplômée en Économie de l'<strong>Universidad Carlos III de Madrid et Paris Dauphine (DTI)</strong>, Marina allie la rigueur analytique macroéconomique à une vocation pour la sécurité mondiale, le renseignement financier et la technologie.<br><br>Son mémoire de fin d'études pionnier sur la Modélisation Structurée du Financement du Terrorisme a été développé dans le cadre des typologies du GAFI.</p>
            <hr style="margin: 12px 0; border-top: 1px solid #eee;">
            <p>{html_flag_gb} Graduated in Economics from <strong>Universidad Carlos III de Madrid and Paris Dauphine (DTI)</strong>, Marina combines macroeconomic analytical rigor with a vocation for global security, financial intelligence, and technology.<br><br>Her pioneering Bachelor's Thesis in Structured Modeling of Terrorism Financing was developed within the framework of FATF typologies.</p>
        </div>
        """, unsafe_allow_html=True)

with col_main:
    col_logo, col_titles = st.columns([1, 2.5])
    with col_logo:
        if assets_status["Combined Logo B64"]:
             st.markdown(f"""<img src='data:image/png;base64,{assets_status["Combined Logo B64"]}' style='width: 100%; height: auto; margin-top: 10px;' />""", unsafe_allow_html=True)
    with col_titles:
        st.markdown("""
        <div style="text-align: left; padding-top: 10px;">
            <div style="margin-bottom: 6px;">
                <span style="font-style: italic; color: #555; font-size: 18px;">Trabajo Fin de Grado - </span>
                <strong style="color: #005691; font-size: 24px;">Doble Titulación Internacional en Economía.</strong>
            </div>
            <div style="color: #222; font-style: italic; font-weight: bold; font-size: 18px; line-height: 1.4;">
                Redes de Financiación del Terrorismo:<br>
                Simulación y análisis mediante Economía de Redes y Teoría de Juegos
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<hr style='margin-top: 15px; margin-bottom: 20px; border-top: 1px solid #111;'>", unsafe_allow_html=True)
    
    if assets_status["Flag FR B64"]: 
        st.markdown(f"""
        <div class="language-block">
            <img src="data:image/png;base64,{assets_status['Flag FR B64']}" class="flag-img" />
            <div class="title-text">
                Réseaux de Financement du Terrorisme:<br>
                Simulation et analyse à travers l'Économie des Réseaux et la Théorie des Jeux
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    if assets_status["Flag GB B64"]: 
        st.markdown(f"""
        <div class="language-block">
            <img src="data:image/png;base64,{assets_status['Flag GB B64']}" class="flag-img" />
            <div class="title-text">
                Terrorism Financing Networks:<br>
                Simulation and Analysis through Network Economics and Game Theory
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<hr style='margin-top: 15px; margin-bottom: 25px; border-top: 1px solid #111;'>", unsafe_allow_html=True)

    st.markdown("## 🎓 Sinopsis del Proyecto")
    with st.expander("Leer Sinopsis / Abstract Completo", expanded=True):
        tab_fr, tab_es, tab_en = st.tabs(["Français", "Español", "English"])
        with tab_fr: st.markdown("""<div style="font-size: 15px; text-align: justify; color: #333;">Le Financement du Terrorisme (FT) consiste en la collecte de fonds, ce qui englobe le processus de sollicitation, de rassemblement, de fourniture et de mise à disposition d'argent ou d'actifs dans le but de faciliter la capacité à mener des activités terroristes. En Espagne, la loi 10/2010 établit un cadre rigoureux contre la fourniture ou la distribution de fonds.<br><br>Les grands groupes organisés, les petites cellules et les acteurs individuels ont besoin d'argent pour développer leurs activités terroristes. La littérature académique s'accorde à dire que le manque de fonds limite considérablement leur capacité opérationnelle, le FT étant un élément structurant du terrorisme mondial.<br><br>Ce travail fonde son analyse sur des informations récentes, en utilisant des exemples représentatifs des dynamiques contemporaines. L'objectif principal est de construire une <strong>simulation d'un réseau de financement du terrorisme</strong> basée sur les preuves recueillies dans la littérature spécialisée (typologies du GAFI, ABE, etc.).<br><br>Une analyse structurelle est réalisée sur cette simulation à l'aide de l'Économie des Réseaux et de la Théorie des Jeux, incluant l'étude des métriques de centralité, l'importance des nœuds clés (points d'étranglement) et la résilience du système face aux interventions policières. Le modèle qui en résulte constitue une représentation réaliste et empiriquement fondée, aboutissant à un outil analytique très utile pour le renseignement financier.</div>""", unsafe_allow_html=True)
        with tab_es: st.markdown("""<div style="font-size: 15px; text-align: justify; color: #333;">La Financiación del Terrorismo (FT) consiste en la captación de fondos, lo cual abarca el proceso de solicitud, recaudación, provisión y puesta a disposición de dinero o activos con el fin de facilitar o potenciar la capacidad de cualquier persona u organización para llevar a cabo actividades relacionadas con el terrorismo. En el caso concreto de España, la Ley 10/2010 establece un marco riguroso frente al suministro, depósito o distribución de fondos.<br><br>Tanto los grandes grupos organizados como las pequeñas células e incluso los actores individuales necesitan dinero para el desarrollo de la actividad terrorista. La literatura académica y los informes institucionales coinciden en que la falta de fondos limita drásticamente su capacidad operativa, siendo la FT un elemento vertebrador del terrorismo global.<br><br>Este trabajo fundamenta su análisis en información reciente, utilizando ejemplos observados en los últimos años que resultan representativos de las dinámicas contemporáneas. El objetivo principal consiste en construir una <strong>simulación de red de financiación del terrorismo</strong> basada en la evidencia recogida en la literatura especializada (tipologías del GAFI, EBA, etc.).<br><br>Sobre dicha simulación se realiza un análisis estructural mediante herramientas propias del análisis de redes (Economía de Redes) y la Teoría de Juegos, incluyendo el estudio de métricas de centralidad, la importancia relativa de los nodos clave (chokepoints) y la resiliencia del sistema ante intervenciones policiales. El modelo resultante constituye una representación realista y fundamentada empíricamente, derivando en un modelo analítico altamente útil para la inteligencia financiera y el diseño de políticas de seguridad.</div>""", unsafe_allow_html=True)
        with tab_en: st.markdown("""<div style="font-size: 15px; text-align: justify; color: #333;">Terrorist Financing (TF) involves the raising of funds, which encompasses the process of soliciting, collecting, providing, and making available money or assets to facilitate or enhance the capacity of any individual or organization to carry out terrorist activities. In Spain, Law 10/2010 establishes a rigorous framework against the supply, deposit, or distribution of funds.<br><br>Large organized groups, small cells, and lone actors require money to carry out terrorist activities. Academic literature and institutional reports agree that a lack of funds drastically limits their operational capacity, making TF a structural backbone of global terrorism.<br><br>This paper bases its analysis on recent information, using examples observed in recent years that are representative of contemporary dynamics. The main objective is to build a <strong>simulation of a terrorist financing network</strong> based on evidence gathered from specialized literature (FATF typologies, EBA, etc.).<br><br>A structural analysis is performed on this simulation using Network Economics and Game Theory, including the study of centrality metrics, the relative importance of key nodes (chokepoints), and the system's resilience to law enforcement interventions. The resulting model constitutes a realistic and empirically grounded representation, resulting in an analytical model highly useful for financial intelligence and security policy design.</div>""", unsafe_allow_html=True)

# ==========================================
# FLUJO DE SIMULACIÓN (3 FASES)
# ==========================================

def leer_tabla_excel(wb, nombre_tabla_buscada):
    for hoja in wb.worksheets:
        if nombre_tabla_buscada in hoja.tables:
            rango = hoja.tables[nombre_tabla_buscada].ref
            datos_celdas = hoja[rango]
            filas = [[celda.value for celda in fila] for fila in datos_celdas]
            return pd.DataFrame(filas[1:], columns=filas[0])
    return pd.DataFrame() # Devuelve vacío si no encuentra la tabla para no romper la app

def aplicar_estilos(df):
    styler = df.style.set_properties(**{'text-align': 'center'})
    styler = styler.set_table_styles([dict(selector='th', props=[('text-align', 'center')])])
    
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

st.markdown("---")
st.markdown("<div style='text-align: center; color: #555; font-size: 14px;'>FLUJO DEL SIMULADOR</div>", unsafe_allow_html=True)

archivo_subido = st.file_uploader("Sube tu archivo Excel alternativo (Opcional)", type=["xlsx"])

wb = None
if archivo_subido is not None:
    wb = load_workbook(archivo_subido, data_only=True)
elif os.path.exists("Diseño Red FT.xlsx"):
    wb = load_workbook("Diseño Red FT.xlsx", data_only=True)
else:
    st.info("💡 Sube un archivo Excel para comenzar, o asegúrate de que el archivo 'Diseño Red FT.xlsx' esté en la misma carpeta que este script para que se cargue automáticamente.")

if wb is not None:
    # Extracción de todas las tablas
    df_tipos = leer_tabla_excel(wb, "tblTiposDeNodo")
    df_nodos_original = leer_tabla_excel(wb, "tblNodos")
    df_enlaces_original = leer_tabla_excel(wb, "tblEnlaces")
    df_pesos = leer_tabla_excel(wb, "tblPesos")
    df_ponderada_excel = leer_tabla_excel(wb, "tblPonderada")

    # ==========================================
    # FASE 1: BASE DE DATOS Y PARAMETRIZACIÓN
    # ==========================================
    st.markdown("## ⚙️ Fase 1: Base de Datos y Parametrización")
    st.markdown("Revisa las entidades, rutas y la metodología de pesos aplicada al modelo.")

    tab_sim1, tab_sim2, tab_sim3, tab_sim4 = st.tabs(["🟢 Nodos (Actores)", "🔗 Enlaces (Rutas)", "🏷️ Tipos de Nodo", "⚖️ Matriz de Pesos (Metodología)"])

    with tab_sim1:
        st.info("💡 Simula el arresto de un actor cambiando su columna **Activo** a 0.")
        cols_nodos = [c for c in ['Activo', 'NodoID', 'Nombre', 'Tipo', 'Descripción'] if c in df_nodos_original.columns]
        if not cols_nodos: cols_nodos = df_nodos_original.columns.tolist()
        df_nodos_editado = st.data_editor(aplicar_estilos(df_nodos_original[cols_nodos]), use_container_width=True, num_rows="dynamic", key="editor_nodos")
    with tab_sim2:
        st.info("💡 Edita la columna **Exposición** o desactiva rutas (1 -> 0).")
        cols_enlaces = [c for c in ['Activo', 'Nodo Origen', 'Nodo Destino', 'Tipo de Enlace', 'Exposición', 'Coste', 'Capacidad', 'Eficiencia'] if c in df_enlaces_original.columns]
        if not cols_enlaces: cols_enlaces = df_enlaces_original.columns.tolist()
        df_enlaces_editado = st.data_editor(aplicar_estilos(df_enlaces_original[cols_enlaces]), use_container_width=True, num_rows="dynamic", key="editor_enlaces")
    with tab_sim3: 
        st.dataframe(aplicar_estilos(df_tipos), use_container_width=True)
    with tab_sim4:
        st.markdown("**Sistema de escalas directa e inversa:** Esta tabla define cómo los atributos cualitativos se traducen matemáticamente para calcular el 'camino de menor resistencia' (fricción).")
        if not df_pesos.empty:
            st.dataframe(aplicar_estilos(df_pesos), use_container_width=True, hide_index=True)
        else:
            st.warning("No se ha encontrado la tabla 'tblPesos' en el archivo Excel.")

    # Procesamiento interno
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

    # ==========================================
    # FASE 2: TOPOLOGÍA DE LA RED
    # ==========================================
    st.markdown("## 🌐 Fase 2: Topología Interactiva de la Red")
    
    if not nodos_activos.empty:
        num_capas = nodos_activos['Capa'].nunique()
        max_nodos_por_capa = nodos_activos['Capa'].value_counts().max()
    else:
        num_capas, max_nodos_por_capa = 1, 1
        
    sep_horizontal = max(int((1440 - 200) / (num_capas - 1 if num_capas > 1 else 1)), 250)
    sep_vertical = max(int((650 - 120) / max_nodos_por_capa), 60)

    net = Network(height="650px", width="100%", directed=True, bgcolor="#ffffff", font_color="black")
    net.from_nx(G)
    net.set_options(f"""
    {{
      "layout": {{ "hierarchical": {{ "enabled": true, "direction": "LR", "sortMethod": "directed", "levelSeparation": {sep_horizontal}, "nodeDistance": {sep_vertical}, "treeSpacing": {sep_vertical}, "parentCentralization": true }} }},
      "physics": {{ "enabled": false }},
      "edges": {{ "smooth": {{ "type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.4 }}, "arrows": {{ "to": {{ "enabled": true, "scaleFactor": 0.5 }} }} }},
      "nodes": {{ "font": {{ "size": 16, "face": "Arial" }} }},
      "interaction": {{ "zoomView": true, "dragNodes": true, "hover": true }}
    }}
    """)

    net.save_graph("mapa_interactivo.html")
    with open("mapa_interactivo.html", 'r', encoding='utf-8') as f:
        codigo_html = f.read()

    gcol1, gcol2 = st.columns([4, 1])
    with gcol1:
        components.html(codigo_html, height=670)

    with gcol2:
        st.subheader("📊 Análisis")
        st.metric("Total Nodos Activos", len(G.nodes))
        st.metric("Total Rutas Activas", len(G.edges))
        
        st.markdown("### Mapa de Calor (Exposición)")
        st.markdown("""
        <div style="line-height: 2; font-size: 15px;">
            🔴 <strong>Bajo:</strong> Canal opaco<br>
            🟠 <strong>Medio-Bajo:</strong> Riesgo latente<br>
            🟡 <strong>Medio:</strong> Neutro<br>
            🟢 <strong>Medio-Alto / Alto:</strong> Canal expuesto (Fricción)
        </div>
        """, unsafe_allow_html=True)

    # ==========================================
    # FASE 3: MODELADO MATEMÁTICO
    # ==========================================
    st.markdown("## 🔢 Fase 3: Modelado Matemático (Teoría de Grafos)")
    st.markdown("Traducción algebraica de la topología superior para el cálculo de equilibrio y cuellos de botella.")

    # Cambio: expanded=True para que cargue abierto.
    with st.expander("Matrices Matemáticas del Sistema (Teoría de Grafos)", expanded=True):
        tab_matriz1, tab_matriz2 = st.tabs(["1️⃣ Matriz de Adyacencia (Topológica)", "2️⃣ Matriz Ponderada (Fricción)"])

        with tab_matriz1:
            st.markdown("**Matriz Binaria:** Representa la existencia de rutas (1 = conectado). Es la base estructural para calcular la centralidad de grado.")
            if len(G.nodes) > 0:
                matriz_adyacencia = nx.to_pandas_adjacency(G, dtype=int)
                matriz_vacia_ady = matriz_adyacencia.replace(0, "")
                st.dataframe(matriz_vacia_ady, use_container_width=True)
            else:
                st.info("La red está vacía.")

        with tab_matriz2:
            st.markdown("**Matriz Ponderada:** Refleja la fricción o capacidad de los canales basándose en los parámetros de la tabla 'tblPonderada' de Excel.")
            if not df_ponderada_excel.empty:
                # Ponemos el índice de la matriz ponderada basándonos en la primera columna
                df_ponderada_excel.set_index(df_ponderada_excel.columns[0], inplace=True)
                # Reemplazamos los ceros (o vacíos) para limpiarla visualmente
                matriz_pond_limpia = df_ponderada_excel.replace(0, "")
                matriz_pond_limpia = matriz_pond_limpia.fillna("")
                st.dataframe(matriz_pond_limpia, use_container_width=True)
            else:
                st.warning("No se ha encontrado la tabla 'tblPonderada' en el archivo Excel.")
