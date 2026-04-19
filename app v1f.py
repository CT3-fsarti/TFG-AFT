import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import networkx as nx
from pyvis.network import Network
from openpyxl import load_workbook
import os
import base64
import json
from io import BytesIO

from ft_excel_bridge import INPUT_TABLES, WorkbookSchemaError, build_artifacts, load_table_frame, prepare_base_tables

# ==========================================
# ESCUDO PROTECTOR PARA VERTEX AI
# Evita que la app colapse si falta configurar algo
# ==========================================
try:
    from google.oauth2 import service_account
    import vertexai
    from vertexai.generative_models import GenerativeModel
    VERTEX_AVAILABLE = True
except ImportError:
    VERTEX_AVAILABLE = False

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
            <a href="https://cloud.google.com/vertex-ai" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">Vertex AI</a>
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
# 3. FUNCIONES DE CARGA Y ESTILADO
# ==========================================
SCENARIO_DEFINITIONS = [
    {
        "id": "estrella",
        "label": "Estrella",
        "description": "Activa o desactiva nodos y enlaces para quedarte con una topologia centralizada.",
    },
    {
        "id": "anillo",
        "label": "Anillo",
        "description": "Usa la Meta-Red como base y deja activas solo las rutas que formen un circuito.",
    },
    {
        "id": "multi_hub",
        "label": "Multi-Hub",
        "description": "Construye una red con varios nodos concentradores comparables entre si.",
    },
]

DEFAULT_WORKBOOKS = ["Diseño Red FT v5a.xlsm", "Diseño Red FT.xlsx"]
METRIC_COLUMN_ORDER = [
    "Distancia media al destino final D1",
    "Distancia media total de la red",
    "Cercania armonica media de la red",
    "Eficiencia global de la red",
    "Centralizacion de la red",
]


def aplicar_estilos(df_in):
    """Estilado para las tablas de la Fase 1: Todo centrado y números a 0 decimales"""
    df_safe = df_in.copy()
    
    df_safe.columns = df_safe.columns.astype(str)
    df_safe.loc[:, ~df_safe.columns.duplicated()]

    styler = df_safe.style.set_properties(**{'text-align': 'center'})
    styler = styler.set_table_styles([
        dict(selector='th', props=[('text-align', 'center')]),
        dict(selector='td', props=[('text-align', 'center')])
    ])
    
    styler = styler.format(lambda v: f"{v:.0f}" if isinstance(v, (int, float)) and pd.notna(v) else v)
    
    if 'Activo' in df_safe.columns:
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

def aplicar_estilo_matriz(df_in, decimales=0):
    """Estilado para matrices de la Fase 3: Todo centrado, ocultación de 0s y control de decimales"""
    df_safe = df_in.copy()
    
    df_safe.index = df_safe.index.astype(str)
    df_safe.columns = df_safe.columns.astype(str)
    df_safe = df_safe.loc[~df_safe.index.duplicated(keep='first')]
    df_safe = df_safe.loc[:, ~df_safe.columns.duplicated()]

    styler = df_safe.style.set_properties(**{'text-align': 'center'})
    styler = styler.set_table_styles([
        dict(selector='th', props=[('text-align', 'center')]),
        dict(selector='td', props=[('text-align', 'center')])
    ])
    
    def formato_celda(v):
        if pd.isna(v) or v == 0 or v == "0" or v == "": 
            return ""
        try: 
            return f"{float(v):.{decimales}f}"
        except: 
            return str(v)
        
    return styler.format(formato_celda)


def copiar_base(base_modelo):
    return {clave: df.copy(deep=True) for clave, df in base_modelo.items()}


def buscar_workbook_por_defecto():
    for candidate in DEFAULT_WORKBOOKS:
        if os.path.exists(candidate):
            return candidate
    return None


@st.cache_data(show_spinner=False)
def cargar_modelo_desde_bytes(workbook_bytes):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, keep_vba=True)
    frames = {
        alias: load_table_frame(workbook, table_name)
        for alias, table_name in INPUT_TABLES.items()
    }
    return prepare_base_tables(frames)


def construir_editor_nodos(base_modelo):
    capas = base_modelo["tipos_nodo"][["tipo", "descripcion", "capa"]].drop_duplicates()
    nodos = base_modelo["nodos"].merge(capas, on="tipo", how="left")
    vista = nodos[["activo", "nodoid", "nombre", "tipo", "descripcion", "capa"]].copy()
    vista["activo"] = vista["activo"].fillna(0).astype(int).eq(1)
    vista = vista.rename(
        columns={
            "activo": "Activo",
            "nodoid": "NodoID",
            "nombre": "Nombre",
            "tipo": "Tipo",
            "descripcion": "Descripcion",
            "capa": "Capa",
        }
    )
    vista.index = vista["NodoID"].astype(str)
    return vista


def construir_editor_enlaces(base_modelo):
    nombres_nodo = base_modelo["nodos"].set_index("nodoid")["nombre"].to_dict()
    enlaces = base_modelo["enlaces"][["activo", "nodo_inicial", "nodo_final", "canal"]].copy()
    enlaces["nombre_origen"] = enlaces["nodo_inicial"].map(nombres_nodo)
    enlaces["nombre_destino"] = enlaces["nodo_final"].map(nombres_nodo)
    vista = enlaces[["activo", "nodo_inicial", "nombre_origen", "nodo_final", "nombre_destino", "canal"]].copy()
    vista["activo"] = vista["activo"].fillna(0).astype(int).eq(1)
    vista = vista.rename(
        columns={
            "activo": "Activo",
            "nodo_inicial": "Nodo Inicial",
            "nombre_origen": "Nombre Origen",
            "nodo_final": "Nodo Final",
            "nombre_destino": "Nombre Destino",
            "canal": "Canal",
        }
    )
    vista.index = base_modelo["enlaces"].index
    return vista


def aplicar_activos(base_modelo, nodos_editados, enlaces_editados):
    escenario = copiar_base(base_modelo)
    nodos_activos = nodos_editados["Activo"].astype(bool).astype(int).to_dict()
    enlaces_activos = enlaces_editados["Activo"].astype(bool).astype(int).to_dict()

    escenario["nodos"]["activo"] = (
        escenario["nodos"]["nodoid"].map(nodos_activos).fillna(0).astype(int)
    )
    escenario["enlaces"]["activo"] = (
        escenario["enlaces"].index.to_series().map(enlaces_activos).fillna(0).astype(int)
    )
    return escenario


def construir_grafo(base_modelo, red):
    colores_tipo = {
        "O": "#D66D75",
        "C": "#E8A66B",
        "I": "#5DA9E9",
        "G": "#7BC47F",
        "D": "#6C5CE7",
    }
    tipos = base_modelo["tipos_nodo"][["tipo", "descripcion", "capa"]].drop_duplicates()
    nodos = base_modelo["nodos"].merge(tipos, on="tipo", how="left")
    nodos_activos = nodos[nodos["activo"] == 1].copy()
    enlaces_activos = red[red["activo"] == 1].copy()

    grafo = nx.DiGraph()
    for row in nodos_activos.itertuples(index=False):
        nombre = str(row.nombre).strip() if pd.notna(row.nombre) else ""
        grafo.add_node(
            str(row.nodoid).strip(),
            label=str(row.nodoid).strip() + (f" - {nombre}" if nombre else ""),
            color=colores_tipo.get(str(row.tipo).strip(), "#BFC8D6"),
            level=int(row.capa) if pd.notna(row.capa) else 0,
            title=(
                f"Tipo: {row.tipo}<br>"
                f"Descripcion: {row.descripcion if pd.notna(row.descripcion) else ''}"
            ),
        )

    for row in enlaces_activos.itertuples(index=False):
        tradeoff = float(row.trade_off_valor_operativo_coste) if pd.notna(row.trade_off_valor_operativo_coste) else 0.0
        if tradeoff >= 3:
            color = "#1B9E77"
        elif tradeoff >= 2:
            color = "#66A61E"
        elif tradeoff >= 1:
            color = "#E6AB02"
        else:
            color = "#D95F02"
        grafo.add_edge(
            str(row.nodo_inicial).strip(),
            str(row.nodo_final).strip(),
            color=color,
            title=(
                f"Coste: {row.coste:.3f}<br>"
                f"Valor operativo: {row.valor_operativo:.3f}<br>"
                f"Trade-off: {tradeoff:.3f}"
            ),
        )
    return grafo


def renderizar_grafo(grafo, scenario_id):
    if len(grafo.nodes) == 0:
        st.info("No hay nodos activos en este escenario.")
        return

    niveles = [data.get("level", 0) for _, data in grafo.nodes(data=True)]
    num_capas = len(set(niveles)) if niveles else 1
    max_nodos_por_capa = max(pd.Series(niveles).value_counts()) if niveles else 1

    sep_horizontal = max(int((1440 - 200) / (num_capas - 1 if num_capas > 1 else 1)), 250)
    sep_vertical = max(int((650 - 120) / max_nodos_por_capa), 60)

    net = Network(height="650px", width="100%", directed=True, bgcolor="#FFFFFF", font_color="#14213D")
    net.from_nx(grafo)
    net.set_options(
        f"""
        {{
          "layout": {{ "hierarchical": {{ "enabled": true, "direction": "LR", "sortMethod": "directed", "levelSeparation": {sep_horizontal}, "nodeDistance": {sep_vertical}, "treeSpacing": {sep_vertical}, "parentCentralization": true }} }},
          "physics": {{ "enabled": false }},
          "edges": {{ "smooth": {{ "type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.35 }}, "arrows": {{ "to": {{ "enabled": true, "scaleFactor": 0.55 }} }} }},
          "nodes": {{ "font": {{ "size": 15, "face": "Arial" }}, "borderWidth": 1.5, "shape": "dot", "size": 18 }},
          "interaction": {{ "zoomView": true, "dragNodes": true, "hover": true }}
        }}
        """
    )
    components.html(net.generate_html(name=f"topologia_{scenario_id}.html"), height=680, scrolling=True)


def obtener_resumen_metricas(metricas_red):
    metricas = metricas_red[["metrica", "valor"]].copy()
    metricas["valor"] = pd.to_numeric(metricas["valor"], errors="coerce").fillna(0.0)
    return metricas.set_index("metrica")["valor"].to_dict()


def construir_fila_comparacion(nombre_topologia, base_modelo, artifacts):
    metricas = obtener_resumen_metricas(artifacts["metricas_red"])
    fila = {
        "Topologia": nombre_topologia,
        "Nodos activos": int((base_modelo["nodos"]["activo"] == 1).sum()),
        "Enlaces activos": int((artifacts["red"]["activo"] == 1).sum()),
    }
    for metric_name in METRIC_COLUMN_ORDER:
        fila[metric_name] = float(metricas.get(metric_name, 0.0))
    return fila


def renderizar_metricas_resumen(metricas_red):
    metricas = obtener_resumen_metricas(metricas_red)
    cols = st.columns(len(METRIC_COLUMN_ORDER))
    for col, metric_name in zip(cols, METRIC_COLUMN_ORDER):
        col.metric(metric_name, f"{metricas.get(metric_name, 0.0):.4f}")


def renderizar_detalle_escenario(base_modelo, artifacts, scenario_id):
    grafo = construir_grafo(base_modelo, artifacts["red"])
    info_col, metric_col = st.columns([4, 1])
    with info_col:
        renderizar_grafo(grafo, scenario_id)
    with metric_col:
        st.subheader("Actividad")
        st.metric("Nodos activos", int((base_modelo["nodos"]["activo"] == 1).sum()))
        st.metric("Enlaces activos", int((artifacts["red"]["activo"] == 1).sum()))
        st.metric("Rutas visibles", len(grafo.edges))

    renderizar_metricas_resumen(artifacts["metricas_red"])

    with st.expander("Matrices y metricas del escenario", expanded=False):
        tab_red, tab_ady, tab_coste, tab_valor, tab_trade, tab_dist, tab_nodos, tab_red_metric = st.tabs([
            "Red agregada",
            "Adyacencia",
            "Costes",
            "Valor operativo",
            "Trade-Off",
            "Distancias minimas",
            "Metricas por nodo",
            "Metricas de red",
        ])

        with tab_red:
            st.dataframe(aplicar_estilos(artifacts["red"]), use_container_width=True, hide_index=True)
        with tab_ady:
            tabla = artifacts["matriz_adyacencia"].set_index("nodo_orig_dest")
            st.dataframe(aplicar_estilo_matriz(tabla, 0), use_container_width=True)
        with tab_coste:
            tabla = artifacts["matriz_costes"].set_index("nodo_orig_dest")
            st.dataframe(aplicar_estilo_matriz(tabla, 2), use_container_width=True)
        with tab_valor:
            tabla = artifacts["matriz_valor_operativo"].set_index("nodo_orig_dest")
            st.dataframe(aplicar_estilo_matriz(tabla, 2), use_container_width=True)
        with tab_trade:
            tabla = artifacts["matriz_tradeoff"].set_index("nodo_orig_dest")
            st.dataframe(aplicar_estilo_matriz(tabla, 2), use_container_width=True)
        with tab_dist:
            tabla = artifacts["matriz_distancias_minimas"].set_index("nodo_orig_dest")
            st.dataframe(aplicar_estilo_matriz(tabla, 3), use_container_width=True)
        with tab_nodos:
            st.dataframe(aplicar_estilos(artifacts["metricas_nodos"]), use_container_width=True, hide_index=True)
        with tab_red_metric:
            st.dataframe(aplicar_estilos(artifacts["metricas_red"]), use_container_width=True, hide_index=True)


def obtener_estado_editor(clave_estado, dataframe_default):
    dataframe_estado = st.session_state.get(clave_estado)
    if isinstance(dataframe_estado, pd.DataFrame):
        return dataframe_estado.copy(deep=True)
    dataframe_inicial = dataframe_default.copy(deep=True)
    st.session_state[clave_estado] = dataframe_inicial
    return dataframe_inicial


def renderizar_pestana_topologia(base_modelo, scenario_id, description, editable):
    st.caption(description)
    nodos_default = construir_editor_nodos(base_modelo)
    enlaces_default = construir_editor_enlaces(base_modelo)
    clave_estado_nodos = f"estado_nodos_{scenario_id}"
    clave_estado_enlaces = f"estado_enlaces_{scenario_id}"
    nodos_actuales = obtener_estado_editor(clave_estado_nodos, nodos_default)
    enlaces_actuales = obtener_estado_editor(clave_estado_enlaces, enlaces_default)

    if editable:
        columnas_nodos_bloqueadas = ["NodoID", "Nombre", "Tipo", "Descripcion", "Capa"]
        columnas_enlaces_bloqueadas = ["Nodo Inicial", "Nombre Origen", "Nodo Final", "Nombre Destino", "Canal"]
        texto_nodos = "Activa o desactiva nodos para construir la topologia sobre la Meta-Red."
        texto_enlaces = "Activa o desactiva enlaces concretos para cerrar o abrir rutas dentro del escenario."
        titulo_configuracion = "Configuracion del escenario"
    else:
        columnas_nodos_bloqueadas = list(nodos_default.columns)
        columnas_enlaces_bloqueadas = list(enlaces_default.columns)
        texto_nodos = "Referencia base de nodos de la Meta-Red."
        texto_enlaces = "Referencia base de enlaces de la Meta-Red."
        titulo_configuracion = "Configuracion de referencia"

    escenario_base = aplicar_activos(base_modelo, nodos_actuales, enlaces_actuales) if editable else copiar_base(base_modelo)
    artifacts = build_artifacts(escenario_base)
    renderizar_detalle_escenario(escenario_base, artifacts, scenario_id)

    with st.expander(titulo_configuracion, expanded=False):
        subtab1, subtab2 = st.tabs(["Nodos", "Enlaces"])
        with subtab1:
            st.markdown(texto_nodos)
            nodos_editados = st.data_editor(
                nodos_actuales,
                key=f"editor_nodos_{scenario_id}",
                hide_index=True,
                use_container_width=True,
                disabled=columnas_nodos_bloqueadas,
                column_config={
                    "Activo": st.column_config.CheckboxColumn("Activo"),
                },
            )
        with subtab2:
            st.markdown(texto_enlaces)
            enlaces_editados = st.data_editor(
                enlaces_actuales,
                key=f"editor_enlaces_{scenario_id}",
                hide_index=True,
                use_container_width=True,
                disabled=columnas_enlaces_bloqueadas,
                column_config={
                    "Activo": st.column_config.CheckboxColumn("Activo"),
                },
            )

    if not nodos_editados.equals(nodos_actuales) or not enlaces_editados.equals(enlaces_actuales):
        st.session_state[clave_estado_nodos] = nodos_editados.copy(deep=True)
        st.session_state[clave_estado_enlaces] = enlaces_editados.copy(deep=True)
        st.rerun()

    return escenario_base, artifacts

# ==========================================
# 4. FLUJO DE DATOS
# ==========================================
st.markdown("---")
st.markdown("<div style='text-align: center; color: #555; font-size: 14px;'>FLUJO DEL SIMULADOR</div>", unsafe_allow_html=True)

archivo_subido = st.file_uploader("Sube tu archivo Excel alternativo (Opcional)", type=["xlsm", "xlsx"])

workbook_bytes = None
workbook_source = None
if archivo_subido is not None:
    workbook_bytes = archivo_subido.getvalue()
    workbook_source = archivo_subido.name
else:
    ruta_por_defecto = buscar_workbook_por_defecto()
    if ruta_por_defecto:
        with open(ruta_por_defecto, "rb") as workbook_file:
            workbook_bytes = workbook_file.read()
        workbook_source = ruta_por_defecto
    else:
        st.info("💡 Sube un archivo Excel para comenzar, o deja en la carpeta del proyecto 'Diseño Red FT v5a.xlsm'.")

if workbook_bytes is not None:
    try:
        modelo_base = cargar_modelo_desde_bytes(workbook_bytes)
        meta_base = copiar_base(modelo_base)
        meta_artifacts = build_artifacts(meta_base)

        st.markdown("## 🧭 Comparador de Topologías sobre la Meta-Red")
        st.caption(f"Fuente cargada: {workbook_source}")
        st.markdown(
            "Cada topología se define encendiendo o apagando la columna `Activo` de Nodos y Enlaces. "
            "La Meta-Red actúa como referencia fija y las demás pestañas recalculan toda la red desde Python."
        )

        pestañas = st.tabs([
            "Meta-Red",
            *[scenario["label"] for scenario in SCENARIO_DEFINITIONS],
            "Comparacion",
        ])

        resultados = {
            "meta_red": {
                "label": "Meta-Red",
                "base": meta_base,
                "artifacts": meta_artifacts,
            }
        }

        with pestañas[0]:
            st.markdown("## Meta-Red")
            meta_base, meta_artifacts = renderizar_pestana_topologia(
                meta_base,
                "meta_red",
                "Escenario base del Excel. Se muestra como referencia fija para comparar el resto de topologías.",
                editable=False,
            )
            resultados["meta_red"] = {
                "label": "Meta-Red",
                "base": meta_base,
                "artifacts": meta_artifacts,
            }

        for index, scenario in enumerate(SCENARIO_DEFINITIONS, start=1):
            with pestañas[index]:
                st.markdown(f"## {scenario['label']}")
                escenario_base, artifacts = renderizar_pestana_topologia(
                    meta_base,
                    scenario["id"],
                    scenario["description"],
                    editable=True,
                )
                resultados[scenario["id"]] = {
                    "label": scenario["label"],
                    "base": escenario_base,
                    "artifacts": artifacts,
                }

        with pestañas[-1]:
            st.markdown("## Comparacion de metricas")
            filas = [
                construir_fila_comparacion(resultado["label"], resultado["base"], resultado["artifacts"])
                for resultado in resultados.values()
            ]
            comparacion_df = pd.DataFrame(filas).set_index("Topologia")
            st.dataframe(
                comparacion_df.style.format(
                    {
                        "Nodos activos": "{:.0f}",
                        "Enlaces activos": "{:.0f}",
                        **{metric_name: "{:.4f}" for metric_name in METRIC_COLUMN_ORDER},
                    }
                ),
                use_container_width=True,
            )

            if "Meta-Red" in comparacion_df.index:
                referencia = comparacion_df.loc["Meta-Red", METRIC_COLUMN_ORDER]
                delta_df = comparacion_df[METRIC_COLUMN_ORDER].subtract(referencia, axis=1)
                delta_df = delta_df.rename_axis("Topologia")
                st.markdown("### Delta respecto a la Meta-Red")
                st.dataframe(delta_df.style.format("{:+.4f}"), use_container_width=True)

            metrica_chart = st.selectbox(
                "Metrica a visualizar",
                METRIC_COLUMN_ORDER,
                key="metrica_comparacion",
            )
            st.bar_chart(comparacion_df[[metrica_chart]], use_container_width=True)

        if VERTEX_AVAILABLE:
            st.markdown("## 🤖 Asistente de Inteligencia Artificial (Vertex AI)")
            st.markdown("Consulta a Gemini sobre el escenario base y las diferencias entre topologías.")
            try:
                if "gcp_service_account_json" not in st.secrets:
                    st.warning("⚠️ El asistente está pausado. Faltan las credenciales en Streamlit Secrets.")
                else:
                    credenciales_json = json.loads(st.secrets["gcp_service_account_json"], strict=False)
                    credenciales_gcp = service_account.Credentials.from_service_account_info(credenciales_json)
                    vertexai.init(project="aft-simulator", location="us-central1", credentials=credenciales_gcp)
                    model = GenerativeModel("gemini-2.5-pro")

                    user_query = st.chat_input("Pregunta a Gemini 2.5 Pro sobre la Meta-Red o la comparativa...")
                    if user_query:
                        st.chat_message("user").write(user_query)
                        comparacion_texto = comparacion_df.to_string()
                        contexto_red = f"""
                        Eres un asistente de inteligencia financiera.
                        Escenario de referencia: Meta-Red.
                        Comparativa actual de topologías:
                        {comparacion_texto}

                        Pregunta del analista: {user_query}
                        """
                        with st.chat_message("assistant"):
                            with st.spinner("Analizando la comparativa con Gemini 2.5 Pro en Google Cloud..."):
                                response = model.generate_content(contexto_red)
                                st.write(response.text)
            except json.JSONDecodeError as e:
                st.error("❌ Error leyendo el archivo JSON de Streamlit Secrets.")
                st.code(str(e))
            except Exception as e:
                st.error("❌ Ha ocurrido un error al conectar con Vertex AI.")
                with st.expander("Ver detalles técnicos del error"):
                    st.code(str(e))

    except WorkbookSchemaError as exc:
        st.error(f"❌ El workbook no cumple el esquema esperado: {exc}")
    except KeyError as exc:
        st.error(f"❌ Falta una tabla requerida en el Excel: {exc}")
    except Exception as exc:
        st.error("❌ No se pudo cargar el modelo de la Meta-Red.")
        with st.expander("Ver detalles técnicos del error"):
            st.code(str(exc))