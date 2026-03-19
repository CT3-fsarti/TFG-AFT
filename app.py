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
    except: return None

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
# 2. ESTRUCTURA DE DISEÑO (COLUMNAS)
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
            <a href="https://www.linkedin.com/in/marina-sarti-pineda-27211b29b/" target="_blank" style="text-decoration: none; font-size: 14px; color: #005691;">🔗 Perfil de LinkedIn</a>
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
            <a href="https://gemini.google.com/" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">Gemini</a>
            <a href="https://github.com/features/copilot" target="_blank" style="text-decoration: none; background-color: #F0F4F8; color: #005691; padding: 3px 10px; border-radius: 12px; font-size: 12.5px; border: 1px solid #D9E2EC; font-weight: 500;">Copilot</a>
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
        st.markdown("""<div style="text-align: left; padding-top: 10px;"><div style="margin-bottom: 6px;"><span style="font-style: italic; color: #555; font-size: 18px;">Trabajo Fin de Grado - </span><strong style="color: #005691; font-size: 24px;">Doble Titulación Internacional en Economía.</strong></div><div style="color: #222; font-style: italic; font-weight: bold; font-size: 18px; line-height: 1.4;">Redes de Financiación del Terrorismo:<br>Simulación y análisis mediante Economía de Redes y Teoría de Juegos</div></div>""", unsafe_allow_html=True)
    
    st.markdown("<hr style='margin-top: 15px; margin-bottom: 20px; border-top: 1px solid #111;'>", unsafe_allow_html=True)
    if assets_status["Flag FR B64"]: st.markdown(f"""<div class="language-block"><img src="data:image/png;base64,{assets_status['Flag FR B64']}" class="flag-img" /><div class="title-text">Réseaux de Financement du Terrorisme:<br>Simulation et analyse à travers l'Économie des Réseaux et la Théorie des Jeux</div></div>""", unsafe_allow_html=True)
    if assets_status["Flag GB B64"]: st.markdown(f"""<div class="language-block"><img src="data:image/png;base64,{assets_status['Flag GB B64']}" class="flag-img" /><div class="title-text">Terrorism Financing Networks:<br>Simulation and Analysis through Network Economics and Game Theory</div></div>""", unsafe_allow_html=True)
    st.markdown("<hr style='margin-top: 15px; margin-bottom: 25px; border-top: 1px solid #111;'>", unsafe_allow_html=True)

    st.markdown("## 🎓 Sinopsis del Proyecto")
    with st.expander("Leer Sinopsis / Abstract Completo", expanded=True):
        tab_fr, tab_es, tab_en = st.tabs(["Français", "Español", "English"])
        with tab_fr: st.markdown("""<div style="font-size: 15px; text-align: justify; color: #333;">Le Financement du Terrorisme (FT) consiste en la collecte de fonds, ce qui englobe le processus de sollicitation, de rassemblement, de fourniture et de mise à disposition d'argent ou d'actifs dans le but de faciliter la capacité à mener des activités terroristes. En Espagne, la loi 10/2010 établit un cadre rigoureux contre la fourniture ou la distribution de fonds.<br><br>Les grands groupes organisés, les petites cellules et les acteurs individuels ont besoin d'argent pour développer leurs activités terroristes. La littérature académique s'accorde à dire que le manque de fonds limite considérablement leur capacité opérationnelle, le FT étant un élément structurant du terrorisme mondial.<br><br>Ce travail fonde son analyse sur des informations récentes, en utilisant des exemples représentatifs des dynamiques contemporaines. L'objectif principal est de construire une <strong>simulation d'un réseau de financement du terrorisme</strong> basée sur les preuves recueillies dans la littérature spécialisée (typologies du GAFI, ABE, etc.).<br><br>Une analyse structurelle est réalisée sur cette simulation à l'aide de l'Économie des Réseaux et de la Théorie des Jeux, incluant l'étude des métriques de centralité, l'importance des nœuds clés (points d'étranglement) et la résilience du système face aux interventions policières. Le modèle qui en résulte constitue une représentation réaliste et empiriquement fondée, aboutissant à un outil analytique très utile pour le renseignement financier.</div>""", unsafe_allow_html=True)
        with tab_es: st.markdown("""<div style="font-size: 15px; text-align: justify; color: #333;">La Financiación del Terrorismo (FT) consiste en la captación de fondos, lo cual abarca el proceso de solicitud, recaudación, provisión y puesta a disposición de dinero o activos con el fin de facilitar o potenciar la capacidad de cualquier persona u organización para llevar a cabo actividades relacionadas con el terrorismo. En el caso concreto de España, la Ley 10/2010 establece un marco riguroso frente al suministro, depósito o distribución de fondos.<br><br>Tanto los grandes grupos organizados como las pequeñas células e incluso los actores individuales necesitan dinero para el desarrollo de la actividad terrorista. La literatura académica y los informes institucionales coinciden en que la falta de fondos limita drásticamente su capacidad operativa, siendo la FT un elemento vertebrador del terrorismo global.<br><br>Este trabajo fundamenta su análisis en información reciente, utilizando ejemplos observados en los últimos años que resultan representativos de las dinámicas contemporáneas. El objetivo principal consiste en construir una <strong>simulación de red de financiación del terrorismo</strong> basada en la evidencia recogida en la literatura especializada (tipologías del GAFI, EBA, etc.).<br><br>Sobre dicha simulación se realiza un análisis estructural mediante herramientas propias del análisis de redes (Economía de Redes) y la Teoría de Juegos, incluyendo el estudio de métricas de centralidad, la importancia relativa de los nodos clave (chokepoints) y la resiliencia del sistema ante intervenciones policiales. El modelo resultante constituye una representación realista y fundamentada empíricamente, derivando en un modelo analítico altamente útil para la inteligencia financiera y el diseño de políticas de seguridad.</div>""", unsafe_allow_html=True)
        with tab_en: st.markdown("""<div style="font-size: 15px; text-align: justify; color: #333;">Terrorist Financing (TF) involves the raising of funds, which encompasses the process of soliciting, collecting, providing, and making available money or assets to facilitate or enhance the capacity of any individual or organization to carry out terrorist activities. In Spain, Law 10/2010 establishes a rigorous framework against the supply, deposit, or distribution of funds.<br><br>Large organized groups, small cells, and lone actors require money to carry out terrorist activities. Academic literature and institutional reports agree that a lack of funds drastically limits their operational capacity, making TF a structural backbone of global terrorism.<br><br>This paper bases its analysis on recent information, using examples observed in recent years that are representative of contemporary dynamics. The main objective is to build a <strong>simulation of a terrorist financing network</strong> based on evidence gathered from specialized literature (FATF typologies, EBA, etc.).<br><br>A structural analysis is performed on this simulation using Network Economics and Game Theory, including the study of centrality metrics, the relative importance of key nodes (chokepoints), and the system's resilience to law enforcement interventions. The resulting model constitutes a realistic and empirically grounded representation, resulting in an analytical model highly useful for financial intelligence and security policy design.</div>""", unsafe_allow_html=True)

# ==========================================
# 3. FUNCIONES Y CARGA DE DATOS
# ==========================================
def leer_tabla(wb, name):
    for s in wb.worksheets:
        if name in s.tables:
            r = s[s.tables[name].ref]
            return pd.DataFrame([[c.value for c in row] for row in r][1:], columns=[c.value for c in r[0]])
    return pd.DataFrame()

def aplicar_estilos_base(df):
    styler = df.style.set_properties(**{'text-align': 'center'})
    if 'Activo' in df.columns:
        def color_a(v):
            try:
                if int(v)==1: return 'background-color: #C6EFCE; color: #006100;'
                if int(v)==0: return 'background-color: #FFC7CE; color: #9C0006;'
            except: pass
            return ''
        styler = styler.applymap(color_a, subset=['Activo'])
    return styler

st.markdown("---")
st.markdown("<div style='text-align: center; color: #555; font-size: 14px;'>FLUJO DEL SIMULADOR</div>", unsafe_allow_html=True)
archivo_subido = st.file_uploader("Sube tu archivo Excel alternativo", type=["xlsx"])

wb = None
if archivo_subido: wb = load_workbook(archivo_subido, data_only=True)
elif os.path.exists("Diseño Red FT.xlsx"): wb = load_workbook("Diseño Red FT.xlsx", data_only=True)

if wb:
    # Carga de tablas
    df_nodos_orig = leer_tabla(wb, "tblNodos")
    df_enlaces_orig = leer_tabla(wb, "tblEnlaces")
    df_tipos = leer_tabla(wb, "tblTiposDeNodo")
    df_pesos = leer_tabla(wb, "tblPesos")
    df_p_costes = leer_tabla(wb, "tblMatrizPonderadaCostes")
    df_p_valor = leer_tabla(wb, "tblMatrizPonderadaValoroperativo")
    df_tradeoff = leer_tabla(wb, "tblMatrizTradeOff")

    # FASE 1
    st.markdown("## ⚙️ Fase 1: Base de Datos")
    t1, t2, t3, t4 = st.tabs(["🟢 Nodos", "🔗 Enlaces", "🏷️ Tipos", "⚖️ Pesos"])
    with t1: df_n_ed = st.data_editor(aplicar_estilos_base(df_nodos_orig), use_container_width=True, key="ed_n")
    with t2: df_e_ed = st.data_editor(aplicar_estilos_base(df_enlaces_orig), use_container_width=True, key="ed_e")
    with t3: st.dataframe(aplicar_estilos_base(df_tipos), use_container_width=True)
    with t4: st.dataframe(aplicar_estilos_base(df_pesos), use_container_width=True)

    # Procesamiento Grafo
    n_ok = df_n_ed[pd.to_numeric(df_n_ed['Activo'], errors='coerce')==1]
    e_ok = df_e_ed[pd.to_numeric(df_e_ed['Activo'], errors='coerce')==1]
    G = nx.DiGraph()
    for _, r in n_ok.iterrows():
        G.add_node(str(r['NodoID']), label=f"{r['NodoID']} - {r.get('Nombre','')}", color={'O':'#FF9999','I':'#99CCFF','G':'#FFCC99','D':'#99FF99'}.get(str(r['Tipo']),'#CCC'))
    for _, r in e_ok.iterrows():
        if str(r['Nodo Origen']) in G.nodes and str(r['Nodo Destino']) in G.nodes:
            G.add_edge(str(r['Nodo Origen']), str(r['Nodo Destino']))

    # FASE 2
    st.markdown("## 🌐 Fase 2: Topología Interactiva")
    net = Network(height="600px", width="100%", directed=True, bgcolor="#ffffff")
    net.from_nx(G)
    net.save_graph("grafo.html")
    components.html(open("grafo.html",'r').read(), height=620)

    # FASE 3
    st.markdown("## 🔢 Fase 3: Modelado Matemático")
    with st.expander("Matrices Matemáticas del Sistema", expanded=True):
        m1, m2, m3, m4 = st.tabs(["1️⃣ Adyacencia", "2️⃣ Costes", "3️⃣ Valor", "4️⃣ Trade-Off"])
        
        with m1:
            st.markdown("**Matriz Binaria:** Topología estructural (1 = conectado).")
            adj = nx.to_pandas_adjacency(G, dtype=int).replace(0, "")
            st.dataframe(adj, use_container_width=True)
            
        with m2:
            st.markdown("**Matriz de Costes:** Escala térmica en rojos para fricción/exposición.")
            if not df_p_costes.empty:
                df_p_costes.set_index(df_p_costes.columns[0], inplace=True)
                st.dataframe(df_p_costes.style.background_gradient(cmap='OrRd').format(na_rep=""), use_container_width=True)
                
        with m3:
            st.markdown("**Matriz de Valor Operativo:** Escala térmica en azules para capacidad/eficiencia.")
            if not df_p_valor.empty:
                df_p_valor.set_index(df_p_valor.columns[0], inplace=True)
                st.dataframe(df_p_valor.style.background_gradient(cmap='Blues').format(na_rep=""), use_container_width=True)
                
        with m4:
            st.markdown("**Matriz de Trade-Off:** Relación coste-beneficio (2 decimales y escala verde).")
            if not df_tradeoff.empty:
                df_tradeoff.set_index(df_tradeoff.columns[0], inplace=True)
                df_num = df_tradeoff.apply(pd.to_numeric, errors='coerce')
                st.dataframe(df_num.style.background_gradient(cmap='YlGn').format("{:.2f}", na_rep=""), use_container_width=True)
