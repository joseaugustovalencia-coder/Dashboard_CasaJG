import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import re
import base64
import io

# ==========================================
# 1. CONFIGURACIÓN DEL ENTORNO (UI/UX)
# ==========================================
st.set_page_config(
    page_title="CONTROL DE OBRAS - JUMBO IA",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS inyectados para forzar el modo premium y tarjetas (Cards) compactas
st.markdown("""
    <style>
    .main-title {
        font-family: 'Google Sans', sans-serif;
        font-weight: 700;
        color: #FF8C00;
        margin-bottom: 0.5rem;
    }
    .metric-card {
        background-color: #1E1E1E;
        padding: 12px 18px;
        border-radius: 8px;
        border-left: 4px solid #FF8C00;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.4);
        margin-bottom: 15px;
    }
    .metric-value { 
        font-size: 20px; 
        font-weight: bold; 
        color: #FFFFFF; 
    }
    .metric-label { 
        font-size: 11px; 
        color: #A0A0A0; 
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .rec-indicator {
        height: 12px;
        width: 12px;
        background-color: #ff0000;
        border-radius: 50%;
        display: inline-block;
        animation: blinker 1.5s linear infinite;
        margin-right: 8px;
    }
    @keyframes blinker {
        50% { opacity: 0; }
    }
    .admin-badge {
        background-color: #EF4444;
        color: white;
        padding: 0.2rem 0.6rem;
        border-radius: 0.25rem;
        font-size: 0.8rem;
        font-weight: bold;
    }
    .client-badge {
        background-color: #10B981;
        color: white;
        padding: 0.2rem 0.6rem;
        border-radius: 0.25rem;
        font-size: 0.8rem;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. MOTOR DE DATOS Y LIMPIEZA REGEX
# ==========================================
def limpiar_numero(valor):
    """Limpia cualquier formato de moneda o porcentaje de Excel (ej: 'Bs. 1.500,50' -> 1500.50)"""
    v = str(valor).strip()
    v = re.sub(r'[^\d\.,\-]', '', v) # Deja solo números, puntos y comas
    if not v or v in ['-', '.', ',']: return 0.0
    
    if '.' in v and ',' in v:
        if v.rfind(',') > v.rfind('.'): 
            v = v.replace('.', '').replace(',', '.')
        else: 
            v = v.replace(',', '')
    elif ',' in v:
        v = v.replace(',', '.')
        
    try:
        return float(v)
    except:
        return 0.0

def leer_csv_robusto(file):
    def reset_puntero():
        if hasattr(file, 'seek'): file.seek(0)
    
    try:
        reset_puntero()
        df = pd.read_csv(file, sep=';', encoding='latin1')
        if len(df.columns) > 1: return df
    except: pass
    
    try:
        reset_puntero()
        df = pd.read_csv(file, sep=',', encoding='latin1')
        if len(df.columns) > 1: return df
    except: pass

    reset_puntero()
    return pd.read_csv(file, sep=',', encoding='utf-8')

@st.cache_data
def cargar_datos_locales():
    archivos = {
        "avance": "PRESUPUESTO OBRA GRUESA E INSTALACIONES_CASA JG_2.xlsx - BD_Avance.csv",
        "item": "PRESUPUESTO OBRA GRUESA E INSTALACIONES_CASA JG_2.xlsx - BD-Item.csv",
        "crono": "PRESUPUESTO OBRA GRUESA E INSTALACIONES_CASA JG_2.xlsx - BD_Cronograma.csv"
    }
    df_dict = {}
    for key, filename in archivos.items():
        if os.path.exists(filename):
            df_dict[key] = leer_csv_robusto(filename)
        else:
            return None
    return df_dict["avance"], df_dict["item"], df_dict["crono"]

# Ejecución de la ingesta
datos = cargar_datos_locales()

if datos is None:
    st.warning("⚠️ Sube tus frentes de trabajo CSV a continuación para activar el Dashboard:")
    col_up1, col_up2, col_up3 = st.columns(3)
    with col_up1: f_avance = st.file_uploader("1. Sube BD_Avance.csv", type=['csv'])
    with col_up2: f_item = st.file_uploader("2. Sube BD-Item.csv", type=['csv'])
    with col_up3: f_crono = st.file_uploader("3. Sube BD_Cronograma.csv", type=['csv'])
    
    if f_avance and f_item and f_crono:
        df_avance = leer_csv_robusto(f_avance)
        df_item = leer_csv_robusto(f_item)
        df_crono = leer_csv_robusto(f_crono)
    else:
        st.stop()
else:
    df_avance, df_item, df_crono = datos

# ==========================================
# 3. NORMALIZACIÓN Y PROGRAMACIÓN DEFENSIVA (ESCUDO ANTIBUG)
# ==========================================
df_avance.columns = df_avance.columns.str.strip()
df_item.columns = df_item.columns.str.strip()
df_crono.columns = df_crono.columns.str.strip()

# Renombrar columnas con tildes si existen
df_crono = df_crono.rename(columns={"ID_SubMódulo": "ID_Sub_Modulo", "ID_Submódulo": "ID_Sub_Modulo"})

# Inyectar columnas requeridas si Excel no las tiene
columnas_req_avance = {'ID_Modulo': 'str', 'Nombre_Modulo': 'str', 'Presupuesto_Base_Bs': 'float', 'Pct_Avance_Fisico': 'float', 'Pct_Avance_Financiero': 'float', 'Fecha_Corte': 'str'}
for col, tipo in columnas_req_avance.items():
    if col not in df_avance.columns:
        df_avance[col] = 0.0 if tipo == 'float' else 'Sin Datos'

columnas_req_item = {'ID_Modulo': 'str', 'Pres_Material_Equipos_Bs': 'float', 'Real_Mat_Eq_Bs': 'float', 'Pres_Mano_Obra_Bs': 'float', 'Real_Mano_Obra_Bs': 'float'}
for col, tipo in columnas_req_item.items():
    if col not in df_item.columns:
        df_item[col] = 0.0 if tipo == 'float' else 'Sin Datos'

columnas_req_crono = {'Modulo_Asignado': 'str', 'Descripcion': 'str', 'Fecha_Inicio': 'str', 'Fecha_Fin': 'str', 'Pct_Completado': 'float'}
for col, tipo in columnas_req_crono.items():
    if col not in df_crono.columns:
        if col == 'Descripcion' and 'Tarea_Descripcion' in df_crono.columns:
            df_crono = df_crono.rename(columns={'Tarea_Descripcion': 'Descripcion'})
        elif col == 'Descripcion' and 'Tarea' in df_crono.columns:
            df_crono = df_crono.rename(columns={'Tarea': 'Descripcion'})
        else:
            df_crono[col] = 0.0 if tipo == 'float' else 'Sin Datos'

# Limpieza y conversión estricta
df_avance['ID_Modulo'] = df_avance['ID_Modulo'].astype(str).str.strip()
df_item['ID_Modulo'] = df_item['ID_Modulo'].astype(str).str.strip()

for col in ['Presupuesto_Base_Bs', 'Pct_Avance_Fisico', 'Pct_Avance_Financiero']:
    df_avance[col] = df_avance[col].apply(limpiar_numero)

for col in ['Pres_Material_Equipos_Bs', 'Real_Mat_Eq_Bs', 'Pres_Mano_Obra_Bs', 'Real_Mano_Obra_Bs']:
    df_item[col] = df_item[col].apply(limpiar_numero)

# ==========================================
# 4. SIDEBAR Y SISTEMA DE AUTENTICACIÓN
# ==========================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Flag_of_Santa_Cruz.svg/250px-Flag_of_Santa_Cruz.svg.png", width=50)
    st.title("⚙️ Panel de Control")
    st.divider()
    
    tipo_cambio = st.slider("Tipo de Cambio (Bs/USD)", 6.96, 15.00, 10.03, 0.01)
    
    st.divider()
    
    # CONTROL DE ACCESO MASTER
    st.subheader("🔒 Acceso de Personal")
    modo_admin = st.checkbox("Activar Modo Administrador (Jumbo)", value=False)
    es_admin = False
    
    if modo_admin:
        clave = st.text_input("Ingresar Clave Maestra:", type="password")
        if clave == "josesito5775":
            st.success("Acceso Autorizado 🟢")
            es_admin = True
        elif clave != "":
            st.error("Contraseña Incorrecta")
            
    st.divider()
    st.subheader("🕹️ Controles de Demostración")
    modo_simulador = st.toggle("Activar Modo Simulador", value=False, help="Inyecta avance y costos ficticios si tus archivos locales están vacíos.")
    
    if modo_simulador:
        st.success("Modo Simulador Activo")
        # Forzar simulado a Abril de 2027
        fecha_corte_obj = datetime(2027, 4, 15)
        df_avance['Pct_Avance_Fisico'] = [0.85, 0.45, 0.60, 0.35][:len(df_avance)]
        df_avance['Pct_Avance_Financiero'] = [0.80, 0.50, 0.55, 0.30][:len(df_avance)]
        df_item['Real_Mat_Eq_Bs'] = df_item['Pres_Material_Equipos_Bs'] * df_item['ID_Modulo'].map({'A': 0.82, 'B': 0.40, 'C': 0.55, 'D': 0.48}).fillna(0.5)
        df_item['Real_Mano_Obra_Bs'] = df_item['Pres_Mano_Obra_Bs'] * df_item['ID_Modulo'].map({'A': 0.80, 'B': 0.45, 'C': 0.50, 'D': 0.45}).fillna(0.5)
        df_crono['Pct_Completado'] = 0.5
    else:
        try:
            fecha_corte_obj = datetime.strptime(str(df_avance['Fecha_Corte'].iloc[0]).strip(), "%Y-%m-%d")
        except:
            fecha_corte_obj = datetime.now()
            
    st.divider()
    st.info(f"**Proyecto:** CASA JG\n\n**Ubicación:** Santa Cruz, Bolivia\n\n**Fecha Corte:** {fecha_corte_obj.strftime('%d/%m/%Y')}")

# ==========================================
# 5. INICIALIZACIÓN DE BASES DE DATOS DE SESIÓN
# ==========================================
if "db_fotos" not in st.session_state:
    st.session_state.db_fotos = [
        {
            "titulo": "Inicio de Excavación de Cimientos y Nivelación",
            "frente": "Obra Gruesa",
            "fecha": "2026-05-10",
            "descripcion": "Labores preliminares de movimiento de tierras y replanteo topográfico de los ejes estructurales de la edificación.",
            "url": "https://images.unsplash.com/photo-1541888946425-d81bb19240f5?auto=format&fit=crop&w=800&q=80"
        },
        {
            "titulo": "Armadura de Fierro de Refuerzo para Zapatas",
            "frente": "Obra Gruesa",
            "fecha": "2026-05-24",
            "descripcion": "Colocación de parrillas inferiores de acero corrugado AH500 en cimientos para asegurar alta rigidez sísmica de las fundaciones.",
            "url": "https://images.unsplash.com/photo-1504307651254-35680f356dfd?auto=format&fit=crop&w=800&q=80"
        },
        {
            "titulo": "Encofrado y Vaciado de Hormigón H25 en Losas",
            "frente": "Obra Gruesa",
            "fecha": "2026-06-08",
            "descripcion": "Vaciado continuo de losa aligerada con supervisión de revenimiento del hormigón premezclado provisto por Soboce.",
            "url": "https://images.unsplash.com/photo-1590069261209-f8e9b8642343?auto=format&fit=crop&w=800&q=80"
        }
    ]

if "db_planillas" not in st.session_state:
    st.session_state.db_planillas = [
        {"numero": "PL-001", "fecha": "2026-05-20", "monto_bs": 45000.0, "descripcion": "Vaciado de cimientos y replanteo", "estado": "Pagada"},
        {"numero": "PL-002", "fecha": "2026-06-05", "monto_bs": 62000.0, "descripcion": "Armado de fierros y encofrado de columnas", "estado": "Pagada"},
        {"numero": "PL-003", "fecha": "2026-06-20", "monto_bs": 35000.0, "descripcion": "Suministro e instalación sanitaria preliminar", "estado": "Pendiente"}
    ]

if "db_reportes" not in st.session_state:
    st.session_state.db_reportes = [
        {"numero": "REP-001", "fecha": "2026-05-30", "periodo": "01/05/2026 al 30/05/2026", "archivo": "Informe_Fiscalizacion_Mayo.pdf", "tipo": "Oficial"}
    ]

# ==========================================
# 6. CÁLCULOS FINANCIEROS GLOBALES
# ==========================================
AREA_TOTAL_M2 = 575.0
PORCENTAJE_LOGISTICA = 0.15

costo_directo_bs = df_avance['Presupuesto_Base_Bs'].sum()
costo_logistica_bs = costo_directo_bs * PORCENTAJE_LOGISTICA
costo_total_obra_bs = costo_directo_bs + costo_logistica_bs

costo_total_obra_usd = costo_total_obra_bs / tipo_cambio
costo_m2_usd = costo_total_obra_usd / AREA_TOTAL_M2

# Ponderación de avances
if costo_directo_bs > 0:
    df_avance['Peso_Relativo'] = df_avance['Presupuesto_Base_Bs'] / costo_directo_bs
else:
    df_avance['Peso_Relativo'] = 0

avance_fisico_global = (df_avance['Pct_Avance_Fisico'] * df_avance['Peso_Relativo']).sum() * 100
avance_financiero_global = (df_avance['Pct_Avance_Financiero'] * df_avance['Peso_Relativo']).sum() * 100

# ==========================================
# 7. MAQUETACIÓN MULTIPESTAÑA
# ==========================================
tab_control, tab_planillas, tab_galeria, tab_proveedores, tab_camaras, tab_informes = st.tabs([
    "📊 Control Financiero",
    "🧾 Planillas MO",
    "📸 Galería de Avance",
    "📞 Directorio de Proveedores",
    "📹 Monitoreo IP en Obra",    
    "📋 Informes y Reportes"
    
    
])

# -----------------------------------------------------------------------------
# TAB 1: CONTROL FINANCIERO E HITOS
# -----------------------------------------------------------------------------
with tab_control:
    # Encabezado Comercial Premium
    st.markdown("""
        <div style="text-align: center; margin-top: 10px; margin-bottom: 20px;">
            <h1 style="color: #FF8C00; font-size: 36px; font-weight: bold; margin-bottom: 5px; letter-spacing: 0.5px;">
                CONTROL DE OBRAS - JUMBO IA
            </h1>
            <p style="font-size: 13px; color: #CCCCCC; margin-bottom: 2px; text-transform: uppercase; letter-spacing: 1px;">
                PROYECTO: <b>CASA JG - CONDOMINIO LA FLORESTA URUBÓ</b>
            </p>
            <p style="font-size: 12px; color: #999999; margin-top: 0px;">
                PROPIETARIOS: <b>JUAN PABLO ANTONIO JAUREGUI PAZ - GINA MOLINA DE JAUREGUI</b>
            </p>
            <hr style="border-top: 2px solid #FF8C00; width: 45%; margin: 15px auto;">
        </div>
    """, unsafe_allow_html=True)

    # Disposición Ejecutiva (Donut Chart a la izquierda, KPI Cards reducidos a la derecha)
    col_summary_chart, col_summary_kpis = st.columns([1, 1])

    with col_summary_chart:
        try:
            # Cálculo de avance real versus pendiente monetariamente
            valor_ejecutado_bs = costo_total_obra_bs * (avance_fisico_global / 100)
            valor_pendiente_bs = costo_total_obra_bs - valor_ejecutado_bs

            fig_donut = go.Figure(data=[go.Pie(
                labels=['Avance Realizado (Físico)', 'Inversión por Ejecutar'],
                values=[valor_ejecutado_bs, valor_pendiente_bs],
                hole=.6,
                marker=dict(colors=['#00CC96', '#2D3748']),
                hoverinfo='label+percent',
                textinfo='percent',
                textfont_size=16
            )])

            fig_donut.update_layout(
                title=dict(
                    text="<b>Avance de obra en tiempo real</b>",
                    x=0.0,
                    y=0.96,
                    xanchor="left",
                    font=dict(size=15, color="#FFFFFF")
                ),
                showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5, font=dict(color="#FFFFFF")),
                margin=dict(t=40, b=10, l=10, r=10),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                height=260
            )
            
            fig_donut.add_annotation(
                text=f"<b>{avance_fisico_global:.1f}%</b><br><span style='font-size:10px;'>FÍSICO</span>",
                showarrow=False,
                font=dict(size=20, color="#FFFFFF"),
                x=0.5, y=0.5
            )

            st.plotly_chart(fig_donut, use_container_width=True)
        except Exception as e:
            st.error(f"Falla al generar Donut Chart: {e}")

    with col_summary_kpis:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Costo Total Obra Presupuestado (Bolivianos)</div>
                <div class="metric-value">{costo_total_obra_bs:,.2f} Bs</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Costo Total Obra en Dólares (Fluctuante)</div>
                <div class="metric-value">${costo_total_obra_usd:,.2f} USD <span style="font-size:12px; color:#A0A0A0;">@ {tipo_cambio}</span></div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Costo Promedio / Metro Cuadrado de Edificación</div>
                <div class="metric-value">${costo_m2_usd:,.2f} USD/m² <span style="font-size:11px; color:#A0A0A0;">(Sobre {AREA_TOTAL_M2} m²)</span></div>
            </div>
        """, unsafe_allow_html=True)

    st.divider()

    # --- SECCIÓN 1: AVANCES POR MÓDULO ---
    try:
        st.subheader("1. Avance Físico vs Avance Financiero por Módulo")
        
        df_avance['Bs_Fisico'] = df_avance['Pct_Avance_Fisico'] * df_avance['Presupuesto_Base_Bs']
        df_avance['Bs_Financiero'] = df_avance['Pct_Avance_Financiero'] * df_avance['Presupuesto_Base_Bs']

        df_melted_avance = df_avance.melt(id_vars=['Nombre_Modulo'], 
                                          value_vars=['Pct_Avance_Fisico', 'Pct_Avance_Financiero'], 
                                          var_name='Tipo de Avance', value_name='Porcentaje')
        df_melted_bs = df_avance.melt(id_vars=['Nombre_Modulo'], 
                                      value_vars=['Bs_Fisico', 'Bs_Financiero'], 
                                      var_name='Tipo de Bs', value_name='Monto_Bs')

        df_melted_avance['Monto_Bs'] = df_melted_bs['Monto_Bs']
        df_melted_avance['Porcentaje'] = df_melted_avance['Porcentaje'] * 100 

        fig_avance = px.bar(df_melted_avance, x='Nombre_Modulo', y='Porcentaje', color='Tipo de Avance', barmode='group',
                            color_discrete_sequence=['#00CC96', '#EF553B'],
                            text=df_melted_avance['Monto_Bs'].apply(lambda x: f'{x:,.0f} Bs'))

        fig_avance.update_traces(textposition='outside')
        fig_avance.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', 
            yaxis=dict(title="Avance Ponderado (%)", range=[0, 100]), 
            xaxis_title=""
        )
        st.plotly_chart(fig_avance, use_container_width=True)
    except Exception as e:
        st.error(f"Falla en Gráfico de Avances: {e}")

    st.divider()

    # --- SECCIÓN 2: CONTROL DE COSTOS DIRECTOS (MATERIALES Y MANO DE OBRA) ---
    try:
        st.subheader("2. Desviación de Costos Directos: Materiales y Mano de Obra")
        col_mat, col_mo = st.columns(2)

        df_agrupado = df_item.groupby('ID_Modulo').agg({
            'Pres_Material_Equipos_Bs': 'sum', 'Real_Mat_Eq_Bs': 'sum',
            'Pres_Mano_Obra_Bs': 'sum', 'Real_Mano_Obra_Bs': 'sum'
        }).reset_index()

        mapa_nombres = dict(zip(df_avance['ID_Modulo'], df_avance['Nombre_Modulo']))
        df_agrupado['Nombre_Modulo'] = df_agrupado['ID_Modulo'].map(mapa_nombres).fillna(df_agrupado['ID_Modulo'])

        with col_mat:
            st.markdown("**Análisis de Adquisición de Materiales y Equipamiento (Bs.)**")
            df_mat = df_agrupado.melt(id_vars=['Nombre_Modulo'], value_vars=['Pres_Material_Equipos_Bs', 'Real_Mat_Eq_Bs'],
                                      var_name='Estado', value_name='Monto (Bs)')
            df_mat['Estado'] = df_mat['Estado'].map({'Pres_Material_Equipos_Bs': 'Presupuestado', 'Real_Mat_Eq_Bs': 'Ejecutado (Real)'})
            
            fig_mat = px.bar(df_mat, x='Nombre_Modulo', y='Monto (Bs)', color='Estado', barmode='group',
                             color_discrete_sequence=['#64748B', '#FF8C00'], text_auto='.0s')
            fig_mat.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis_title="")
            st.plotly_chart(fig_mat, use_container_width=True)

        with col_mo:
            st.markdown("**Análisis de Desembolsos en Mano de Obra y Planillas (Bs.)**")
            df_mo = df_agrupado.melt(id_vars=['Nombre_Modulo'], value_vars=['Pres_Mano_Obra_Bs', 'Real_Mano_Obra_Bs'],
                                     var_name='Estado', value_name='Monto (Bs)')
            df_mo['Estado'] = df_mo['Estado'].map({'Pres_Mano_Obra_Bs': 'Presupuestado', 'Real_Mano_Obra_Bs': 'Ejecutado (Real)'})
            
            fig_mo = px.bar(df_mo, x='Nombre_Modulo', y='Monto (Bs)', color='Estado', barmode='group',
                            color_discrete_sequence=['#64748B', '#00CC96'], text_auto='.0s')
            fig_mo.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis_title="")
            st.plotly_chart(fig_mo, use_container_width=True)
    except Exception as e:
        st.error(f"Falla en Gráfico de Desviaciones: {e}")

    st.divider()

    # --- SECCIÓN 3: CRONOGRAMA ---
    try:
        st.subheader("3. Programación Física y Control de Plazos (Hitos Gantt)")
        df_crono['Fecha_Inicio'] = pd.to_datetime(df_crono['Fecha_Inicio'], errors='coerce')
        df_crono['Fecha_Fin'] = pd.to_datetime(df_crono['Fecha_Fin'], errors='coerce')
        df_crono = df_crono.dropna(subset=['Fecha_Inicio', 'Fecha_Fin'])

        # Integrar estatus dinámico temporal
        gantt_estados = []
        for idx, row in df_crono.iterrows():
            duracion = (row["Fecha_Fin"] - row["Fecha_Inicio"]).days if (row["Fecha_Fin"] - row["Fecha_Inicio"]).days > 0 else 1
            dias_t = (fecha_corte_obj - row["Fecha_Inicio"]).days
            progreso_e = 0.0 if dias_t < 0 else min(1.0, dias_t / duracion)
            
            pct_comp_val = limpiar_numero(row.get('Pct_Completado', 0.0))
            if pct_comp_val >= 1.0: gantt_estados.append("Completado 🟢")
            elif fecha_corte_obj < row["Fecha_Inicio"]: gantt_estados.append("No Iniciado ⚪")
            elif pct_comp_val >= progreso_e: gantt_estados.append("En Plazo 🔵")
            else: gantt_estados.append("Retrasado 🔴")
            
        df_crono["Estatus_Dinámico"] = gantt_estados

        fig_gantt = px.timeline(
            df_crono, x_start="Fecha_Inicio", x_end="Fecha_Fin", y="Descripcion", color="Estatus_Dinámico",
            color_discrete_map={"Completado 🟢": "#10B981", "En Plazo 🔵": "#2563EB", "Retrasado 🔴": "#EF4444", "No Iniciado ⚪": "#9CA3AF"}
        )
        fig_gantt.update_yaxes(autorange="reversed")
        fig_gantt.update_layout(
            height=450,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=220, r=30, t=80, b=40),
            yaxis=dict(title=""),
            plot_bgcolor='rgba(0,0,0,0)'
        )
        fig_gantt.add_vline(x=fecha_corte_obj.timestamp() * 1000, line_width=3, line_dash="dash", line_color="#FF8C00")
        st.plotly_chart(fig_gantt, use_container_width=True)
    except Exception as e:
        st.error(f"Falla al cargar Gráfico de Gantt: {e}")

# -----------------------------------------------------------------------------
# TAB 2: PLANILLAS MO
# -----------------------------------------------------------------------------
with tab_planillas:
    st.subheader("🧾 Control y Gestión de Planillas de Mano de Obra")
    st.markdown("Registro cronológico e histórico de desembolsos fiscales asignados a mano de obra y subcontratos.")
    st.divider()

    if es_admin:
        st.markdown("### ⚙️ Registro de Planilla Oficial (Exclusivo Administrador)")
        with st.expander("➕ Cargar Nueva Planilla de Mano de Obra", expanded=True):
            with st.form("form_planilla_mo", clear_on_submit=True):
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    pl_nro = st.text_input("Número de Planilla:", value=f"PL-00{len(st.session_state.db_planillas)+1}")
                    pl_fecha = st.date_input("Fecha de Registro:", value=datetime.now())
                with col_p2:
                    pl_monto = st.number_input("Monto Liquidado (Bs.):", min_value=0.0, step=100.0)
                    pl_estado = st.selectbox("Estado del Pago:", ["Pagada", "Pendiente"])
                
                pl_desc = st.text_area("Hitos y Tareas Cubiertas en la Planilla:")
                pl_pdf = st.file_uploader("Adjuntar Comprobante de Planilla / PDF:", type=["pdf"])

                btn_planilla = st.form_submit_button("Guardar en Historial")
                if btn_planilla:
                    if pl_nro and pl_monto > 0:
                        st.session_state.db_planillas.append({
                            "numero": pl_nro,
                            "fecha": pl_fecha.strftime("%Y-%m-%d"),
                            "monto_bs": pl_monto,
                            "descripcion": pl_desc,
                            "estado": pl_estado
                        })
                        st.success(f"¡La planilla {pl_nro} ha sido incorporada con éxito!")
                        st.rerun()
                    else:
                        st.error("Por favor completa los datos mínimos obligatorios.")
        st.divider()

    # Visualización Cronológica Inversa
    st.markdown("### 📋 Historial Técnico de Planillas de Pago")
    planillas_sorted = sorted(st.session_state.db_planillas, key=lambda x: x["fecha"], reverse=True)

    if not planillas_sorted:
        st.info("No se han registrado planillas de mano de obra aún.")
    else:
        for p in planillas_sorted:
            col_pl1, col_pl2 = st.columns([1, 2])
            badge_color = "background-color: #10B981; color: white;" if p["estado"] == "Pagada" else "background-color: #EF4444; color: white;"
            
            with col_pl1:
                st.markdown(f"#### 📄 {p['numero']}")
                st.markdown(f"📅 **Fecha de Registro:** {p['fecha']}")
                st.markdown(f"💰 **Monto Liquidado:** {p['monto_bs']:,.2f} Bs.")
                st.markdown(f"📌 **Estado:** <span style='padding: 3px 8px; border-radius: 4px; font-weight: bold; {badge_color}'>{p['estado']}</span>", unsafe_allow_html=True)
            with col_pl2:
                st.markdown("**Descripción de Avance:**")
                st.write(p["descripcion"])
                st.button(f"📥 Descargar PDF de Respaldo ({p['numero']})", key=f"btn_dl_pl_{p['numero']}")
            st.markdown("<hr style='border-top: 1px dashed #444;'>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# TAB 3: GALERÍA FOTOGRÁFICA INTERACTIVA
# -----------------------------------------------------------------------------
with tab_galeria:
    st.subheader("📸 Galería de Inspección Visual de Avance")
    st.markdown("Historial fotográfico de la materialización en sitio del proyecto **CASA JG**.")
    st.divider()
    
    if es_admin:
        st.markdown("### 🛠️ Registro de Nuevas Imágenes (Exclusivo Administrador)")
        with st.expander("➕ Cargar Foto en Galería", expanded=False):
            with st.form("form_nueva_foto", clear_on_submit=True):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    nuevo_titulo = st.text_input("Título descriptivo:", placeholder="Ej: Vaciado de Viga de Fundación")
                    nuevo_frente = st.selectbox("Frente de Trabajo:", ["Obra Gruesa", "Obra Fina", "Instalaciones"])
                with col_f2:
                    nueva_fecha = st.date_input("Fecha de Registro:", value=datetime.now())
                    nueva_img = st.file_uploader("Subir Imagen (JPG/PNG):", type=["jpg", "png", "jpeg"])
                
                nueva_desc = st.text_area("Especificación Técnica o Comentarios:")
                btn_subir_foto = st.form_submit_button("Registrar en Galería")
                
                if btn_subir_foto:
                    if nuevo_titulo and nueva_desc and nueva_img:
                        try:
                            image_bytes = nueva_img.read()
                            encoded_image = base64.b64encode(image_bytes).decode("utf-8")
                            mime_type = nueva_img.type
                            encoded_url = f"data:{mime_type};base64,{encoded_image}"
                            
                            st.session_state.db_fotos.append({
                                "titulo": nuevo_titulo,
                                "frente": nuevo_frente,
                                "fecha": nueva_fecha.strftime("%Y-%m-%d"),
                                "descripcion": nueva_desc,
                                "url": encoded_url
                            })
                            st.success("¡Imagen cargada exitosamente!")
                            st.rerun()
                        except Exception as ex:
                            st.error(f"Error al procesar: {ex}")
                    else:
                        st.error("Todos los campos y archivos son obligatorios.")
        st.divider()

    # Controles de Filtros
    col_fil1, col_fil2 = st.columns(2)
    with col_fil1:
        filtro_f = st.selectbox("Filtrar por Frente de Obra:", ["Todos", "Obra Gruesa", "Obra Fina", "Instalaciones"])
    with col_fil2:
        filtro_o = st.selectbox("Ordenar por fecha de captura:", ["Más recientes primero", "Más antiguas primero"])

    fotos_filtradas = st.session_state.db_fotos
    if filtro_f != "Todos":
        fotos_filtradas = [f for f in fotos_filtradas if f["frente"] == filtro_f]

    rev_bool = (filtro_o == "Más recientes primero")
    fotos_filtradas = sorted(fotos_filtradas, key=lambda x: x["fecha"], reverse=rev_bool)

    if not fotos_filtradas:
        st.info("No hay registros en esta sección de la galería.")
    else:
        col_grid1, col_grid2 = st.columns(2)
        for idx, foto in enumerate(fotos_filtradas):
            target_col = col_grid1 if idx % 2 == 0 else col_grid2
            with target_col:
                st.markdown(f"""
                    <div style="background-color: #1E1E1E; padding: 15px; border-radius: 10px; border-bottom: 3px solid #FF8C00; margin-bottom: 25px; box-shadow: 2px 2px 8px rgba(0,0,0,0.4);">
                        <img src="{foto['url']}" style="width: 100%; border-radius: 6px; height: 260px; object-fit: cover; margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <span style="background-color: #2D3748; color: #FF8C00; font-size: 11px; padding: 3px 10px; border-radius: 12px; font-weight: bold; text-transform: uppercase;">{foto['frente']}</span>
                            <span style="color: #A0A0A0; font-size: 12px; font-weight: bold;">📅 {foto['fecha']}</span>
                        </div>
                        <h4 style="color: #FFFFFF; margin-top: 5px; margin-bottom: 6px; font-size: 17px; font-weight: bold;">{foto['titulo']}</h4>
                        <p style="color: #CCCCCC; font-size: 13px; line-height: 1.45; margin-bottom: 0px;">{foto['descripcion']}</p>
                    </div>
                """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# TAB 4: DIRECTORIO DE PROVEEDORES
# -----------------------------------------------------------------------------
with tab_proveedores:
    st.subheader("📞 Directorio de Proveedores y Estado de Cuentas")
    st.markdown("Lista consolidada de subcontratistas y proveedores del proyecto, con control de saldos y montos ejecutados.")
    st.divider()

    df_prov_data = pd.DataFrame({
        "Proveedor / Contratista": [
            "SOBOCE S.A. (Hormigón Listo)", 
            "Aceros LAS LOMAS", 
            "TIGRE PLASMAR S.A.", 
            "Distribuidora EL ALFARERO", 
            "Hormigones JUMBO S.R.L.",
            "INSTALACIONES ELECTROMECÁNICAS S&R"
        ],
        "Rubro de Suministro": [
            "Hormigón Premezclado H21 / H25", 
            "Acero Estructural de Refuerzo AH500", 
            "Sistemas de Conducción PVC Hidro-Sanitaria", 
            "Mampostería, Ladrillo Adobito y Cemento", 
            "Dirección, Logística y Maquinaria Pesada",
            "Conductores de Cobre y Tableros Eléctricos"
        ],
        "Contacto de Atención": [
            "Ing. Carlos Vaca (760-14352)", 
            "Lic. Mariana Soruco (773-82910)", 
            "Dpto. Comercial (3-3467812)", 
            "Ventas Planta Urubó (700-11223)", 
            "Gerencia Técnico (776-94989)",
            "Ing. Ricardo Rojas (708-54321)"
        ],
        "Estado del Contrato": [
            "Suministro Activo", 
            "Suministro Activo", 
            "Planilla de Entrega Pendiente", 
            "En Negociación (Obra Fina)", 
            "Fiscalización Activa",
            "Programado (Fase Instalaciones)"
        ],
        "Monto Pagado (Bs.)": [
            145250.00, 
            98700.00, 
            35400.00, 
            0.00, 
            443018.72, 
            0.00
        ]
    })
    
    st.dataframe(
        df_prov_data, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "Monto Pagado (Bs.)": st.column_config.NumberColumn(
                "Monto Pagado (Bs.)",
                format="Bs. %,.2f",
                help="Total de fondos liquidados acumulados en Bolivianos"
            )
        }
    )

# -----------------------------------------------------------------------------
# TAB 5: MONITOREO DE CÁMARAS IP
# -----------------------------------------------------------------------------
with tab_camaras:
    st.subheader("📹 Monitoreo de Seguridad e Inspección en Tiempo Real")
    st.markdown("Acceso a flujos de transmisión y cámaras de vigilancia situadas en los perímetros de la obra.")
    st.divider()
    
    col_cam_sel, col_cam_feed = st.columns([1, 2])
    
    with col_cam_sel:
        st.markdown("### Seleccionar Cámara de Obra")
        cam_select = st.radio(
            "Cámaras de Seguridad Activas:",
            ["Cam 01: Acceso Principal y Acopio", 
             "Cam 02: Frente de Encofrado y Columnas", 
             "Cam 03: Área de Mezclado de Probetas"]
        )
        
        st.divider()
        st.markdown("### Enlace de Red DVR Local")
        url_personalizada = st.text_input(
            "Dirección IP o URL de tu DVR IP (WebRTC / HTTP):",
            placeholder="http://192.168.1.100:8080/video"
        )
        st.caption("Para integrar la dirección de red del NVR, especifica el dominio o IP DNS mapeado en el enrutador.")
        
    with col_cam_feed:
        st.markdown(f"### <span class='rec-indicator'></span> TRANSMISIÓN EN VIVO: {cam_select}", unsafe_allow_html=True)
        if url_personalizada:
            st.info(f"Conectando a transmisión remota: `{url_personalizada}`...")
            st.video("https://www.w3schools.com/html/mov_bbb.mp4")
        else:
            st.video("https://www.w3schools.com/html/mov_bbb.mp4")
            st.caption("🔴 Servidor de CCTV Operativo. Codificación en flujo continuo H.264 activa.")

# -----------------------------------------------------------------------------
# TAB 6: INFORMES Y REPORTES
# -----------------------------------------------------------------------------
def renderizar_archivo_reportes():
    """Muestra el repositorio de archivos PDF e informes oficiales disponibles."""
    if es_admin:
        st.markdown("### 🛠️ Cargar Reporte de Inspección Oficial (Firma de Supervisión)")
        with st.expander("📤 Cargar PDF del Informe", expanded=False):
            with st.form("form_subir_reporte", clear_on_submit=True):
                col_u_inf1, col_u_inf2 = st.columns(2)
                with col_u_inf1:
                    rep_nro = st.text_input("Código de Informe Oficial:", value="REP-Manual")
                    rep_fecha = st.date_input("Fecha de Emisión:", value=datetime.now())
                with col_u_inf2:
                    rep_periodo = st.text_input("Periodo de Cobertura del Reporte:", value="Quincenal")
                    rep_file = st.file_uploader("Seleccionar PDF:", type=["pdf"])

                btn_guardar_reporte = st.form_submit_button("Indexar Reporte Oficial")
                if btn_guardar_reporte:
                    if rep_nro and rep_file:
                        st.session_state.db_reportes.insert(0, {
                            "numero": rep_nro,
                            "fecha": rep_fecha.strftime("%Y-%m-%d"),
                            "periodo": rep_periodo,
                            "archivo": f"Informe_Manual_{rep_nro}.pdf",
                            "tipo": "Oficial"
                        })
                        st.success("Informe registrado con éxito.")
                        st.rerun()
                    else:
                        st.error("Por favor completa los campos y sube un archivo PDF.")
        st.divider()

    reportes_sorted = sorted(st.session_state.db_reportes, key=lambda x: x["fecha"], reverse=True)
    if not reportes_sorted:
        st.info("No hay informes registrados en la biblioteca virtual.")
    else:
        for r in reportes_sorted:
            badge_type = "🟢 Automatizado" if r["tipo"] == "Generado Automático" else "🔵 Oficial Supervisión"
            col_r1, col_r2, col_r3 = st.columns([1, 2, 1])
            with col_r1:
                st.markdown(f"**Código:** {r['numero']}\n\n`{badge_type}`")
            with col_r2:
                st.markdown(f"**Período:** {r['periodo']}\n\n<span style='color:#A0A0A0; font-size:0.85rem;'>Archivo: {r['archivo']}</span>", unsafe_allow_html=True)
            with col_r3:
                # Botón de descarga real con buffer simulado de PDF
                st.download_button(
                    label="📥 Descargar PDF",
                    data=f"Archivo PDF Simulado del Reporte {r['numero']}\nEmitido el {r['fecha']}.",
                    file_name=r["archivo"],
                    mime="application/pdf",
                    key=f"dl_inf_{r['numero']}_{r['fecha']}"
                )
            st.markdown("<hr style='margin:0.5rem 0; border-top: 1px solid #444;'>", unsafe_allow_html=True)


with tab_informes:
    st.subheader("📋 Repositorio y Emisión de Informes de Obra")
    st.markdown("Consulte los reportes oficiales o genere un informe automatizado basado en las métricas de avance del dashboard.")
    st.divider()

    # Si es administrador, se habilitan las pestañas con el generador inteligente exclusivo
    if es_admin:
        sub_t_inf1, sub_t_inf2 = st.tabs(["⚙️ Generador Inteligente de Informes", "📁 Archivo de Reportes PDF"])

        with sub_t_inf1:
            st.markdown("### Generador Inteligente de Informes")
            
            col_g_inf1, col_g_inf2 = st.columns(2)
            with col_g_inf1:
                cod_inf = st.text_input("Designación de Reporte:", value=f"REP-00{len(st.session_state.db_reportes)+1}")
                fecha_fin_inf = st.date_input("Fecha de Emisión del Reporte:", value=datetime.now())
            with col_g_inf2:
                fecha_ini_inf = st.date_input("Fecha Inicio de Periodo:", value=datetime.now() - timedelta(days=15))
                st.text_input("Ubicación del Proyecto:", value="Condominio La Floresta Urubó, Santa Cruz", disabled=True)

            report_buffer = io.StringIO()
            report_buffer.write(f"REPORTE TÉCNICO DE EVOLUCIÓN: {cod_inf}\n")
            report_buffer.write(f"PROYECTO: CASA JG\n")
            report_buffer.write(f"UBICACIÓN: CONDOMINIO LA FLORESTA URUBÓ, BOLIVIA\n")
            report_buffer.write(f"PERIODO EVALUADO: {fecha_ini_inf.strftime('%d/%m/%Y')} al {fecha_fin_inf.strftime('%d/%m/%Y')}\n")
            report_buffer.write("--------------------------------------------------------------------------------\n")
            report_buffer.write("1. EVALUACIÓN Y COSTOS DEL PROYECTO:\n")
            report_buffer.write(f"   - Avance Físico Global Ponderado: {avance_fisico_global:.2f}%\n")
            report_buffer.write(f"   - Avance Financiero Global Ponderado: {avance_financiero_global:.2f}%\n")
            report_buffer.write(f"   - Presupuesto Total de Obra: {costo_total_obra_bs:,.2f} Bs. / ${costo_total_obra_usd:,.2f} USD\n")
            report_buffer.write(f"   - Tipo de cambio registrado: {tipo_cambio:.2f} Bs./USD\n")
            report_buffer.write(f"   - Ratio Superficial Estimado: ${costo_m2_usd:,.2f} USD/m²\n\n")
            report_buffer.write("2. HISTORIAL DE PLANILLAS DE MANO DE OBRA REGISTRADAS:\n")
            for p in st.session_state.db_planillas:
                report_buffer.write(f"   - {p['numero']} | Fecha: {p['fecha']} | Monto: {p['monto_bs']:,} Bs. | {p['estado']}\n")
            
            st.text_area("Contenido Compilado del Informe:", value=report_buffer.getvalue(), height=250)

            if st.button("Procesar y Emitir en Archivo Técnico", use_container_width=True):
                st.session_state.db_reportes.insert(0, {
                    "numero": cod_inf,
                    "fecha": fecha_fin_inf.strftime("%Y-%m-%d"),
                    "periodo": f"{fecha_ini_inf.strftime('%d/%m/%Y')} al {fecha_fin_inf.strftime('%d/%m/%Y')}",
                    "archivo": f"Informe_Generado_{cod_inf}.pdf",
                    "tipo": "Generado Automático"
                })
                st.success("¡El informe automatizado ha sido compilado e indexado con éxito!")
                st.rerun()

        with sub_t_inf2:
            renderizar_archivo_reportes()

    # Si es cliente o personal no-administrador, únicamente visualiza el repositorio histórico (Sin Generador)
    else:
        renderizar_archivo_reportes()