import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re
import os

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="ERP Financiero - Control de Pagos, Cheques y OCs",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .metric-card {
        background-color: #1e222d;
        border-radius: 10px;
        padding: 15px;
        border-left: 5px solid #00b4d8;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }
    .metric-title {
        color: #8d99ae;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 5px;
    }
    .metric-value {
        color: #ffffff;
        font-size: 1.5rem;
        font-weight: bold;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #111827;
        border-radius: 8px 8px 0px 0px;
        color: #9ca3af;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1f2937 !important;
        color: #38bdf8 !important;
        border-bottom: 3px solid #38bdf8 !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# RUTAS DE ARCHIVOS Y CONEXIÓN A LA RED
# ---------------------------------------------------------
RUTAS_RED = {
    "op": r"H:\Comercial\COMPRAS\8 - INDICADORES\4 - ANÁLISIS DE PAGOS\ORDENES DE PAGO.xlsx",
    "oc": r"H:\Comercial\COMPRAS\5 - SEGUIMIENTO DE COMPRAS\2026. SEGUIMIENTO OC.xlsm",
    "facturas": r"H:\Comercial\COMPRAS\8 - INDICADORES\3 - ANÁLISIS DE COSTOS\FACTURAS - PRESUPUESTO VS COSTO.xlsm",
    "saldos": r"H:\Comercial\COMPRAS\8 - INDICADORES\4 - ANÁLISIS DE PAGOS\SALDOS PENDIENTES Y PAGOS REALIZADOS POR PROYECTOS.xlsm"
}

RUTAS_LOCALES = {
    "op": "ORDENES DE PAGO.xlsx",
    "oc": "2026. SEGUIMIENTO OC.xlsm",
    "facturas": "FACTURAS - PRESUPUESTO VS COSTO.xlsm",
    "saldos": "SALDOS PENDIENTES Y PAGOS REALIZADOS POR PROYECTOS.xlsm"
}

def obtener_path(clave):
    if os.path.exists(RUTAS_RED[clave]):
        return RUTAS_RED[clave]
    return RUTAS_LOCALES[clave]

# ---------------------------------------------------------
# FUNCIÓN DE PROCESAMIENTO AVANZADO DE ORDENES DE COMPRA
# ---------------------------------------------------------
def procesar_ordenes_compra(df_oc):
    # Asegurar que la columna PROYECTO sea de tipo texto para evitar errores con Datetime
    df_oc = df_oc.copy()
    df_oc['PROYECTO'] = df_oc['PROYECTO'].astype(object)
    
    parsed_rows = []
    
    for _, row in df_oc.iterrows():
        valor_orig = row['VALOR OC']
        proyecto_raw = str(row['PROYECTO']).strip() if pd.notnull(row['PROYECTO']) else 'SIN PROYECTO'
        
        # 1. Ajuste de Moneda (<100,000 en USD -> *1515) e IVA (+21%)
        if pd.isna(valor_orig) or valor_orig == 0:
            valor_base_ars = 0.0
        else:
            v = float(valor_orig)
            if v < 100000:
                v = v * 1515.0  # Conversión de USD a ARS
            valor_base_ars = v * 1.21  # Incorporación del 21% de IVA

        # 2. Desglose de Proyectos y Porcentajes (ej. (34) CAGNOLI & (66) ORMAZABAL)
        if '&' in proyecto_raw:
            parts = [p.strip() for p in proyecto_raw.split('&')]
            pct_list = []
            proj_list = []
            has_pct = False
            
            for part in parts:
                match = re.match(r'^\((\d+)\%?\)\s*(.*)', part)
                if match:
                    has_pct = True
                    pct_list.append(float(match.group(1)))
                    proj_list.append(match.group(2).strip())
                else:
                    pct_list.append(1.0)
                    proj_list.append(part)
                    
            if has_pct:
                total_pct = sum(pct_list) if sum(pct_list) > 0 else 100.0
                for proj, pct in zip(proj_list, pct_list):
                    if pct > 0:
                        r = row.to_dict()
                        r['PROYECTO'] = str(proj)
                        r['VALOR OC'] = valor_base_ars * (pct / total_pct)
                        parsed_rows.append(r)
            else:
                n = len(parts)
                for proj in parts:
                    r = row.to_dict()
                    r['PROYECTO'] = str(proj)
                    r['VALOR OC'] = valor_base_ars / n
                    parsed_rows.append(r)
        else:
            match = re.match(r'^\((\d+)\%?\)\s*(.*)', proyecto_raw)
            r = row.to_dict()
            if match:
                r['PROYECTO'] = str(match.group(2).strip())
            else:
                r['PROYECTO'] = str(proyecto_raw)
            r['VALOR OC'] = valor_base_ars
            parsed_rows.append(r)
            
    df_parsed = pd.DataFrame(parsed_rows)
    df_activas = df_parsed[df_parsed['ESTADO'].isin(['PENDIENTE', 'PARCIAL'])].copy()
    df_activas['FECHA ESTIMADA DE PAGO'] = pd.to_datetime(df_activas['FECHA ESTIMADA DE PAGO'], errors='coerce')
    df_activas['PROYECTO'] = df_activas['PROYECTO'].fillna('SIN PROYECTO').astype(str)
    df_activas['PROVEEDOR'] = df_activas['PROVEEDOR'].fillna('PROVEEDOR S/N').astype(str)
    
    return df_activas

# ---------------------------------------------------------
# CARGA DE DATOS GENERAL
# ---------------------------------------------------------
@st.cache_data
def cargar_y_preparar_datos():
    # 1. Facturas
    df_fact = pd.read_excel(obtener_path("facturas"), sheet_name='Facturas')
    df_fact['Comprobante_clean'] = df_fact['Comprobante'].astype(str).str.strip()
    df_fact['Fecha'] = pd.to_datetime(df_fact['Fecha'], errors='coerce')
    df_fact['Centro de Costo'] = df_fact['Centro de Costo'].fillna('SIN CENTRO DE COSTO').astype(str)
    df_fact['Proveedor'] = df_fact['Proveedor'].fillna('PROVEEDOR S/N').astype(str)

    df_fact_map = df_fact[['Comprobante_clean', 'Proveedor', 'Centro de Costo']].drop_duplicates(subset=['Comprobante_clean'])

    # 2. Órdenes de Pago (Facturas Aplicadas y Valores Entregados)
    xls_op = pd.ExcelFile(obtener_path("op"))
    df_op_fact = pd.read_excel(xls_op, sheet_name='Facturas_Aplicadas')
    df_op_val = pd.read_excel(xls_op, sheet_name='Valores_Entregados')

    df_op_fact['Factura_clean'] = df_op_fact['Factura'].astype(str).str.strip()
    df_op_fact_mapped = pd.merge(
        df_op_fact,
        df_fact_map,
        left_on='Factura_clean',
        right_on='Comprobante_clean',
        how='left'
    )
    df_op_fact_mapped['Proveedor'] = df_op_fact_mapped['Proveedor'].fillna('PROVEEDOR S/N').astype(str)
    df_op_fact_mapped['Centro de Costo'] = df_op_fact_mapped['Centro de Costo'].fillna('SIN CENTRO DE COSTO').astype(str)

    df_op_fact_mapped['ImporteFactura_num'] = (
        df_op_fact_mapped['ImporteFactura']
        .astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
    )
    df_op_fact_mapped['ImporteFactura_num'] = pd.to_numeric(df_op_fact_mapped['ImporteFactura_num'], errors='coerce').fillna(0)

    df_op_val['ImporteCheque_num'] = (
        df_op_val['ImporteCheque']
        .astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
    )
    df_op_val['ImporteCheque_num'] = pd.to_numeric(df_op_val['ImporteCheque_num'], errors='coerce').fillna(0)
    df_op_val['FechaCheque_dt'] = pd.to_datetime(df_op_val['FechaCheque'], errors='coerce', dayfirst=True)

    op_set = set(df_op_fact['Factura'].dropna().astype(str).str.strip())
    df_fact['Estado_Pago'] = df_fact['Comprobante_clean'].apply(lambda x: 'PAGADO' if x in op_set else 'PENDIENTE')

    df_cheques_full = pd.merge(df_op_val, df_op_fact_mapped, on='NroOrdenPago', how='left')

    def clasificar_medio(tipo):
        t = str(tipo).upper()
        if 'E-CHEQ' in t or 'ECHEQ' in t:
            return 'E-CHEQ'
        elif 'CHEQUE' in t:
            return 'Cheque Físico'
        elif 'CUENTA' in t or 'TRANSFERENCIA' in t:
            return 'Transferencia / Cta'
        return 'Otros / Efectivo'

    df_cheques_full['Categoria_Pago'] = df_cheques_full['TipoValor'].apply(clasificar_medio)

    hoy = pd.to_datetime(datetime.now().date())
    df_cheques_full['Estado_Vencimiento'] = df_cheques_full['FechaCheque_dt'].apply(
        lambda x: 'Vencido / Pasado' if pd.notnull(x) and x < hoy else 'Próximo Vencimiento'
    )

    # 3. Saldos Pendientes (Hoja1)
    df_saldos = pd.read_excel(obtener_path("saldos"), sheet_name='Hoja1')
    df_saldos['Fecha factura'] = pd.to_datetime(df_saldos['Fecha factura'], errors='coerce')
    df_saldos['Fecha vencimiento'] = pd.to_datetime(df_saldos['Fecha vencimiento'], errors='coerce')
    df_saldos['Centro de costo'] = df_saldos['Centro de costo'].fillna('SIN CENTRO DE COSTO').astype(str)
    df_saldos['Proveedor'] = df_saldos['Proveedor'].fillna('PROVEEDOR S/N').astype(str)

    # 4. Órdenes de Compra Procesadas
    df_oc = pd.read_excel(obtener_path("oc"), sheet_name='ENTREGA PENDIENTE')
    df_oc_activas = procesar_ordenes_compra(df_oc)

    return df_fact, df_saldos, df_oc_activas, df_cheques_full

try:
    df_facturas_raw, df_saldos_raw, df_oc_raw, df_cheques_raw = cargar_y_preparar_datos()
except Exception as e:
    st.error(f"⚠️ Error cargando archivos desde la red o locales: {e}")
    st.stop()

# ---------------------------------------------------------
# FILTROS LATERALES DINÁMICOS
# ---------------------------------------------------------
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2830/2830284.png", width=65)
st.sidebar.title("🎛️ Filtros Globales")
st.sidebar.markdown("---")

proyectos_unicos = sorted(list(set(
    df_saldos_raw['Centro de costo'].astype(str).unique().tolist() +
    df_facturas_raw['Centro de Costo'].astype(str).unique().tolist() +
    df_oc_raw['PROYECTO'].astype(str).unique().tolist() +
    df_cheques_raw['Centro de Costo'].dropna().astype(str).unique().tolist()
)))

proveedores_unicos = sorted(list(set(
    df_saldos_raw['Proveedor'].astype(str).unique().tolist() +
    df_facturas_raw['Proveedor'].astype(str).unique().tolist() +
    df_oc_raw['PROVEEDOR'].astype(str).unique().tolist() +
    df_cheques_raw['Proveedor'].dropna().astype(str).unique().tolist()
)))

proyectos_sel = st.sidebar.multiselect("📂 Proyecto (Centro de Costo):", options=proyectos_unicos, default=[])
proveedores_sel = st.sidebar.multiselect("🏢 Proveedor:", options=proveedores_unicos, default=[])

def filtrar_df(df, col_proj, col_prov):
    df_filtered = df.copy()
    if proyectos_sel and col_proj in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_proj].astype(str).isin(proyectos_sel)]
    if proveedores_sel and col_prov in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_prov].astype(str).isin(proveedores_sel)]
    return df_filtered

df_saldos = filtrar_df(df_saldos_raw, 'Centro de costo', 'Proveedor')
df_facturas = filtrar_df(df_facturas_raw, 'Centro de Costo', 'Proveedor')
df_oc_activas = filtrar_df(df_oc_raw, 'PROYECTO', 'PROVEEDOR')
df_cheques = filtrar_df(df_cheques_raw, 'Centro de Costo', 'Proveedor')

# ---------------------------------------------------------
# ENCABEZADO Y KPIS PRINCIPALES
# ---------------------------------------------------------
st.title("📊 ERP Financiero - Control de Deudas, Cheques & Cash Flow")
st.markdown("##### Tablero ejecutivo interconectado de gestión de compras, facturas y flujo de fondos")

monto_deuda_pendiente = df_saldos['Importe'].sum()
monto_compromiso_oc = df_oc_activas['VALOR OC'].sum()

cheques_proximos_df = df_cheques[df_cheques['Estado_Vencimiento'] == 'Próximo Vencimiento']
echeqs_proximos_monto = cheques_proximos_df[cheques_proximos_df['Categoria_Pago'] == 'E-CHEQ']['ImporteFactura_num'].sum()
total_proximo_vencimiento = cheques_proximos_df['ImporteFactura_num'].sum()

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #ef4444;">
        <div class="metric-title">🔴 DEUDA PENDIENTE (SALDOS)</div>
        <div class="metric-value">AR$ {monto_deuda_pendiente:,.2f}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #f59e0b;">
        <div class="metric-title">🟡 COMPROMISOS OCs ACTIVAS (INC. IVA)</div>
        <div class="metric-value">AR$ {monto_compromiso_oc:,.2f}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #00b4d8;">
        <div class="metric-title">⚡ PRÓXIMOS E-CHEQS A VENCER</div>
        <div class="metric-value">AR$ {echeqs_proximos_monto:,.2f}</div>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #10b981;">
        <div class="metric-title">🗓️ TOTAL VENCIMIENTOS PRÓXIMOS</div>
        <div class="metric-value">AR$ {total_proximo_vencimiento:,.2f}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ---------------------------------------------------------
# PESTAÑAS PRINCIPALES
# ---------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "💳 Calendario E-Cheqs & Valores",
    "🔴 Deuda Pendiente & Vencimientos",
    "🗓️ Cash Flow - Compromisos OCs",
    "📈 Histórico de Facturación"
])

# =========================================================
# TAB 1: CALENDARIO E-CHEQS Y VALORES
# =========================================================
with tab1:
    st.subheader("💳 Monitoreo de E-Cheqs y Valores por Importe Facturado")
    st.markdown("Visualización detallada del dinero comprometido a nivel de factura cancelada, proveedor, proyecto y fecha de débito.")

    col_f1, col_f2 = st.columns([1, 1])
    with col_f1:
        filtro_estado_venc = st.radio(
            "Filtrar Vencimientos:",
            options=["Solo Próximos Vencimientos", "Solo Histórico (Vencidos / Pasados)", "Mostrar Todos"],
            index=0,
            horizontal=True
        )
    with col_f2:
        medios_disponibles = list(df_cheques['Categoria_Pago'].unique()) if len(df_cheques) > 0 else []
        filtro_medios = st.multiselect(
            "Filtrar por Medio de Pago:",
            options=medios_disponibles,
            default=medios_disponibles
        )

    df_ch_view = df_cheques[df_cheques['Categoria_Pago'].isin(filtro_medios)].copy() if len(df_cheques) > 0 else df_cheques.copy()
    
    if filtro_estado_venc == "Solo Próximos Vencimientos":
        df_ch_view = df_ch_view[df_ch_view['Estado_Vencimiento'] == 'Próximo Vencimiento']
    elif filtro_estado_venc == "Solo Histórico (Vencidos / Pasados)":
        df_ch_view = df_ch_view[df_ch_view['Estado_Vencimiento'] == 'Vencido / Pasado']

    if len(df_ch_view) > 0:
        df_ch_view['FechaCheque_Str'] = df_ch_view['FechaCheque_dt'].dt.strftime('%Y-%m-%d')

        st.markdown("### 🗓️ Cronograma de Vencimientos (Base: Importe Facturado)")
        df_timeline = df_ch_view.groupby(['FechaCheque_Str', 'Categoria_Pago'])['ImporteFactura_num'].sum().reset_index()
        
        fig_ch = px.bar(
            df_timeline,
            x='FechaCheque_Str',
            y='ImporteFactura_num',
            color='Categoria_Pago',
            title="Salida de Dinero Requerida por Fecha (Diferenciando E-CHEQ vs Otros Medios)",
            labels={'FechaCheque_Str': 'Fecha de Vencimiento / Pago', 'ImporteFactura_num': 'Importe Facturado Cancelado (AR$)'},
            color_discrete_map={
                'E-CHEQ': '#00b4d8',
                'Cheque Físico': '#8b5cf6',
                'Transferencia / Cta': '#10b981',
                'Otros / Efectivo': '#f59e0b'
            },
            barmode='stack'
        )
        fig_ch.update_layout(hovermode="x unified")
        st.plotly_chart(fig_ch, use_container_width=True)

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            df_ch_view['AñoMes_Venc'] = df_ch_view['FechaCheque_dt'].dt.to_period('M').astype(str)
            df_mensual = df_ch_view.groupby(['AñoMes_Venc', 'Categoria_Pago'])['ImporteFactura_num'].sum().reset_index()
            
            fig_m = px.bar(
                df_mensual,
                x='AñoMes_Venc',
                y='ImporteFactura_num',
                color='Categoria_Pago',
                title="Proyección Mensual de Compromisos de Pago (AR$)",
                color_discrete_map={
                    'E-CHEQ': '#00b4d8',
                    'Cheque Físico': '#8b5cf6',
                    'Transferencia / Cta': '#10b981',
                    'Otros / Efectivo': '#f59e0b'
                },
                barmode='group'
            )
            st.plotly_chart(fig_m, use_container_width=True)

        with col_m2:
            df_pie_medios = df_ch_view.groupby('Categoria_Pago')['ImporteFactura_num'].sum().reset_index()
            fig_p = px.pie(
                df_pie_medios,
                values='ImporteFactura_num',
                names='Categoria_Pago',
                title="Distribución del Importe Facturado por Instrumento",
                hole=0.4,
                color='Categoria_Pago',
                color_discrete_map={
                    'E-CHEQ': '#00b4d8',
                    'Cheque Físico': '#8b5cf6',
                    'Transferencia / Cta': '#10b981',
                    'Otros / Efectivo': '#f59e0b'
                }
            )
            st.plotly_chart(fig_p, use_container_width=True)

        st.markdown("### 🔍 Detalle Unificado: Órdenes de Pago, Facturas Aplicadas y Cheques")
        st.dataframe(
            df_ch_view[[
                'NroOrdenPago', 'Proveedor', 'Centro de Costo', 'Factura',
                'ImporteFactura_num', 'FechaCheque_dt', 'Categoria_Pago',
                'TipoValor', 'NroCheque', 'Estado_Vencimiento'
            ]].rename(columns={
                'ImporteFactura_num': 'Importe Factura (AR$)',
                'FechaCheque_dt': 'Fecha Vencimiento/Cheque',
                'Centro de Costo': 'Proyecto / Centro de Costo'
            }),
            use_container_width=True
        )
    else:
        st.info("No hay registros de valores/cheques para los filtros seleccionados.")

# =========================================================
# TAB 2: DEUDA PENDIENTE
# =========================================================
with tab2:
    st.subheader("🔴 Análisis de Deuda Acreedora Actual (Saldos Pendientes)")
    
    col_a, col_b = st.columns([1, 1])
    
    with col_a:
        df_prov_deuda = df_saldos.groupby('Proveedor')['Importe'].sum().reset_index()
        df_prov_deuda = df_prov_deuda.sort_values(by='Importe', ascending=False).head(10)
        
        fig_prov = px.bar(
            df_prov_deuda,
            x='Importe',
            y='Proveedor',
            orientation='h',
            title="Top 10 Proveedores con Mayor Deuda Pendiente (AR$)",
            color='Importe',
            color_continuous_scale='Reds'
        )
        fig_prov.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
        st.plotly_chart(fig_prov, use_container_width=True)

    with col_b:
        df_proj_deuda = df_saldos.groupby('Centro de costo')['Importe'].sum().reset_index()
        df_proj_deuda = df_proj_deuda.sort_values(by='Importe', ascending=False).head(10)
        
        fig_proj = px.bar(
            df_proj_deuda,
            x='Importe',
            y='Centro de costo',
            orientation='h',
            title="Deuda Pendiente por Proyecto / Centro de Costo (AR$)",
            color='Importe',
            color_continuous_scale='Oranges'
        )
        fig_proj.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
        st.plotly_chart(fig_proj, use_container_width=True)

    st.markdown("### 🗓️ Curva de Vencimientos de Deuda Pendiente")
    df_saldos['FechaVenc_Str'] = df_saldos['Fecha vencimiento'].dt.strftime('%Y-%m-%d')
    df_venc = df_saldos.groupby('FechaVenc_Str')['Importe'].sum().reset_index()
    
    fig_venc = px.bar(
        df_venc,
        x='FechaVenc_Str',
        y='Importe',
        title="Vencimientos de Deuda Pendiente por Fecha (AR$)",
        labels={'FechaVenc_Str': 'Fecha de Vencimiento', 'Importe': 'Monto Deuda'},
        color_discrete_sequence=['#ef4444']
    )
    st.plotly_chart(fig_venc, use_container_width=True)

    with st.expander("🔍 Ver Detalle Completo de Facturas en Deuda Pendiente"):
        st.dataframe(
            df_saldos[['Fecha factura', 'Fecha vencimiento', 'Comprobante', 'Proveedor', 'Centro de costo', 'Importe', 'Mora']],
            use_container_width=True
        )

# =========================================================
# TAB 3: COMPROMISOS OC (CASH FLOW)
# =========================================================
with tab3:
    st.subheader("🗓️ Calendario de Compromisos de Pago por Órdenes de Compra (PENDIENTE / PARCIAL)")
    st.caption("Nota: Los valores de las OCs incluyen el 21% de IVA y la conversión de moneda cuando corresponde.")
    
    if len(df_oc_activas) > 0:
        df_oc_activas['FechaPago_Str'] = df_oc_activas['FECHA ESTIMADA DE PAGO'].dt.strftime('%Y-%m-%d')
        df_oc_cf = df_oc_activas.groupby(['FechaPago_Str', 'PROYECTO'])['VALOR OC'].sum().reset_index()
        
        fig_cf = px.bar(
            df_oc_cf,
            x='FechaPago_Str',
            y='VALOR OC',
            color='PROYECTO',
            title="Requerimiento de Fondos para OCs por Fecha Estimada de Pago (AR$)",
            labels={'FechaPago_Str': 'Fecha Estimada de Pago', 'VALOR OC': 'Monto Requerido (AR$ con IVA)'},
            barmode='stack'
        )
        st.plotly_chart(fig_cf, use_container_width=True)
        
        col_oc1, col_oc2 = st.columns(2)
        with col_oc1:
            df_oc_prov = df_oc_activas.groupby('PROVEEDOR')['VALOR OC'].sum().reset_index().sort_values(by='VALOR OC', ascending=False).head(10)
            fig_oc_p = px.pie(df_oc_prov, values='VALOR OC', names='PROVEEDOR', title="Compromiso OCs por Proveedor", hole=0.4)
            st.plotly_chart(fig_oc_p, use_container_width=True)
            
        with col_oc2:
            df_oc_proj = df_oc_activas.groupby('PROYECTO')['VALOR OC'].sum().reset_index().sort_values(by='VALOR OC', ascending=False).head(10)
            fig_oc_pj = px.pie(df_oc_proj, values='VALOR OC', names='PROYECTO', title="Compromiso OCs por Proyecto", hole=0.4)
            st.plotly_chart(fig_oc_pj, use_container_width=True)

        st.markdown("### 📋 Listado Detallado de Órdenes de Compra en Ejecución")
        st.dataframe(
            df_oc_activas[['Nº DE OC ', 'PROVEEDOR', 'PROYECTO', 'VALOR OC', 'FECHA ESTIMADA DE PAGO', 'ESTADO', 'DESCRIPCIÓN']].rename(
                columns={'VALOR OC': 'VALOR OC (AR$ c/IVA)'}
            ),
            use_container_width=True
        )
    else:
        st.info("No hay Órdenes de Compra en ejecución para los filtros seleccionados.")

# =========================================================
# TAB 4: HISTÓRICO DE FACTURACIÓN
# =========================================================
with tab4:
    st.subheader("📈 Evolución Histórica de Facturación y Pagos Realizados")
    
    df_facturas['AñoMes'] = df_facturas['Fecha'].dt.to_period('M').astype(str)
    df_evo = df_facturas.groupby(['AñoMes', 'Estado_Pago'])['Importe Facturado'].sum().reset_index()
    
    fig_evo = px.bar(
        df_evo,
        x='AñoMes',
        y='Importe Facturado',
        color='Estado_Pago',
        title="Facturación Mensual Histórica: Pagas vs. Pendientes",
        color_discrete_map={'PAGADO': '#10b981', 'PENDIENTE': '#ef4444'},
        barmode='stack'
    )
    st.plotly_chart(fig_evo, use_container_width=True)

    with st.expander("📋 Ver Registro General de Facturas"):
        st.dataframe(
            df_facturas[['Proveedor', 'Fecha', 'Comprobante', 'Centro de Costo', 'Importe Facturado', 'Estado_Pago']],
            use_container_width=True
        )