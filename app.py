import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import requests
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Dashboard SEGPRO - Quejas y Reclamos",
    page_icon="📊",
    layout="wide"
)

# CSS personalizado
st.markdown("""
<style>
    .stMetric {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# Función para convertir link de SharePoint
def convertir_link_sharepoint(url):
    try:
        # Quitar parámetros y agregar download=1
        base_url = url.split('?')[0]
        return base_url + '?download=1'
    except:
        return url

# Función para cargar datos
@st.cache_data(ttl=60)
def cargar_datos_excel(archivo):
    try:
        df = pd.read_excel(archivo, engine='openpyxl')
        return df, None
    except Exception as e:
        return None, str(e)

# Función para cargar desde URL
@st.cache_data(ttl=60)
def cargar_desde_url(url):
    try:
        url = convertir_link_sharepoint(url)
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
        response.raise_for_status()
        
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        return df, None
        
    except Exception as e:
        return None, f"Error: {str(e)}"

# Función para clasificar tipo de error según SEGPRO
def clasificar_tipo_error(texto):
    if pd.isna(texto):
        return "Sin clasificar"
    
    texto = str(texto).lower()
    
    # Mapeo exacto de tipos de error SEGPRO
    if 'defectuoso' in texto or 'defecto' in texto or 'falla' in texto:
        return "Producto defectuoso"
    elif 'talla' in texto or 'tamaño' in texto or 'número' in texto:
        return "Error de talla"
    elif 'faltante' in texto or 'falta' in texto or 'incompleto' in texto or 'pieza' in texto:
        return "Pieza faltante"
    elif 'color' in texto:
        return "Color incorrecto"
    elif 'no coincide' in texto or 'equivocado' in texto or 'otro producto' in texto or 'incorrecto' in texto:
        return "Producto no coincide con lo solicitado"
    elif 'transporte' in texto or 'dañado' in texto or 'empaque' in texto or 'golpeado' in texto:
        return "Daño por transporte"
    elif 'fábrica' in texto:
        return "Producto con fallas de fábrica"
    else:
        return "Otros"

# Función para identificar productos SEGPRO
def identificar_producto(texto):
    if pd.isna(texto):
        return "Sin especificar"
    
    texto = str(texto).lower()
    
    # Productos específicos SEGPRO
    if 'multi' in texto or 'flex' in texto or 'guante' in texto:
        return "Guante Multi Flex"
    elif 'harder' in texto or 'zapato' in texto:
        return "Zapatos Harder"
    elif 'cono' in texto or 'naranja' in texto:
        return "Cono Naranja"
    else:
        return texto.title()[:50]  # Limitar longitud

# Función para clasificar estado
def clasificar_estado(respuesta):
    if pd.isna(respuesta) or str(respuesta).strip() == '':
        return "Pendiente"
    
    texto = str(respuesta).lower()
    
    if any(p in texto for p in ['resuelto', 'solucionado', 'cerrado', 'completado', 'atendido']):
        return "Resuelto"
    elif any(p in texto for p in ['proceso', 'gestionando', 'revisando', 'evaluando']):
        return "En Proceso"
    else:
        return "Pendiente"

# Función para calcular satisfacción
def calcular_satisfaccion(texto, es_inicial=True):
    if pd.isna(texto):
        return 2 if es_inicial else None
    
    texto = str(texto).lower()
    
    if any(p in texto for p in ['pésimo', 'horrible', 'terrible']):
        return 1
    elif any(p in texto for p in ['malo', 'insatisfecho', 'molesto']):
        return 2
    elif any(p in texto for p in ['regular', 'aceptable']):
        return 3
    elif any(p in texto for p in ['bueno', 'satisfecho', 'bien']):
        return 4 if not es_inicial else 3
    elif any(p in texto for p in ['excelente', 'perfecto', 'genial']):
        return 5 if not es_inicial else 4
    else:
        return 2 if es_inicial else 3

# Título principal
col1, col2 = st.columns([3, 1])
with col1:
    st.title("📊 Dashboard de Quejas y Reclamos")
    st.markdown("**SEGPRO** - Análisis en Tiempo Real")
with col2:
    try:
        st.image("https://wardia.com.pe/segpro/wp-content/uploads/2024/08/logo-segpro-300x163.png", width=150)
    except:
        pass

st.markdown("---")

# Sidebar
st.sidebar.header("⚙️ Configuración")

opcion_carga = st.sidebar.radio(
    "Método de carga:",
    ["URL de SharePoint/OneDrive", "Subir archivo Excel", "Datos de ejemplo"]
)

df = None

if opcion_carga == "URL de SharePoint/OneDrive":
    st.sidebar.info("💡 Pega el link de SharePoint")
    url = st.sidebar.text_input(
        "URL del archivo:",
        value="https://usilpe-my.sharepoint.com/:x:/g/personal/andres_gallardo_usil_pe/ESwYnIB2V95Do5-WZWL1GEIBsaF17Mc-gJZWQjgHSFuOGQ?e=HrAUrS"
    )
    
    if st.sidebar.button("🔄 Cargar desde URL"):
        with st.spinner("🌐 Cargando..."):
            df, error = cargar_desde_url(url)
            if df is not None:
                st.sidebar.success(f"✅ {len(df)} registros")
            else:
                st.sidebar.error(f"❌ {error}")

elif opcion_carga == "Subir archivo Excel":
    archivo = st.sidebar.file_uploader("Cargar archivo", type=['xlsx', 'xls'])
    if archivo:
        df, error = cargar_datos_excel(archivo)
        if df is not None:
            st.sidebar.success(f"✅ {len(df)} registros")

else:  # Datos de ejemplo
    df = pd.DataFrame({
        'fecha': pd.date_range(end=datetime.now(), periods=20, freq='D'),
        'cliente': ['Cliente ' + str(i) for i in range(20)],
        'producto': ['Guante Multi Flex', 'Zapatos Harder', 'Cono Naranja'] * 7 + ['Guante Multi Flex'],
        'tipo_error': ['Producto defectuoso', 'Error de talla', 'Color incorrecto', 'Daño por transporte'] * 5,
        'descripcion': ['Queja ejemplo'] * 20,
        'respuesta': ['Solucionado'] * 15 + [None] * 5
    })
    st.sidebar.info("📝 Datos de ejemplo")

# Procesar datos
if df is not None and len(df) > 0:
    
    # Mostrar columnas detectadas
    with st.expander("🔍 Vista previa y columnas detectadas"):
        st.write(f"**Total de registros:** {len(df)}")
        st.write(f"**Columnas encontradas:** {list(df.columns)}")
        st.dataframe(df.head(10), use_container_width=True)
    
    # Normalizar nombres de columnas
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    cols = list(df.columns)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Mapeo de Columnas")
    
    # Detectar automáticamente columnas relevantes
    col_fecha = None
    col_cliente = None
    col_producto = None
    col_tipo_error = None
    col_descripcion = None
    col_respuesta = None
    
    for col in cols:
        if 'fecha' in col or 'date' in col:
            col_fecha = col
        elif 'cliente' in col or 'nombre' in col:
            col_cliente = col
        elif 'producto' in col or 'item' in col:
            col_producto = col
        elif 'tipo' in col and 'error' in col:
            col_tipo_error = col
        elif 'descripcion' in col or 'detalle' in col or 'queja' in col or 'reclamo' in col or 'asunto' in col:
            col_descripcion = col
        elif 'respuesta' in col or 'solucion' in col or 'resolucion' in col:
            col_respuesta = col
    
    # Permitir selección manual
    col_fecha_sel = st.sidebar.selectbox("📅 Fecha:", cols, index=cols.index(col_fecha) if col_fecha else 0)
    col_producto_sel = st.sidebar.selectbox("📦 Producto:", ['Auto-detectar'] + cols, index=0)
    col_tipo_error_sel = st.sidebar.selectbox("⚠️ Tipo de Error:", ['Auto-detectar'] + cols, index=0)
    col_descripcion_sel = st.sidebar.selectbox("📝 Descripción/Queja:", cols, index=cols.index(col_descripcion) if col_descripcion else 0)
    col_respuesta_sel = st.sidebar.selectbox("✅ Respuesta:", ['Sin respuesta'] + cols, index=0)
    
    # Crear DataFrame procesado
    df_proc = pd.DataFrame()
    
    # Fecha
    df_proc['fecha'] = pd.to_datetime(df[col_fecha_sel], errors='coerce')
    
    # Cliente (si existe)
    if col_cliente:
        df_proc['cliente'] = df[col_cliente].fillna('Sin especificar')
    else:
        df_proc['cliente'] = 'Cliente ' + df.index.astype(str)
    
    # Descripción/Queja
    df_proc['descripcion'] = df[col_descripcion_sel].fillna('Sin descripción').astype(str)
    
    # Producto - Auto-detectar o usar columna
    if col_producto_sel == 'Auto-detectar':
        df_proc['producto'] = df_proc['descripcion'].apply(identificar_producto)
    else:
        df_proc['producto'] = df[col_producto_sel].fillna('Sin especificar').apply(identificar_producto)
    
    # Tipo de Error - Auto-detectar o usar columna
    if col_tipo_error_sel == 'Auto-detectar':
        df_proc['tipo_error'] = df_proc['descripcion'].apply(clasificar_tipo_error)
    else:
        df_proc['tipo_error'] = df[col_tipo_error_sel].fillna('').apply(clasificar_tipo_error)
    
    # Respuesta
    if col_respuesta_sel != 'Sin respuesta':
        df_proc['respuesta'] = df[col_respuesta_sel].fillna('')
    else:
        df_proc['respuesta'] = ''
    
    # Estado
    df_proc['estado'] = df_proc['respuesta'].apply(clasificar_estado)
    
    # Satisfacción
    df_proc['satisfaccion_inicial'] = df_proc['descripcion'].apply(lambda x: calcular_satisfaccion(x, True))
    df_proc['satisfaccion_final'] = df_proc.apply(
        lambda row: calcular_satisfaccion(row['respuesta'], False) if row['estado'] == 'Resuelto' else None,
        axis=1
    )
    
    # Días de resolución (simulado)
    df_proc['dias_resolucion'] = df_proc.apply(
        lambda row: ((datetime.now() - row['fecha']).days if pd.notna(row['fecha']) else 5) 
        if row['estado'] == 'Resuelto' else None,
        axis=1
    )
    
    # Filtros
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros")
    
    # Filtro de fechas
    df_con_fecha = df_proc[df_proc['fecha'].notna()]
    
    if len(df_con_fecha) > 0:
        fecha_min = df_con_fecha['fecha'].min()
        fecha_max = df_con_fecha['fecha'].max()
        
        fecha_inicio = st.sidebar.date_input("Desde:", fecha_min.date())
        fecha_fin = st.sidebar.date_input("Hasta:", fecha_max.date())
        
        mask_fecha = (df_proc['fecha'].dt.date >= fecha_inicio) & (df_proc['fecha'].dt.date <= fecha_fin)
    else:
        mask_fecha = [True] * len(df_proc)
        st.sidebar.warning("⚠️ No hay fechas válidas")
    
    # Filtro por tipo de error
    tipos_error_disponibles = sorted(df_proc['tipo_error'].unique())
    tipos_error_filtro = st.sidebar.multiselect(
        "Tipo de Error:",
        options=tipos_error_disponibles,
        default=tipos_error_disponibles
    )
    
    # Filtro por estado
    estados_disponibles = sorted(df_proc['estado'].unique())
    estados_filtro = st.sidebar.multiselect(
        "Estado:",
        options=estados_disponibles,
        default=estados_disponibles
    )
    
    # Aplicar filtros
    mask_tipo = df_proc['tipo_error'].isin(tipos_error_filtro)
    mask_estado = df_proc['estado'].isin(estados_filtro)
    df_filtrado = df_proc[mask_fecha & mask_tipo & mask_estado]
    
    # Métricas principales
    st.header("📈 Métricas Principales")
    
    col1, col2, col3, col4 = st.columns(4)
    
    total = len(df_filtrado)
    resueltos = len(df_filtrado[df_filtrado['estado'] == 'Resuelto'])
    tasa = (resueltos / total * 100) if total > 0 else 0
    
    tiempo_prom = df_filtrado[df_filtrado['dias_resolucion'].notna()]['dias_resolucion'].mean()
    tiempo_prom = tiempo_prom if pd.notna(tiempo_prom) else 0
    
    sat_inicial = df_filtrado['satisfaccion_inicial'].mean()
    sat_final = df_filtrado[df_filtrado['satisfaccion_final'].notna()]['satisfaccion_final'].mean()
    mejora = sat_final - sat_inicial if pd.notna(sat_final) else 0
    
    with col1:
        st.metric("Total Quejas", total)
    with col2:
        st.metric("Tasa Resolución", f"{tasa:.1f}%", delta=f"{resueltos} resueltas")
    with col3:
        st.metric("Tiempo Promedio", f"{tiempo_prom:.1f} días")
    with col4:
        st.metric("Mejora Satisfacción", f"+{mejora:.1f}", delta=f"{sat_inicial:.1f} → {sat_final:.1f}")
    
    st.markdown("---")
    
    # Gráficos
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Por Tipo de Error")
        tipo_counts = df_filtrado['tipo_error'].value_counts()
        st.bar_chart(tipo_counts)
    
    with col2:
        st.subheader("📈 Por Estado")
        estado_counts = df_filtrado['estado'].value_counts()
        st.bar_chart(estado_counts)
    
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("🏆 Top 5 Productos")
        producto_counts = df_filtrado['producto'].value_counts().head(5)
        st.bar_chart(producto_counts)
    
    with col4:
        st.subheader("⭐ Satisfacción")
        sat_data = pd.DataFrame({
            'Inicial': [sat_inicial],
            'Final': [sat_final if pd.notna(sat_final) else 0]
        })
        st.bar_chart(sat_data)
    
    # Tendencias
    st.subheader("📅 Tendencias Temporales")
    df_temporal = df_filtrado[df_filtrado['fecha'].notna()].groupby(df_filtrado['fecha'].dt.date).size()
    if len(df_temporal) > 0:
        st.line_chart(df_temporal)
    else:
        st.info("No hay datos con fechas válidas para mostrar tendencias")
    
    # Tabla
    st.markdown("---")
    st.subheader("📋 Detalle de Quejas")
    
    df_mostrar = df_filtrado[['fecha', 'cliente', 'producto', 'tipo_error', 'estado', 
                               'satisfaccion_inicial', 'satisfaccion_final']].copy()
    df_mostrar['fecha'] = df_mostrar['fecha'].apply(
        lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else 'Sin fecha'
    )
    
    st.dataframe(df_mostrar, use_container_width=True, height=400)
    
    # Exportar
    st.markdown("---")
    csv = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        "📥 Descargar Datos Procesados (CSV)",
        data=csv,
        file_name=f"quejas_segpro_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

else:
    st.info("👈 Selecciona un método de carga en el menú lateral")
    
    st.markdown("""
    ### 📝 Instrucciones:
    
    **Productos detectados automáticamente:**
    - 🧤 Guante Multi Flex
    - 👞 Zapatos Harder
    - 🚧 Cono Naranja
    
    **Tipos de error clasificados:**
    - Producto defectuoso
    - Error de talla
    - Pieza faltante
    - Color incorrecto
    - Producto no coincide con lo solicitado
    - Daño por transporte
    - Producto con fallas de fábrica
    
    El dashboard detectará automáticamente estos valores en tu Excel.
    """)

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Dashboard SEGPRO © 2024</div>", unsafe_allow_html=True)
