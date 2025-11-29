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
    if pd.isna(texto) or str(texto).strip() == '':
        return "Sin clasificar"
    
    texto = str(texto).lower()
    
    # Mapeo exacto de tipos de error SEGPRO
    if 'defectuoso' in texto or 'defecto' in texto:
        return "Producto defectuoso"
    elif 'talla' in texto:
        return "Error de talla"
    elif 'faltante' in texto or 'incompleto' in texto or 'pieza' in texto:
        return "Pieza faltante"
    elif 'color' in texto:
        return "Color incorrecto"
    elif 'no coincide' in texto or 'equivocado' in texto:
        return "Producto no coincide con lo solicitado"
    elif 'transporte' in texto or 'dañado' in texto or 'empaque' in texto:
        return "Daño por transporte"
    elif 'fábrica' in texto or 'falla' in texto:
        return "Producto con fallas de fábrica"
    else:
        return "Otros"

# Función para normalizar nombre de productos
def normalizar_producto(texto):
    if pd.isna(texto) or str(texto).strip() == '':
        return "Sin especificar"
    
    texto = str(texto).strip()
    texto_lower = texto.lower()
    
    # Productos específicos SEGPRO
    if 'multi' in texto_lower and 'flex' in texto_lower:
        return "Guante Multi Flex"
    elif 'harder' in texto_lower:
        return "Zapatos Harder"
    elif 'cono' in texto_lower and 'naranja' in texto_lower:
        return "Cono Naranja"
    else:
        # Retornar el nombre original limpio
        return texto.title()

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
    ["URL de SharePoint/OneDrive", "Subir archivo Excel"]
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
                st.sidebar.success(f"✅ {len(df)} registros cargados")
            else:
                st.sidebar.error(f"❌ {error}")

elif opcion_carga == "Subir archivo Excel":
    archivo = st.sidebar.file_uploader("Cargar archivo", type=['xlsx', 'xls'])
    if archivo:
        df, error = cargar_datos_excel(archivo)
        if df is not None:
            st.sidebar.success(f"✅ {len(df)} registros cargados")

# Procesar datos
if df is not None and len(df) > 0:
    
    # Mostrar columnas detectadas
    with st.expander("🔍 Vista previa del Excel cargado"):
        st.write(f"**Total de registros:** {len(df)}")
        st.write(f"**Columnas encontradas:** {list(df.columns)}")
        st.dataframe(df.head(10), use_container_width=True)
    
    # Normalizar nombres de columnas
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    cols = list(df.columns)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Seleccionar Columnas")
    
    # Detectar automáticamente columnas
    col_fecha = next((c for c in cols if 'fecha' in c or 'date' in c), cols[0] if cols else None)
    col_email = next((c for c in cols if 'email' in c or 'correo' in c or 'mail' in c), None)
    col_cliente = next((c for c in cols if 'cliente' in c or 'nombre' in c), None)
    col_producto = next((c for c in cols if 'producto' in c or 'item' in c or 'artículo' in c or 'articulo' in c), None)
    col_tipo_error = next((c for c in cols if ('tipo' in c and 'error' in c) or 'categoria' in c), None)
    col_descripcion = next((c for c in cols if 'descripcion' in c or 'detalle' in c or 'queja' in c or 'reclamo' in c or 'asunto' in c or 'mensaje' in c), None)
    col_respuesta = next((c for c in cols if 'respuesta' in c or 'solucion' in c), None)
    
    # Selectores manuales
    col_fecha_sel = st.sidebar.selectbox("📅 Fecha:", cols, index=cols.index(col_fecha) if col_fecha else 0)
    
    col_email_sel = st.sidebar.selectbox("📧 Email/Correo:", ['No tiene'] + cols, 
                                          index=cols.index(col_email)+1 if col_email else 0)
    
    col_cliente_sel = st.sidebar.selectbox("👤 Cliente/Nombre:", ['No tiene'] + cols,
                                            index=cols.index(col_cliente)+1 if col_cliente else 0)
    
    col_producto_sel = st.sidebar.selectbox("📦 Producto:", ['No tiene'] + cols,
                                             index=cols.index(col_producto)+1 if col_producto else 0)
    
    col_tipo_error_sel = st.sidebar.selectbox("⚠️ Tipo de Error:", ['Auto-clasificar'] + cols,
                                               index=cols.index(col_tipo_error)+1 if col_tipo_error else 0)
    
    col_descripcion_sel = st.sidebar.selectbox("📝 Descripción:", ['No tiene'] + cols,
                                                index=cols.index(col_descripcion)+1 if col_descripcion else 0)
    
    col_respuesta_sel = st.sidebar.selectbox("✅ Respuesta:", ['No tiene'] + cols,
                                              index=cols.index(col_respuesta)+1 if col_respuesta else 0)
    
    # Crear DataFrame procesado
    df_proc = pd.DataFrame()
    
    # Fecha
    df_proc['fecha'] = pd.to_datetime(df[col_fecha_sel], errors='coerce')
    
    # Email
    if col_email_sel != 'No tiene':
        df_proc['email'] = df[col_email_sel].fillna('Sin email').astype(str)
    else:
        df_proc['email'] = 'No disponible'
    
    # Cliente
    if col_cliente_sel != 'No tiene':
        df_proc['cliente'] = df[col_cliente_sel].fillna('Sin nombre').astype(str)
    else:
        df_proc['cliente'] = 'No disponible'
    
    # Producto
    if col_producto_sel != 'No tiene':
        df_proc['producto'] = df[col_producto_sel].fillna('Sin especificar').apply(normalizar_producto)
    else:
        df_proc['producto'] = 'No especificado'
    
    # Descripción
    if col_descripcion_sel != 'No tiene':
        df_proc['descripcion'] = df[col_descripcion_sel].fillna('Sin descripción').astype(str)
    else:
        df_proc['descripcion'] = 'No disponible'
    
    # Tipo de Error
    if col_tipo_error_sel == 'Auto-clasificar':
        df_proc['tipo_error'] = df_proc['descripcion'].apply(clasificar_tipo_error)
    else:
        df_proc['tipo_error'] = df[col_tipo_error_sel].fillna('').apply(clasificar_tipo_error)
    
    # Respuesta
    if col_respuesta_sel != 'No tiene':
        df_proc['respuesta'] = df[col_respuesta_sel].fillna('Sin respuesta').astype(str)
    else:
        df_proc['respuesta'] = 'No disponible'
    
    # Estado
    df_proc['estado'] = df_proc['respuesta'].apply(clasificar_estado)
    
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
    
    # Filtro por producto
    productos_disponibles = sorted(df_proc['producto'].unique())
    productos_filtro = st.sidebar.multiselect(
        "Producto:",
        options=productos_disponibles,
        default=productos_disponibles
    )
    
    # Aplicar filtros
    mask_tipo = df_proc['tipo_error'].isin(tipos_error_filtro)
    mask_estado = df_proc['estado'].isin(estados_filtro)
    mask_producto = df_proc['producto'].isin(productos_filtro)
    df_filtrado = df_proc[mask_fecha & mask_tipo & mask_estado & mask_producto]
    
    # Métricas principales
    st.header("📈 Métricas Principales")
    
    col1, col2, col3, col4 = st.columns(4)
    
    total = len(df_filtrado)
    resueltos = len(df_filtrado[df_filtrado['estado'] == 'Resuelto'])
    en_proceso = len(df_filtrado[df_filtrado['estado'] == 'En Proceso'])
    pendientes = len(df_filtrado[df_filtrado['estado'] == 'Pendiente'])
    tasa = (resueltos / total * 100) if total > 0 else 0
    
    with col1:
        st.metric("Total Quejas", total)
    with col2:
        st.metric("Resueltas", resueltos, delta=f"{tasa:.1f}%")
    with col3:
        st.metric("En Proceso", en_proceso)
    with col4:
        st.metric("Pendientes", pendientes)
    
    st.markdown("---")
    
    # Gráficos
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Quejas por Tipo de Error")
        tipo_counts = df_filtrado['tipo_error'].value_counts()
        st.bar_chart(tipo_counts)
    
    with col2:
        st.subheader("📈 Quejas por Estado")
        estado_counts = df_filtrado['estado'].value_counts()
        st.bar_chart(estado_counts)
    
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("📦 Quejas por Producto")
        producto_counts = df_filtrado['producto'].value_counts()
        st.bar_chart(producto_counts)
    
    with col4:
        st.subheader("📅 Tendencia en el Tiempo")
        df_temporal = df_filtrado[df_filtrado['fecha'].notna()].groupby(df_filtrado['fecha'].dt.date).size()
        if len(df_temporal) > 0:
            st.line_chart(df_temporal)
        else:
            st.info("Sin datos de fechas")
    
    # Tabla completa con correos
    st.markdown("---")
    st.subheader("📋 Detalle Completo de Quejas")
    
    # Preparar tabla para mostrar
    columnas_mostrar = ['fecha', 'email', 'cliente', 'producto', 'tipo_error', 'estado', 'descripcion']
    
    # Solo incluir columnas que existen
    columnas_a_mostrar = [col for col in columnas_mostrar if col in df_filtrado.columns]
    
    df_mostrar = df_filtrado[columnas_a_mostrar].copy()
    
    # Formatear fecha
    df_mostrar['fecha'] = df_mostrar['fecha'].apply(
        lambda x: x.strftime('%Y-%m-%d %H:%M') if pd.notna(x) else 'Sin fecha'
    )
    
    # Truncar descripción larga
    if 'descripcion' in df_mostrar.columns:
        df_mostrar['descripcion'] = df_mostrar['descripcion'].apply(
            lambda x: (str(x)[:100] + '...') if len(str(x)) > 100 else str(x)
        )
    
    # Renombrar columnas para mejor visualización
    nombres_columnas = {
        'fecha': 'Fecha',
        'email': 'Correo Electrónico',
        'cliente': 'Cliente',
        'producto': 'Producto',
        'tipo_error': 'Tipo de Error',
        'estado': 'Estado',
        'descripcion': 'Descripción'
    }
    df_mostrar = df_mostrar.rename(columns=nombres_columnas)
    
    # Mostrar tabla con altura ajustable
    st.dataframe(df_mostrar, use_container_width=True, height=500)
    
    # Resumen por categorías
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Resumen por Tipo de Error")
        resumen_tipo = df_filtrado.groupby('tipo_error').agg({
            'estado': lambda x: (x == 'Resuelto').sum()
        }).reset_index()
        resumen_tipo.columns = ['Tipo de Error', 'Casos Resueltos']
        resumen_tipo['Total Casos'] = df_filtrado['tipo_error'].value_counts().values
        resumen_tipo['% Resolución'] = (resumen_tipo['Casos Resueltos'] / resumen_tipo['Total Casos'] * 100).round(1)
        st.dataframe(resumen_tipo, use_container_width=True)
    
    with col2:
        st.subheader("📦 Resumen por Producto")
        resumen_producto = df_filtrado.groupby('producto').agg({
            'estado': lambda x: (x == 'Resuelto').sum()
        }).reset_index()
        resumen_producto.columns = ['Producto', 'Casos Resueltos']
        resumen_producto['Total Casos'] = df_filtrado['producto'].value_counts().values
        resumen_producto['% Resolución'] = (resumen_producto['Casos Resueltos'] / resumen_producto['Total Casos'] * 100).round(1)
        st.dataframe(resumen_producto, use_container_width=True)
    
    # Exportar
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        csv = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Descargar Datos Completos (CSV)",
            data=csv,
            file_name=f"quejas_segpro_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )
    
    with col2:
        # Exportar solo tabla resumida
        csv_resumen = df_mostrar.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Descargar Tabla Mostrada (CSV)",
            data=csv_resumen,
            file_name=f"tabla_quejas_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

else:
    st.info("👈 Carga tu archivo Excel en el menú lateral")
    
    st.markdown("""
    ### 📝 Instrucciones:
    
    **Este dashboard NO inventa datos.** Solo analiza lo que viene en tu Excel.
    
    **Columnas que puede detectar:**
    - 📅 Fecha de la queja
    - 📧 Email/Correo del cliente
    - 👤 Nombre del cliente
    - 📦 Producto (se normalizará automáticamente)
    - ⚠️ Tipo de error
    - 📝 Descripción del problema
    - ✅ Respuesta o solución
    
    **Productos específicos de SEGPRO:**
    - Guante Multi Flex
    - Zapatos Harder
    - Cono Naranja
    
    **Tipos de error clasificados:**
    - Producto defectuoso
    - Error de talla
    - Pieza faltante
    - Color incorrecto
    - Producto no coincide con lo solicitado
    - Daño por transporte
    - Producto con fallas de fábrica
    """)

st.markdown("---")
st.markdown("<div style='text-align: center; color: gray;'>Dashboard SEGPRO © 2024 | Datos reales del Excel</div>", unsafe_allow_html=True)
