import streamlit as st
import pandas as pd
from datetime import datetime
import psycopg2 
from psycopg2 import extras
import time
import unicodedata
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Base de Datos - Camacho Construcciones",
    page_icon="📊",
    layout="wide"
)

if 'ultimo_eliminado' not in st.session_state:
    st.session_state['ultimo_eliminado'] = None

# =========================================================
# 1. ENLACE DE CONEXIÓN (SECRETS DE STREAMLIT)
# =========================================================
DB_URL = st.secrets.get("DATABASE_URL")

if not DB_URL:
    st.error("⚠️ Faltan los Secrets. Por favor, ve a la configuración de Streamlit y asegúrate de haber puesto tu DATABASE_URL correctamente.")
    st.stop() 

def obtener_conexion():
    return psycopg2.connect(DB_URL)

# =========================================================
# 2. FUNCIONES Y HERRAMIENTAS
# =========================================================
def mostrar_mensaje_central(mensaje, tipo="success"):
    color_fondo = "rgba(40, 167, 69, 0.95)" if tipo == "success" else "rgba(220, 53, 69, 0.95)" if tipo == "error" else "rgba(0, 123, 255, 0.95)"
    marcador = st.empty()
    marcador.markdown(f"""
        <div style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); 
                    background-color: {color_fondo}; color: white; padding: 20px 40px; 
                    border-radius: 10px; z-index: 99999; box-shadow: 0px 8px 20px rgba(0,0,0,0.3);
                    pointer-events: none; font-size: 1.3rem; font-weight: bold; text-align: center;">
            {mensaje}
        </div>
    """, unsafe_allow_html=True)
    time.sleep(2)
    marcador.empty()

@st.dialog("⚠️ Confirmar eliminación")
def confirmar_eliminar(id_registro, fila):
    st.write("¿Estás seguro de que deseas eliminar este registro de forma permanente?")
    st.info(f"👤 **Nombre:** {fila['NOMBRE']}\n\n💳 **Cédula:** {fila['CEDULA']}\n\n🆔 **ID:** {fila['ID']}\n\n💰 **Valor:** ${fila['VALOR']:,.0f}")
    
    col1, col2 = st.columns(2)
    if col1.button("✅ Sí, eliminar"):
        st.session_state['ultimo_eliminado'] = fila.to_dict()
        
        # Eliminación dirigida mediante SQL directo
        conn = obtener_conexion()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM ventas WHERE \"ID\" = %s", (str(id_registro),))
        conn.commit()
        conn.close()
        
        st.cache_data.clear()
        st.toast("🗑️ Registro eliminado. (Puedes deshacerlo en la pestaña de Edición)", icon="🚨")
        st.rerun()
    if col2.button("❌ No, cancelar"):
        st.rerun()

# =========================================================
# 3. CONFIGURACIÓN DE LA PÁGINA Y BASE DE DATOS
# =========================================================
st.markdown("<h1 style='font-size: 45px; font-weight: bold;'>Base de Datos - Camacho Construcciones</h1>", unsafe_allow_html=True)
st.image("logo.png", width=180) 

@st.cache_data(ttl=600)
def cargar_datos():
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT EXISTS (
            SELECT FROM information_schema.tables 
            WHERE table_schema = 'public' AND table_name = 'ventas'
        );
    """)
    tabla_existe = cursor.fetchone()[0]
    
    if not tabla_existe:
        st.info("🔄 Migrando datos iniciales a la nube de Neon. Esto puede tardar unos segundos...")
        SHEET_ID = "1fgY5F9PsYMu-7mff8vzAieA_yezmo0-D"
        URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
        df = pd.read_csv(URL)
        df.columns = [c.strip() for c in df.columns]
        
        columnas_validas = [c for c in df.columns if not c.startswith("Unnamed")]
        df = df[columnas_validas]
        
        from sqlalchemy import create_engine
        engine = create_engine(DB_URL.replace("postgresql://", "postgresql+psycopg2://"))
        df.to_sql('ventas', engine, if_exists='replace', index=False)
        conn.close()
        return df
    else:
        df = pd.read_sql("SELECT * FROM ventas", conn)
    
    conn.close()
    
    columnas_validas = [c for c in df.columns if not c.startswith("Unnamed")]
    df = df[columnas_validas]
    
    df["FECHA_DT"] = pd.to_datetime(df["FECHA"], errors='coerce', dayfirst=True)
    df["ANIO"] = df["FECHA_DT"].dt.year.fillna(datetime.now().year).astype(int)
    
    for col in ["CEDULA", "TELEFONO", "ID"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')

    if "VALOR" in df.columns:
        df["VALOR"] = df["VALOR"].astype(str).replace({r'\$': '', r',': '', r'\s': ''}, regex=True)
        df["VALOR"] = pd.to_numeric(df["VALOR"], errors='coerce').fillna(0)
    else:
        df["VALOR"] = 0
        
    df["Pago Vanti (8%)"] = df["VALOR"] * 0.08
    df["Comision Vendedor (35%)"] = df["VALOR"] * 0.35
    df["Ganancia Camacho"] = df["VALOR"] * 0.57
            
    return df

def enviar_respaldo_automatico(df_limpio):
    correo_emisor = "camachoconstruccionespqr@gmail.com"
    password = "inunlpcwyxqhwnkx"
    correo_receptor = "camachoconstruccionespqr@gmail.com"
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_limpio.to_excel(writer, index=False)
    buffer.seek(0)
    
    msg = MIMEMultipart()
    msg['From'] = correo_emisor
    msg['To'] = correo_receptor
    msg['Subject'] = f"📆 RESPALDO SEMANAL AUTOMÁTICO - {datetime.now().strftime('%d/%m/%Y')}"
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(buffer.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= Respaldo_Camacho_Semanal.xlsx")
    msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(correo_emisor, password)
        server.sendmail(correo_emisor, correo_receptor, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        return False

def verificar_y_ejecutar_respaldo():
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS control_respaldo (id INTEGER PRIMARY KEY, fecha TEXT)''')
    conn.commit()
    
    cursor.execute("SELECT fecha FROM control_respaldo WHERE id = 1")
    resultado = cursor.fetchone()
    fecha_hoy = datetime.now().date()
    debe_respaldar = False
    
    if resultado is None:
        debe_respaldar = True
    else:
        ultima_fecha = datetime.strptime(resultado[0], "%Y-%m-%d").date()
        if (fecha_hoy - ultima_fecha).days >= 7:
            debe_respaldar = True
            
    if debe_respaldar:
        df_respaldo = pd.read_sql("SELECT * FROM ventas", conn)
        if enviar_respaldo_automatico(df_respaldo):
            cursor.execute("INSERT INTO control_respaldo (id, fecha) VALUES (1, %s) ON CONFLICT (id) DO UPDATE SET fecha = EXCLUDED.fecha", (fecha_hoy.strftime("%Y-%m-%d"),))
            conn.commit()
    conn.close()

df_base = cargar_datos()
df_base = df_base.sort_values(by="FECHA_DT", ascending=False, na_position='last').reset_index(drop=True)

# =========================================================
# 4. INTERFAZ Y PESTAÑAS
# =========================================================
tab_visualizar, tab_registrar, tab_editar = st.tabs(["📊 Datos Generales", "➕ Registrar Nuevo", "✏️ Editar o Eliminar"])

with tab_visualizar:
    col_titulo, col_boton = st.columns([3, 1])
    with col_titulo:
        st.subheader("💰 Resumen Financiero")
    with col_boton:
        csv = df_base.drop(columns=["FECHA_DT", "ANIO"], errors='ignore').to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar a Excel",
            data=csv,
            file_name=f'Base_Camacho_{datetime.now().strftime("%d-%m-%Y")}.csv',
            mime='text/csv',
            use_container_width=True
        )
    
    opcion_metricas = st.radio("Selecciona el período a consultar:", ["Esta Semana", "Este Mes", "Este Año", "Total"], horizontal=True)
    hoy = pd.Timestamp(datetime.now().date())
    df_metricas = df_base.copy()
    texto_periodo = ""
    
    if opcion_metricas == "Esta Semana":
        inicio_semana = hoy - pd.Timedelta(days=hoy.weekday())
        fin_semana = inicio_semana + pd.Timedelta(days=6)
        df_metricas = df_metricas[(df_metricas["FECHA_DT"] >= inicio_semana) & (df_metricas["FECHA_DT"] <= fin_semana)]
        texto_periodo = f"Del Lunes {inicio_semana.strftime('%d/%m/%Y')} al Domingo {fin_semana.strftime('%d/%m/%Y')}"
    elif opcion_metricas == "Este Mes":
        df_metricas = df_metricas[(df_metricas["FECHA_DT"].dt.month == hoy.month) & (df_metricas["FECHA_DT"].dt.year == hoy.year)]
        texto_periodo = "Mes actual"
    elif opcion_metricas == "Este Año":
        df_metricas = df_metricas[df_metricas["FECHA_DT"].dt.year == hoy.year]
        texto_periodo = f"Año {hoy.year}"
    else: 
        texto_periodo = "Todo el histórico"
        
    st.caption(f"📅 Calculando: {texto_periodo}")
    
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Ventas ({opcion_metricas})", f"${df_metricas['VALOR'].sum():,.0f}")
    c2.metric("A Pagar a Vanti (8%)", f"${df_metricas['Pago Vanti (8%)'].sum():,.0f}")
    c3.metric("Ganancia Limpia (57%)", f"${df_metricas['Ganancia Camacho'].sum():,.0f}")
    
    st.divider()
    st.subheader("📋 Base de Datos y Buscador General")
    
    col_filtro1, col_filtro2 = st.columns(2)
    with col_filtro1:
        anio_actual = datetime.now().year
        lista_anios = ["TODOS"] + sorted(df_base["ANIO"].unique().tolist(), reverse=True)
        index_defecto = lista_anios.index(anio_actual) if anio_actual in lista_anios else 1
        anio_seleccionado = st.selectbox("📅 Filtrar tabla por Año:", lista_anios, index=index_defecto)
        
    with col_filtro2:
        criterio_busqueda = st.text_input("👤 Buscar en la tabla por Nombre, Cédula, ID o Vendedor:")

    df_filtrado = df_base.copy()
    if anio_seleccionado != "TODOS":
        df_filtrado = df_filtrado[df_filtrado["ANIO"] == int(anio_seleccionado)]
        
    if criterio_busqueda:
        def limpiar_texto(texto):
            if pd.isna(texto): return ""
            return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').lower()
        criterio_limpio = limpiar_texto(criterio_busqueda)
        df_filtrado = df_filtrado[
            df_filtrado.apply(lambda fila: 
                criterio_limpio in limpiar_texto(fila["NOMBRE"]) or
                criterio_limpio in limpiar_texto(fila["CEDULA"]) or
                criterio_limpio in limpiar_texto(fila["ID"]) or
                criterio_limpio in limpiar_texto(fila["VENDEDOR"]), 
            axis=1)
        ]

    df_vista = df_filtrado.copy()
    for col in ["VALOR", "Pago Vanti (8%)", "Comision Vendedor (35%)", "Ganancia Camacho"]:
        if col in df_vista.columns:
            df_vista[col] = df_vista[col].map("${:,.0f}".format)

    st.dataframe(df_vista.drop(columns=["FECHA_DT", "ANIO"], errors='ignore'), use_container_width=True, height=500)

    if not df_metricas.empty:
        st.write("")
        col_graf1, col_graf2 = st.columns(2)
        with col_graf1:
            st.caption("📊 Cantidad de Ventas Realizadas por Vendedor")
            df_graf_vendedores = df_metricas.groupby("VENDEDOR").size().reset_index(name="VENTAS REALIZADAS")
            st.bar_chart(df_graf_vendedores.set_index("VENDEDOR"))
        with col_graf2:
            st.caption("📈 Comparativo de Ventas vs Período Anterior")
            df_comparativo = df_base.copy()
            if opcion_metricas == "Esta Semana":
                inicio_act = hoy - pd.Timedelta(days=hoy.weekday())
                df_act = df_comparativo[df_comparativo["FECHA_DT"] >= inicio_act]
                df_ant = df_comparativo[(df_comparativo["FECHA_DT"] >= (inicio_act - pd.Timedelta(days=7))) & (df_comparativo["FECHA_DT"] < inicio_act)]
                nombres_dias = {0: "1-Lun", 1: "2-Mar", 2: "3-Mié", 3: "4-Jue", 4: "5-Vie", 5: "6-Sáb", 6: "7-Dom"}
                s_act = df_act.groupby(df_act["FECHA_DT"].dt.weekday).size()
                s_ant = df_ant.groupby(df_ant["FECHA_DT"].dt.weekday).size()
                df_plot = pd.DataFrame({"Esta Semana": s_act, "Semana Pasada": s_ant}).fillna(0)
                df_plot.index = df_plot.index.map(nombres_dias)
            elif opcion_metricas == "Este Mes":
                df_act = df_comparativo[(df_comparativo["FECHA_DT"].dt.month == hoy.month) & (df_comparativo["FECHA_DT"].dt.year == hoy.year)]
                mes_ant = hoy.month - 1 if hoy.month > 1 else 12
                anio_ant = hoy.year if hoy.month > 1 else hoy.year - 1
                df_ant = df_comparativo[(df_comparativo["FECHA_DT"].dt.month == mes_ant) & (df_comparativo["FECHA_DT"].dt.year == anio_ant)]
                s_act = df_act.groupby(df_act["FECHA_DT"].dt.day).size()
                s_ant = df_ant.groupby(df_ant["FECHA_DT"].dt.day).size()
                df_plot = pd.DataFrame({"Este Mes": s_act, "Mes Pasado": s_ant}).fillna(0)
            elif opcion_metricas == "Este Año":
                df_act = df_comparativo[df_comparativo["FECHA_DT"].dt.year == hoy.year]
                df_ant = df_comparativo[df_comparativo["FECHA_DT"].dt.year == (hoy.year - 1)]
                nombres_meses = {1:"01-Ene", 2:"02-Feb", 3:"03-Mar", 4:"04-Abr", 5:"05-May", 6:"06-Jun", 7:"07-Jul", 8:"08-Ago", 9:"09-Sep", 10:"10-Oct", 11:"11-Nov", 12:"12-Dic"}
                s_act = df_act.groupby(df_act["FECHA_DT"].dt.month).size()
                s_ant = df_ant.groupby(df_ant["FECHA_DT"].dt.month).size()
                df_plot = pd.DataFrame({"Este Año": s_act, "Año Pasado": s_ant}).fillna(0)
                df_plot.index = df_plot.index.map(nombres_meses)
            else:
                s_act = df_comparativo.groupby(df_comparativo["FECHA_DT"].dt.year).size()
                df_plot = pd.DataFrame({"Ventas Históricas Totales": s_act}).fillna(0)
            st.line_chart(df_plot)

with tab_registrar:
    st.subheader("📝 Agregar Nuevo Registro")
    with st.form("form_nuevo", clear_on_submit=True):
        c_reg1, c_reg2, c_reg3 = st.columns(3)
        with c_reg1:
            f_fecha = st.date_input("FECHA", format="DD/MM/YYYY")
            f_cedula = st.text_input("CEDULA")
            f_id = st.text_input("ID / RADICADO")
            f_nombre = st.text_input("NOMBRE DEL CLIENTE")
        with c_reg2:
            f_direccion = st.text_input("DIRECCION")
            f_telefono = st.text_input("TELEFONO")
            f_tecnico = st.text_input("TECNICO ASIGNADO")
            f_valor = st.number_input("VALOR DE LA VENTA ($)", min_value=0, step=10000)
        with c_reg3:
            f_vendedor = st.text_input("VENDEDOR / ASESOR")
            f_est_final = st.text_input("ESTADO FINAL")
            f_est_pago = st.text_input("ESTADO PAGO")
            f_cuenta = st.text_input("NUMERO DE CUENTA")
            
        f_observaciones = st.text_area("OBSERVACIONES")
        
        if st.form_submit_button("💾 Guardar en Base de Datos"):
            # Inserción eficiente de una única fila
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO ventas ("FECHA", "CEDULA", "ID", "NOMBRE", "DIRECCION", "TELEFONO", "TECNICO", "VALOR", "VENDEDOR", "ESTADO FINAL", "ESTADO PAGO", "NUMERO DE CUENTA", "OBSERVACIONES")
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                f"{f_fecha.day}/{f_fecha.month}/{f_fecha.year}", f_cedula, f_id, f_nombre.upper(),
                f_direccion.upper(), f_telefono, f_tecnico.upper(), float(f_valor), f_vendedor.upper(),
                f_est_final.upper(), f_est_pago.upper(), f_cuenta, f_observaciones
            ))
            conn.commit()
            conn.close()
            
            st.cache_data.clear()
            verificar_y_ejecutar_respaldo()
            mostrar_mensaje_central("✅ ¡Cliente registrado con éxito!", "success")
            st.rerun()

with tab_editar:
    st.subheader("⚙️ Modificar o Eliminar un Registro")
    if st.session_state['ultimo_eliminado'] is not None:
        st.warning("⚠️ Acabas de eliminar un registro recientemente. ¿Deseas recuperarlo?")
        c_des1, c_des2 = st.columns(2)
        with c_des1:
            if st.button("↩️ Sí, deshacer y recuperar el registro"):
                recup = st.session_state['ultimo_eliminado']
                
                conn = obtener_conexion()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO ventas ("FECHA", "CEDULA", "ID", "NOMBRE", "DIRECCION", "TELEFONO", "TECNICO", "VALOR", "VENDEDOR", "ESTADO FINAL", "ESTADO PAGO", "NUMERO DE CUENTA", "OBSERVACIONES")
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    recup["FECHA"], recup["CEDULA"], recup["ID"], recup["NOMBRE"],
                    recup["DIRECCION"], recup["TELEFONO"], recup["TECNICO"], float(recup["VALOR"]),
                    recup["VENDEDOR"], recup["ESTADO FINAL"], recup["ESTADO PAGO"], recup["NUMERO DE CUENTA"], recup["OBSERVACIONES"]
                ))
                conn.commit()
                conn.close()
                
                st.cache_data.clear()
                st.session_state['ultimo_eliminado'] = None
                mostrar_mensaje_central("✅ ¡Registro recuperado con éxito!", "success")
                st.rerun()
        with c_des2:
            if st.button("🗑️ No, dejarlo eliminado"):
                st.session_state['ultimo_eliminado'] = None
                st.rerun()
        st.divider()

    if not df_base.empty:
        diccionario_busqueda = {}
        for index, row in df_base.iterrows():
            texto_busqueda = f"👤 {row['NOMBRE']} | 🆔 ID: {row['ID']} | 💳 Céd: {row['CEDULA']} | 💼 Vend: {row['VENDEDOR']}"
            diccionario_busqueda[texto_busqueda] = row['ID']
            
        opcion_seleccionada = st.selectbox("🔍 Escribe aquí para buscar el registro:", options=list(diccionario_busqueda.keys()))
        id_real = diccionario_busqueda[opcion_seleccionada]
        fila_datos = df_base[df_base['ID'] == id_real].iloc[0]
        
        texto_whatsapp = (
            f"👤 *Nombre:* {fila_datos['NOMBRE']}\n"
            f"💳 *Cédula:* {fila_datos['CEDULA']}\n"
            f"🆔 *ID/Radicado:* {fila_datos['ID']}\n"
            f"🏠 *Dirección:* {fila_datos['DIRECCION']}\n"
            f"📞 *Teléfono:* {fila_datos['TELEFONO']}\n"
            f"💰 *Valor:* ${fila_datos['VALOR']:,.0f}"
        )
        
        st.caption("📲 Presiona el ícono de copiar (arriba a la derecha) para enviarlo por WhatsApp:")
        st.code(texto_whatsapp, language="markdown")
        st.write("")
        
        with st.form("form_editar"):
            c_ed1, c_ed2, c_ed3 = st.columns(3)
            with c_ed1:
                e_fecha = st.text_input("FECHA (DD/MM/YYYY)", value=str(fila_datos["FECHA"]))
                e_cedula = st.text_input("CEDULA", value=str(fila_datos["CEDULA"]))
                e_nombre = st.text_input("NOMBRE", value=str(fila_datos["NOMBRE"]))
            with c_ed2:
                e_direccion = st.text_input("DIRECCION", value=str(fila_datos["DIRECCION"]))
                e_telefono = st.text_input("TELEFONO", value=str(fila_datos["TELEFONO"]))
                e_tecnico = st.text_input("TECNICO", value=str(fila_datos["TECNICO"]))
            with c_ed3:
                e_valor = st.number_input("VALOR DE VENTA ($)", value=int(fila_datos["VALOR"]), step=5000)
                e_vendedor = st.text_input("VENDEDOR", value=str(fila_datos["VENDEDOR"]))
                e_est_final = st.text_input("ESTADO FINAL", value=str(fila_datos["ESTADO FINAL"]))
                
            c_ed4, c_ed5 = st.columns(2)
            with c_ed4:
                e_est_pago = st.text_input("ESTADO PAGO", value=str(fila_datos["ESTADO PAGO"]))
                e_cuenta = st.text_input("NUMERO DE CUENTA", value=str(fila_datos["NUMERO DE CUENTA"]))
            with c_ed5:
                e_observaciones = st.text_area("OBSERVACIONES", value=str(fila_datos["OBSERVACIONES"]))
                
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                btn_actualizar = st.form_submit_button("🔄 Guardar Cambios Permanentes")
            with col_b2:
                btn_eliminar = st.form_submit_button("❌ Eliminar Registro Definitivamente")
                
            if btn_actualizar:
                # Update dirigido únicamente a las columnas del ID correspondiente
                conn = obtener_conexion()
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE ventas 
                    SET "FECHA"=%s, "CEDULA"=%s, "NOMBRE"=%s, "DIRECCION"=%s, "TELEFONO"=%s, "TECNICO"=%s, 
                        "VALOR"=%s, "VENDEDOR"=%s, "ESTADO FINAL"=%s, "ESTADO PAGO"=%s, "NUMERO DE CUENTA"=%s, "OBSERVACIONES"=%s
                    WHERE "ID"=%s
                """, (
                    e_fecha, e_cedula, e_nombre.upper(), e_direccion.upper(), e_telefono, e_tecnico.upper(),
                    float(e_valor), e_vendedor.upper(), e_est_final.upper(), e_est_pago.upper(), e_cuenta, e_observaciones,
                    str(id_real)
                ))
                conn.commit()
                conn.close()
                
                st.cache_data.clear()
                verificar_y_ejecutar_respaldo()
                mostrar_mensaje_central("🔄 ¡Registro actualizado correctamente!", "success")
                st.rerun()
                
            if btn_eliminar:
                confirmar_eliminar(id_real, fila_datos)
