import streamlit as st
import pandas as pd
from datetime import datetime
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
# 1. CONEXIÓN ÚNICA Y ROBUSTA (reemplaza obtener_conexion)
# =========================================================
# En st.secrets debes tener:
# [connections.neon]
# dialect = "postgresql"
# url = "postgresql://..."   ← tu DATABASE_URL actual

conn_st = st.connection("neon", type="sql")

def ejecutar_query(sql, params=None):
    """Ejecuta INSERT, UPDATE o DELETE con reconexión automática."""
    with conn_st.session as session:
        session.execute(st.text(sql), params or {})
        session.commit()

# =========================================================
# 2. FUNCIONES Y HERRAMIENTAS
# =========================================================
def mostrar_mensaje_central(mensaje, tipo="success"):
    """Muestra toast instantáneo sin bloquear la UI con sleep."""
    icono = "✅" if tipo == "success" else "❌" if tipo == "error" else "ℹ️"
    st.toast(mensaje, icon=icono)

@st.dialog("⚠️ Confirmar eliminación")
def confirmar_eliminar(id_registro, fila):
    st.write("¿Estás seguro de que deseas eliminar este registro de forma permanente?")
    st.info(f"👤 **Nombre:** {fila['NOMBRE']}\n\n💳 **Cédula:** {fila['CEDULA']}\n\n🆔 **ID:** {fila['ID']}\n\n💰 **Valor:** ${fila['VALOR']:,.0f}")
    
    col1, col2 = st.columns(2)
    if col1.button("✅ Sí, eliminar"):
        st.session_state['ultimo_eliminado'] = fila.to_dict()
        ejecutar_query('DELETE FROM ventas WHERE "ID" = :id', {"id": str(id_registro)})
        st.cache_data.clear()
        st.toast("🗑️ Registro eliminado. (Puedes deshacerlo en la pestaña de Edición)", icon="🚨")
        st.rerun()
    if col2.button("❌ No, cancelar"):
        st.rerun()

# =========================================================
# 3. BASE DE DATOS
# =========================================================
st.markdown("<h1 style='font-size: 45px; font-weight: bold;'>Base de Datos - Camacho Construcciones</h1>", unsafe_allow_html=True)
st.image("logo.png", width=180)

@st.cache_data(ttl=600)
def cargar_datos():
    df = conn_st.query("SELECT * FROM ventas", ttl=0)

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

    df["Pago Vanti (8%)"]          = df["VALOR"] * 0.08
    df["Comision Vendedor (35%)"]  = df["VALOR"] * 0.35
    df["Ganancia Camacho"]         = df["VALOR"] * 0.57

    return df

def enviar_respaldo_automatico(df_limpio):
    correo_emisor   = "camachoconstruccionespqr@gmail.com"
    # ✅ SEGURIDAD: contraseña en secrets, no en el código
    password        = st.secrets.get("GMAIL_PASSWORD", "")
    correo_receptor = "camachoconstruccionespqr@gmail.com"

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_limpio.to_excel(writer, index=False)
    buffer.seek(0)

    msg = MIMEMultipart()
    msg['From']    = correo_emisor
    msg['To']      = correo_receptor
    msg['Subject'] = f"📆 RESPALDO SEMANAL AUTOMÁTICO - {datetime.now().strftime('%d/%m/%Y')}"

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(buffer.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename=Respaldo_Camacho_Semanal.xlsx")
    msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(correo_emisor, password)
        server.sendmail(correo_emisor, correo_receptor, msg.as_string())
        server.quit()
        return True
    except Exception:
        return False

def verificar_y_ejecutar_respaldo():
    """
    Corre el respaldo solo si hace más de 7 días.
    Usa conn_st para no abrir conexión extra.
    """
    conn_st.query(
        "CREATE TABLE IF NOT EXISTS control_respaldo (id INTEGER PRIMARY KEY, fecha TEXT)",
        ttl=0
    )
    resultado = conn_st.query("SELECT fecha FROM control_respaldo WHERE id = 1", ttl=0)
    fecha_hoy = datetime.now().date()

    debe_respaldar = resultado.empty
    if not debe_respaldar:
        ultima_fecha = datetime.strptime(resultado.iloc[0]["fecha"], "%Y-%m-%d").date()
        if (fecha_hoy - ultima_fecha).days >= 7:
            debe_respaldar = True

    if debe_respaldar:
        df_respaldo = conn_st.query("SELECT * FROM ventas", ttl=0)
        if enviar_respaldo_automatico(df_respaldo):
            ejecutar_query(
                "INSERT INTO control_respaldo (id, fecha) VALUES (1, :fecha) "
                "ON CONFLICT (id) DO UPDATE SET fecha = EXCLUDED.fecha",
                {"fecha": fecha_hoy.strftime("%Y-%m-%d")}
            )

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
        fin_semana    = inicio_semana + pd.Timedelta(days=6)
        df_metricas   = df_metricas[(df_metricas["FECHA_DT"] >= inicio_semana) & (df_metricas["FECHA_DT"] <= fin_semana)]
        texto_periodo = f"Del Lunes {inicio_semana.strftime('%d/%m/%Y')} al Domingo {fin_semana.strftime('%d/%m/%Y')}"
    elif opcion_metricas == "Este Mes":
        df_metricas   = df_metricas[(df_metricas["FECHA_DT"].dt.month == hoy.month) & (df_metricas["FECHA_DT"].dt.year == hoy.year)]
        texto_periodo = "Mes actual"
    elif opcion_metricas == "Este Año":
        df_metricas   = df_metricas[df_metricas["FECHA_DT"].dt.year == hoy.year]
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
        anio_actual       = datetime.now().year
        lista_anios       = ["TODOS"] + sorted(df_base["ANIO"].unique().tolist(), reverse=True)
        index_defecto     = lista_anios.index(anio_actual) if anio_actual in lista_anios else 1
        anio_seleccionado = st.selectbox("📅 Filtrar tabla por Año:", lista_anios, index=index_defecto)
    with col_filtro2:
        criterio_busqueda = st.text_input("👤 Buscar por Nombre, Cédula, ID o Vendedor:")

    df_filtrado = df_base.copy()
    if anio_seleccionado != "TODOS":
        df_filtrado = df_filtrado[df_filtrado["ANIO"] == int(anio_seleccionado)]

    if criterio_busqueda:
        # ✅ MEJORA: búsqueda vectorizada en vez de apply+lambda fila por fila
        def normalizar_serie(serie):
            return serie.fillna("").astype(str).apply(
                lambda t: ''.join(c for c in unicodedata.normalize('NFD', t)
                                  if unicodedata.category(c) != 'Mn').lower()
            )
        criterio_limpio = ''.join(
            c for c in unicodedata.normalize('NFD', criterio_busqueda)
            if unicodedata.category(c) != 'Mn'
        ).lower()

        mask = (
            normalizar_serie(df_filtrado["NOMBRE"]).str.contains(criterio_limpio, regex=False) |
            normalizar_serie(df_filtrado["CEDULA"]).str.contains(criterio_limpio, regex=False) |
            normalizar_serie(df_filtrado["ID"]).str.contains(criterio_limpio, regex=False) |
            normalizar_serie(df_filtrado["VENDEDOR"]).str.contains(criterio_limpio, regex=False)
        )
        df_filtrado = df_filtrado[mask]

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
                inicio_act   = hoy - pd.Timedelta(days=hoy.weekday())
                df_act       = df_comparativo[df_comparativo["FECHA_DT"] >= inicio_act]
                df_ant       = df_comparativo[(df_comparativo["FECHA_DT"] >= (inicio_act - pd.Timedelta(days=7))) & (df_comparativo["FECHA_DT"] < inicio_act)]
                nombres_dias = {0:"1-Lun", 1:"2-Mar", 2:"3-Mié", 3:"4-Jue", 4:"5-Vie", 5:"6-Sáb", 6:"7-Dom"}
                s_act        = df_act.groupby(df_act["FECHA_DT"].dt.weekday).size()
                s_ant        = df_ant.groupby(df_ant["FECHA_DT"].dt.weekday).size()
                df_plot      = pd.DataFrame({"Esta Semana": s_act, "Semana Pasada": s_ant}).fillna(0)
                df_plot.index = df_plot.index.map(nombres_dias)
            elif opcion_metricas == "Este Mes":
                df_act   = df_comparativo[(df_comparativo["FECHA_DT"].dt.month == hoy.month) & (df_comparativo["FECHA_DT"].dt.year == hoy.year)]
                mes_ant  = hoy.month - 1 if hoy.month > 1 else 12
                anio_ant = hoy.year if hoy.month > 1 else hoy.year - 1
                df_ant   = df_comparativo[(df_comparativo["FECHA_DT"].dt.month == mes_ant) & (df_comparativo["FECHA_DT"].dt.year == anio_ant)]
                s_act    = df_act.groupby(df_act["FECHA_DT"].dt.day).size()
                s_ant    = df_ant.groupby(df_ant["FECHA_DT"].dt.day).size()
                df_plot  = pd.DataFrame({"Este Mes": s_act, "Mes Pasado": s_ant}).fillna(0)
            elif opcion_metricas == "Este Año":
                df_act        = df_comparativo[df_comparativo["FECHA_DT"].dt.year == hoy.year]
                df_ant        = df_comparativo[df_comparativo["FECHA_DT"].dt.year == (hoy.year - 1)]
                nombres_meses = {1:"01-Ene", 2:"02-Feb", 3:"03-Mar", 4:"04-Abr", 5:"05-May", 6:"06-Jun", 7:"07-Jul", 8:"08-Ago", 9:"09-Sep", 10:"10-Oct", 11:"11-Nov", 12:"12-Dic"}
                s_act         = df_act.groupby(df_act["FECHA_DT"].dt.month).size()
                s_ant         = df_ant.groupby(df_ant["FECHA_DT"].dt.month).size()
                df_plot       = pd.DataFrame({"Este Año": s_act, "Año Pasado": s_ant}).fillna(0)
                df_plot.index = df_plot.index.map(nombres_meses)
            else:
                s_act   = df_comparativo.groupby(df_comparativo["FECHA_DT"].dt.year).size()
                df_plot = pd.DataFrame({"Ventas Históricas Totales": s_act}).fillna(0)
            st.line_chart(df_plot)

with tab_registrar:
    st.subheader("📝 Agregar Nuevo Registro")
    with st.form("form_nuevo", clear_on_submit=True):
        c_reg1, c_reg2, c_reg3 = st.columns(3)
        with c_reg1:
            f_fecha      = st.date_input("FECHA", format="DD/MM/YYYY")
            f_cedula     = st.text_input("CEDULA")
            f_id         = st.text_input("ID / RADICADO")
            f_nombre     = st.text_input("NOMBRE DEL CLIENTE")
        with c_reg2:
            f_direccion  = st.text_input("DIRECCION")
            f_telefono   = st.text_input("TELEFONO")
            f_tecnico    = st.text_input("TECNICO ASIGNADO")
            f_valor      = st.number_input("VALOR DE LA VENTA ($)", min_value=0, step=10000)
        with c_reg3:
            f_vendedor   = st.text_input("VENDEDOR / ASESOR")
            f_est_final  = st.text_input("ESTADO FINAL")
            f_est_pago   = st.text_input("ESTADO PAGO")
            f_cuenta     = st.text_input("NUMERO DE CUENTA")

        f_observaciones = st.text_area("OBSERVACIONES")

        if st.form_submit_button("💾 Guardar en Base de Datos"):
            ejecutar_query("""
                INSERT INTO ventas ("FECHA","CEDULA","ID","NOMBRE","DIRECCION","TELEFONO","TECNICO",
                                    "VALOR","VENDEDOR","ESTADO FINAL","ESTADO PAGO","NUMERO DE CUENTA","OBSERVACIONES")
                VALUES (:fecha,:cedula,:id,:nombre,:direccion,:telefono,:tecnico,
                        :valor,:vendedor,:est_final,:est_pago,:cuenta,:observaciones)
            """, {
                "fecha":        f"{f_fecha.day}/{f_fecha.month}/{f_fecha.year}",
                "cedula":       f_cedula,
                "id":           f_id,
                "nombre":       f_nombre.upper(),
                "direccion":    f_direccion.upper(),
                "telefono":     f_telefono,
                "tecnico":      f_tecnico.upper(),
                "valor":        float(f_valor),
                "vendedor":     f_vendedor.upper(),
                "est_final":    f_est_final.upper(),
                "est_pago":     f_est_pago.upper(),
                "cuenta":       f_cuenta,
                "observaciones": f_observaciones,
            })
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
                ejecutar_query("""
                    INSERT INTO ventas ("FECHA","CEDULA","ID","NOMBRE","DIRECCION","TELEFONO","TECNICO",
                                        "VALOR","VENDEDOR","ESTADO FINAL","ESTADO PAGO","NUMERO DE CUENTA","OBSERVACIONES")
                    VALUES (:fecha,:cedula,:id,:nombre,:direccion,:telefono,:tecnico,
                            :valor,:vendedor,:est_final,:est_pago,:cuenta,:observaciones)
                """, {
                    "fecha":        recup["FECHA"],    "cedula":       recup["CEDULA"],
                    "id":           recup["ID"],       "nombre":       recup["NOMBRE"],
                    "direccion":    recup["DIRECCION"],"telefono":     recup["TELEFONO"],
                    "tecnico":      recup["TECNICO"],  "valor":        float(recup["VALOR"]),
                    "vendedor":     recup["VENDEDOR"], "est_final":    recup["ESTADO FINAL"],
                    "est_pago":     recup["ESTADO PAGO"], "cuenta":    recup["NUMERO DE CUENTA"],
                    "observaciones": recup["OBSERVACIONES"],
                })
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
        # ✅ MEJORA: construcción vectorizada del diccionario de búsqueda
        opciones = (
            "👤 " + df_base["NOMBRE"] +
            " | 🆔 ID: " + df_base["ID"] +
            " | 💳 Céd: " + df_base["CEDULA"] +
            " | 💼 Vend: " + df_base["VENDEDOR"]
        )
        diccionario_busqueda = dict(zip(opciones, df_base["ID"]))

        opcion_seleccionada = st.selectbox("🔍 Escribe aquí para buscar el registro:", options=list(diccionario_busqueda.keys()))
        id_real    = diccionario_busqueda[opcion_seleccionada]
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
                e_fecha   = st.text_input("FECHA (DD/MM/YYYY)", value=str(fila_datos["FECHA"]))
                e_cedula  = st.text_input("CEDULA", value=str(fila_datos["CEDULA"]))
                e_nombre  = st.text_input("NOMBRE", value=str(fila_datos["NOMBRE"]))
            with c_ed2:
                e_direccion = st.text_input("DIRECCION", value=str(fila_datos["DIRECCION"]))
                e_telefono  = st.text_input("TELEFONO", value=str(fila_datos["TELEFONO"]))
                e_tecnico   = st.text_input("TECNICO", value=str(fila_datos["TECNICO"]))
            with c_ed3:
                e_valor     = st.number_input("VALOR DE VENTA ($)", value=int(fila_datos["VALOR"]), step=5000)
                e_vendedor  = st.text_input("VENDEDOR", value=str(fila_datos["VENDEDOR"]))
                e_est_final = st.text_input("ESTADO FINAL", value=str(fila_datos["ESTADO FINAL"]))

            c_ed4, c_ed5 = st.columns(2)
            with c_ed4:
                e_est_pago = st.text_input("ESTADO PAGO", value=str(fila_datos["ESTADO PAGO"]))
                e_cuenta   = st.text_input("NUMERO DE CUENTA", value=str(fila_datos["NUMERO DE CUENTA"]))
            with c_ed5:
                e_observaciones = st.text_area("OBSERVACIONES", value=str(fila_datos["OBSERVACIONES"]))

            col_b1, col_b2 = st.columns(2)
            with col_b1:
                btn_actualizar = st.form_submit_button("🔄 Guardar Cambios Permanentes")
            with col_b2:
                btn_eliminar = st.form_submit_button("❌ Eliminar Registro Definitivamente")

            if btn_actualizar:
                ejecutar_query("""
                    UPDATE ventas 
                    SET "FECHA"=:fecha, "CEDULA"=:cedula, "NOMBRE"=:nombre, "DIRECCION"=:direccion,
                        "TELEFONO"=:telefono, "TECNICO"=:tecnico, "VALOR"=:valor, "VENDEDOR"=:vendedor,
                        "ESTADO FINAL"=:est_final, "ESTADO PAGO"=:est_pago,
                        "NUMERO DE CUENTA"=:cuenta, "OBSERVACIONES"=:observaciones
                    WHERE "ID"=:id
                """, {
                    "fecha": e_fecha, "cedula": e_cedula, "nombre": e_nombre.upper(),
                    "direccion": e_direccion.upper(), "telefono": e_telefono,
                    "tecnico": e_tecnico.upper(), "valor": float(e_valor),
                    "vendedor": e_vendedor.upper(), "est_final": e_est_final.upper(),
                    "est_pago": e_est_pago.upper(), "cuenta": e_cuenta,
                    "observaciones": e_observaciones, "id": str(id_real),
                })
                st.cache_data.clear()
                verificar_y_ejecutar_respaldo()
                mostrar_mensaje_central("🔄 ¡Registro actualizado correctamente!", "success")
                st.rerun()

            if btn_eliminar:
                confirmar_eliminar(id_real, fila_datos)
