import streamlit as st
import pandas as pd
import os
from PIL import Image
import win32net
import io
import base64
import altair as alt
from datetime import datetime, timedelta
from PIL import Image
import pyperclip





# --- CONFIGURACIÓN INICIAL ---
# Conexión al servidor
def conectar_carpeta_servidor():
    network_path = r"\\10.80.24.66\10.80.24.51\Testlog\ALL LIST\0.ENGINEER\AOI TEAM\22.- LIBRARY PN"
    user = r"SFQMM\shopfloor-t"
    password = "Qmmsf#789"
    try:
        
        win32net.NetUseAdd(None, 2, {'remote': network_path, 'username': user, 'password': password}, None)
    except Exception as e:
        st.warning("No se pudo conectar al servidor. Verifica la red.")

# Cargar datos iniciales
@st.cache_data
def cargar_datos():
    try:
       # df_inventario = pd.read_csv("inventario.csv")  # Archivo de Reject Rate
        df_inventario = pd.read_csv(r"\\10.80.24.32\TestLog\0.0 ENGINEERING\0.0 SMT\RejectRate\fujidb.csv")  # Archivo de Reject Rate
        df_componentes = pd.read_excel(r"\\10.80.24.32\TestLog\0.0 ENGINEERING\Componentes\componentes.xlsx")  # Componentes

        # Convertir la columna "DATE" al tipo datetime
        if "DATE" in df_inventario.columns:
            df_inventario["DATE"] = pd.to_datetime(df_inventario["DATE"], errors="coerce")  # Convierte y maneja errores
        
        return df_inventario, df_componentes
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        return None, None

# --- INTERFAZ PRINCIPAL ---
st.title("QMMC Dahsboard")
opcion = st.sidebar.radio("Selecciona una opción:", ["Reject Rate", "Part Number Data","Thermal Profile","Hit Rate","SPI"])
st.sidebar.divider()
# --- REJECT RATE ---
if opcion == "Reject Rate":
    st.subheader("Reject Rate")
    df_inventario, _ = cargar_datos()
    
    if df_inventario is not None:
        # Campos de filtro
        #lineas = st.sidebar.multiselect("Selecciona LINE:", options=df_inventario["LINE"].unique())
        #lados = st.sidebar.multiselect("Selecciona SIDE:", options=df_inventario["SIDE"].unique())
        
        #Opcion personalizada
        lineas = st.sidebar.multiselect("Selecciona LINE:", ["D31", "D21","D11","C41","C31","C21","C11","B61","B51","G21", "A11"])
        lados = st.sidebar.multiselect("Selecciona SIDE:", ["S", "C"])
        # Configurar fechas predeterminadas (hoy y 7 días antes)
        fecha_actual = datetime.now().date()
        fecha_predeterminada_inicio = fecha_actual - timedelta(days=7)
        fecha_predeterminada_fin = fecha_actual

        # Mostrar selector de rango de fechas
        rango_fechas = st.sidebar.date_input(
            "Rango de fechas:",
            value=(fecha_predeterminada_inicio, fecha_predeterminada_fin),  # Valores predeterminados
            min_value=datetime(2022, 1, 1).date(),  # Fecha mínima permitida
            max_value=fecha_actual  # Fecha máxima permitida
        )
        
        try:
            #df_inventario["DATE"] = df_inventario["DATE"].dt.date #convertir a tipo dateime.date
            if rango_fechas and len(rango_fechas) == 2:
                fecha_inicio = pd.to_datetime(rango_fechas[0])
                fecha_fin = pd.to_datetime(rango_fechas[1])
            datos_filtrados = df_inventario[
                (df_inventario["LINE"].isin(lineas)) &
                (df_inventario["SIDE"].isin(lados)) &
                (df_inventario["DATE"].between(fecha_inicio, fecha_fin))
            ]
       
        
        # ----------------------------------- Graficos TOP LINE -> ERROR -----------------------------------#
        
            # Agrupar y ordenar los datos por "LINE" basado en "TOTALERROR"
            lines = ["D31", "D21","D11","C41","C31","C21","C11","B61","B51","G21", "A11"]
            datos_date = df_inventario[df_inventario["DATE"].between(fecha_inicio, fecha_fin)]
            df_filtered = datos_date[datos_date["LINE"].isin(lines)]
            df_grouped = df_filtered.groupby("LINE")["TOTALERROR"].sum().reset_index()
            df_grouped = df_grouped.sort_values(by="TOTALERROR", ascending=False)

            # Crear gráfica con Altair
            chart_line = alt.Chart(df_grouped).mark_bar().encode(
            x=alt.X("TOTALERROR:Q", title="Total Error"),
            y=alt.Y("LINE:N", sort="-x", title="Line"),  # Ordenar por valores descendentes
            tooltip=["LINE", "TOTALERROR"]
            ).properties(
                title="Reject Rate by Line",
                width=700,
                height=400
            )
            # Mostrar la gráfica
            st.altair_chart(chart_line, use_container_width=True)
        

        
        
        
        
        # ----------------------------------- Tabla PN By Line, Side and Date -----------------------------------#     
            # Mostrar tabla
            df_summary = (
            datos_filtrados.groupby("PartNumber")
                .agg(
                    Pickup_Sum=("Pickup", "sum"),
                    Error_Sum=("TOTALERROR", "sum")
                )
                .reset_index()
            )
            
            st.dataframe(df_summary,use_container_width=True)
       
        except NameError as e:
            st.error(f"Seleccione un rango de fechas")



# --- PART NUMBER DATA ---
elif opcion == "Part Number Data":
    st.subheader("Part Number Data")
    _, df_componentes = cargar_datos()

    if df_componentes is not None:
        # Autocompletar entrada de búsqueda
        part_number = st.selectbox("Seleccione o ingrese el número de parte:", 
                                   options=[""]+ list(df_componentes["PART NUMBER"].unique()), 
                                   index=0, 
                                   help="Seleccione un número de parte para obtener información.")

        if part_number:
            resultado = df_componentes[df_componentes["PART NUMBER"] == part_number]
            
            if not resultado.empty:
                categoria = resultado["CATEGORY"].values[0]
                descripcion = resultado["DESCRIPTION"].values[0]
                mfg = resultado["MFG PN"].values[0]
                
                st.write(f"**Categoría:** {categoria}")
                st.write(f"**Descripción:** {descripcion}")
                st.write(f"**MFG PN:** {mfg}")

                # Mostrar imagen
                img_path = os.path.join(r"\\10.80.24.66\10.80.24.51\Testlog\ALL LIST\0.ENGINEER\AOI TEAM\22.- LIBRARY PN\ALL PARTNUMBERS", f"{part_number}.png")
                if os.path.exists(img_path):
                    img = Image.open(img_path)
                    st.image(img, caption=f"Imagen de {part_number}", use_column_width=True)
                else:
                    st.warning("No se encontró imagen para este número de parte.")
                    
                
                
                #-------------------------------- Visualizar PDF --------------------------------------------------------------#
                pdf_path = os.path.join(r"\\10.80.24.32\TestLog\0.0 ENGINEERING\0.0 SMT\Datasheets", f"{part_number}.pdf")

                #Datasheet
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()

                        # Botón para visualizar el PDF
                        #Visualizar datasheet inFrame    
                        base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
                            # Mostrar el PDF incrustado
                        st.markdown(pdf_display, unsafe_allow_html=True)
 
                        #Boton para descargar datasheet
                        st.download_button(
                            label="Descargar Datasheet",
                            data=pdf_bytes,
                            file_name=f"{part_number}.pdf",
                            mime="application/pdf"
                        )
                else:
                    st.warning("No se encontró datasheet para este número de parte.")
                    
            #----------------------------------------------------------------------------------------------
   
                
                
                
            else:
                st.error("Número de parte no encontrado.")


# --- Thermal Profile --- #
elif opcion == "Thermal Profile":
    # Selección de línea y lado
    linea = st.sidebar.selectbox("Selecciona LINE:", ["D31", "D21"])
    #lado = st.sidebar.selectbox("Selecciona SIDE:", ["S", "C"])
    
    # Fecha actual y selector de fecha
    fecha_actual = datetime.now().date()
    fecha_seleccionada = st.sidebar.date_input(
        "Selecciona una fecha:",
        value=fecha_actual,  # Fecha predeterminada
        min_value=datetime(2022, 1, 1).date(),  # Fecha mínima permitida
        max_value=fecha_actual  # Fecha máxima permitida
    )

    # Construcción dinámica de la ruta del PDF según la selección
    anio = fecha_seleccionada.year
    mes = fecha_seleccionada.strftime("%B")  # Nombre del mes en texto (e.g., Enero)
    dia = fecha_seleccionada.strftime("%d")  # Día con dos dígitos

    # Ruta base
    ruta_base = r"\\10.80.24.32\TestLog\0.0 ENGINEERING\0.1 MASS PRODUCTION\OVEN\Thermal Profiles"
    
    # Ruta completa construida dinámicamente
    ruta_pdf = os.path.join(ruta_base, linea, str(anio), mes, dia) #agregar lado voluntario
    
    # Mostrar información al usuario
    #st.write(f"Ruta del archivo seleccionada: `{ruta_pdf}`")
    
    if os.path.exists(ruta_pdf):
        # Listar archivos PDF en el directorio
        archivos_pdf = [f for f in os.listdir(ruta_pdf) if f.endswith('.pdf')]
        
        if archivos_pdf:
            # Crear un selector de archivo
            #archivo_seleccionado = st.selectbox("Selecciona un archivo PDF:", archivos_pdf)
            archivo_seleccionado2 = st.sidebar.selectbox("Selecciona un archivo PDF:", archivos_pdf)
            
            # Ruta completa del archivo seleccionado
            ruta_pdf = os.path.join(ruta_pdf, archivo_seleccionado2)
            
            # Botón para visualizar el PDF
            if st.button("Visualizar Thermal Profile"):
                with open(ruta_pdf, "rb") as pdf_file:
                    pdf_bytes = pdf_file.read()
                    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
            
            # Botón para descargar el PDF
            with open(ruta_pdf, "rb") as pdf_file:
                st.download_button(
                    label="Descargar Thermal Profile",
                    data=pdf_file,
                    file_name=archivo_seleccionado2,
                    mime="application/pdf"
                )
        else:
            st.error("No se encontraron archivos PDF en la ruta seleccionada.")
    else:
        st.error("No esta presente el archivo")
        
        
# --- NUEVA OPCIÓN: HIT RATE --- #
elif opcion == "Hit Rate":
    st.subheader("Hit Rate - Producción vs Target")

    # Cargar archivo desde ruta especificada
    ruta_archivo = r"\\10.80.24.32\TestLog\0.0 ENGINEERING\0.0 SMT\HitRate\InputOutputWIPDailyReport_2025-02-07.xlsx"

    #try:
        # Leer hojas disponibles en el archivo Excel
    hojas_disponibles = pd.ExcelFile(ruta_archivo).sheet_names

        # Selector para elegir la línea (representada por el nombre de la hoja)
    hoja_seleccionada = st.selectbox("Selecciona la línea y modelo:", hojas_disponibles)

    if hoja_seleccionada:
            # Leer datos de la hoja seleccionada
            df_hit_rate = pd.read_excel(ruta_archivo, sheet_name=hoja_seleccionada, skiprows=2)

            # Limpiar datos eliminando filas de totales y filas vacías
            df_hit_rate = df_hit_rate[df_hit_rate["SSP Target"].notna() & df_hit_rate["SSP ACTUAL"].notna()]
            df_hit_rate = df_hit_rate[~df_hit_rate["2025-02-07"].str.contains("Total", na=False)]

            # Renombrar columnas para mejor manejo
            df_hit_rate.rename(columns={
                "2025-02-07": "Hora",
                "SSP Target": "Target",
                "SSP ACTUAL": "Actual"
            }, inplace=True)

            # Convertir "Hora" en categoría para orden correcto
            df_hit_rate["Hora"] = pd.Categorical(df_hit_rate["Hora"], categories=df_hit_rate["Hora"].unique(), ordered=True)

            # Graficar los datos
            chart = alt.Chart(df_hit_rate).mark_bar(color="steelblue").encode(
                x=alt.X("Hora:N", title="Hora"),
                y=alt.Y("Actual:Q", title="Producción Real"),
                tooltip=["Hora", "Actual", "Target"]
            ).properties(
                title=f"Producción vs Target - {hoja_seleccionada}",
                width=800,
                height=400
            )

            # Agregar línea para el target
            line = alt.Chart(df_hit_rate).mark_line(color="red").encode(
                x="Hora:N",
                y="Target:Q"
            )

            # Combinar barra y línea
            st.altair_chart(chart + line, use_container_width=True)
        

            # Mostrar tabla de datos
            st.dataframe(df_hit_rate, use_container_width=True)
            
            #Grafica de pastel
            # Calcular el promedio de SSP Hit Rate
            promedio_hit_rate = df_hit_rate["SSP Hit rate"].mean()
            promedio_hit_rate = promedio_hit_rate * 100

            # Preparar datos para gráfica de pastel
            df_pie = pd.DataFrame({
                "Categoría": ["Average Hit rate", "Pending"],
                "Porcentaje": [promedio_hit_rate, 100 - promedio_hit_rate]
            })

            # Graficar pastel
            pie_chart = alt.Chart(df_pie).mark_arc().encode(
                theta=alt.Theta(field="Porcentaje", type="quantitative"),
                color=alt.Color(field="Categoría", type="nominal"),
                tooltip=["Categoría", "Porcentaje"]
            ).properties(
                title="Promedio de SSP Hit Rate",
                width=400,
                height=400
            )

            st.altair_chart(pie_chart, use_container_width=True)
            pie_chart.mark_text(align='left', dx=2)
            
            
            

    #except Exception as e:
        #st.error(f"Error al cargar o procesar los datos: {e}")
        
        
# --- SPI SECTION --- #
elif opcion == "SPI":
    # Selección de línea y lado
    linea = st.sidebar.selectbox("Selecciona LINE:", ["D31", "D21","D11","C41","C31","C21","C11","B61","B51","A11"])
    lado = st.sidebar.selectbox("Selecciona SIDE:", ["S", "C"])
    opcion_valor = st.sidebar.selectbox("Selecciona opcion:", ["IMAGE"])
    
    # Fecha actual y selector de fecha  
    fecha_actual = datetime.now().date()
    fecha_seleccionada = st.sidebar.date_input(
        "Selecciona una fecha:",
        value=fecha_actual,  # Fecha predeterminada
        min_value=datetime(2022, 1, 1).date(),  # Fecha mínima permitida
        max_value=fecha_actual  # Fecha máxima permitida
    )
    
    
    # Construcción dinámica de la ruta de la imagen según la selección
    anio = fecha_seleccionada.year
    mes = fecha_seleccionada.strftime("%m")  # Nombre del mes en texto (e.g., Enero) %B es con texto
    dia = fecha_seleccionada.strftime("%d")  # Día con dos dígitos

    # Ruta base
    ruta_base = r"\\10.80.24.65\10.80.24.51\Testlog"
    
    # Ruta completa construida dinámicamente
    
    if lado == "S":
        ruta_lado = "SS"
    else:
        ruta_lado = "CS"
    
    if (opcion_valor == "IMAGE"):
        ruta_lado = ruta_lado + "_IMAGE"
        ruta_imagen = os.path.join(ruta_base, linea, "SPI", ruta_lado, str(anio), str(mes), dia) #agregar lado voluntario
    else:
         ruta_reporte = 0
    
    
    # Mostrar información al usuario
    #st.write(f"Ruta del archivo seleccionada: `{ruta_pdf}`")
    # Aumentar el límite de tamaño permitido
    Image.MAX_IMAGE_PIXELS = None
    if os.path.exists(ruta_imagen):
        # Listar archivos PDF en el directorio
        archivos_imagen = [f for f in os.listdir(ruta_imagen) if f.endswith('.jpg')]
        
        if archivos_imagen:
            # Crear un selector de archivo
            #archivo_seleccionado = st.selectbox("Selecciona un archivo PDF:", archivos_pdf)
            archivo_seleccionado2 = st.sidebar.selectbox("Selecciona un archivo PDF:", archivos_imagen)
            
            # Ruta completa del archivo seleccionado
            ruta_imagen = os.path.join(ruta_imagen, archivo_seleccionado2)
            
            
            # Mostrar imagen
            #img_path = os.path.join(r"\\10.80.24.66\10.80.24.51\Testlog\ALL LIST\0.ENGINEER\AOI TEAM\22.- LIBRARY PN\ALL PARTNUMBERS", f"{part_number}.png")
            if st.sidebar.button("Check"):
                if os.path.exists(ruta_imagen):
                    img = Image.open(ruta_imagen)
                    #img.thumbnail((1920, 1080))
                    st.image(img, caption=f"Imagen de {ruta_imagen}", use_column_width=True)
                else:
                    st.warning("No se encontró imagen para este número de parte.")
        else:
            st.error("No se encontraron archivos PDF en la ruta seleccionada.")
    else:
        st.error("No esta presente el archivo")

elif opcion == "Search SN:":
    # Selector para elegir la línea
    linea_seleccionada = st.selectbox(
        "Selecciona la línea:",
        ["D31", "D21", "D11", "C41", "C31", "C21", "C11", "B61", "B51", "G21", "A11"]
    )

    # Selector para elegir la estación
    estacion_seleccionada = st.selectbox(
        "Selecciona la estación:",
        ["AOI", "AXI", "SPI"]
    )

    # Input para ingresar el valor a buscar
    valor_buscar = st.text_input("Ingresa el valor a buscar:", placeholder="Search Serial Number")

    # Botón para iniciar la búsqueda
    if st.button("Buscar archivo"):
        if linea_seleccionada and estacion_seleccionada and valor_buscar:
            try:
                # Extraer la terminación del valor ingresado (después del último ":")
                terminacion = valor_buscar.split(":")[-1]

                # Construir la ruta base según la línea y la estación seleccionadas
                ruta_base = os.path.join(r"\\10.80.24.65\10.80.24.51\Testlog", linea_seleccionada, estacion_seleccionada)

                # Lista para almacenar los resultados encontrados
                resultados = []

                # Buscar en la estructura de directorios de la ruta seleccionada
                for root, dirs, files in os.walk(ruta_base):
                    for file in files:
                        if terminacion in file:
                            resultados.append(os.path.join(root, file))

                # Mostrar los resultados
                if resultados:
                    st.success(f"Se encontraron {len(resultados)} resultado(s):")
                    for resultado in resultados:
                        st.write(resultado)
                else:
                    st.error("No se encontraron archivos con la terminación proporcionada.")
            except Exception as e:
                st.error(f"Ocurrió un error: {e}")
        else:
            st.warning("Por favor, selecciona una línea, una estación e ingresa un valor para buscar.")


# --- FINAL ---
st.sidebar.divider()
st.sidebar.info("Developed by: M2302209")


