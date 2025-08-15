import streamlit as st
import json
import io
import os
import math
import matplotlib.pyplot as plt
import geopandas as gpd
from shapely.geometry import Point
import contextily as cx
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from staticmap import StaticMap, CircleMarker

def get_map_png_bytes(lon, lat, buffer_m=300, width_px=900, height_px=700, zoom=17):
    """
    Genera un PNG (bytes) de un mapa satelital con marcador en (lon, lat).
    - buffer_m: radio en metros alrededor del punto (controla "zoom").
    - zoom: nivel de teselas (18-19 suele ser bueno).
    """
    # Crear punto y reproyectar a Web Mercator
    gdf = gpd.GeoDataFrame(geometry=[Point(lon, lat)], crs="EPSG:4326").to_crs(epsg=3857)
    pt = gdf.geometry.iloc[0]
    
    # Calcular bounding box
    bbox = (pt.x - buffer_m, pt.y - buffer_m, pt.x + buffer_m, pt.y + buffer_m)

    # Crear figura
    fig, ax = plt.subplots(figsize=(width_px/100, height_px/100), dpi=100)
    ax.set_xlim(bbox[0], bbox[2])
    ax.set_ylim(bbox[1], bbox[3])

    # Añadir basemap (Esri World Imagery)
    cx.add_basemap(ax, source=cx.providers.Esri.WorldImagery, crs="EPSG:3857", zoom=zoom)

    # Dibujar marcador
    gdf.plot(ax=ax, markersize=40, color="red")

    ax.set_axis_off()
    plt.tight_layout(pad=0)

    # Guardar a buffer en memoria
    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()



def convertir_a_mayusculas(data):
    if isinstance(data, str):
        return data.upper()
    elif isinstance(data, dict):
        return {k: convertir_a_mayusculas(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [convertir_a_mayusculas(v) for v in data]
    elif isinstance(data, tuple):
        return tuple(convertir_a_mayusculas(v) for v in data)
    else:
        return data  # cualquier otro tipo se deja igual


def obtener_valor_por_temperatura(temperatura_prueba: float, tipo_aislamiento: str) -> float:
    """Obtiene el valor de resistencia de aislamiento basado en la temperatura de prueba y tipo de aislamiento.

    Esta función busca el valor de resistencia de aislamiento correspondiente a la temperatura más cercana
    en una lista predefinida, dependiendo del tipo de aislamiento (Aceite o Seco).

    Args:
        temperatura_prueba (float): Temperatura de prueba en grados Celsius.
        tipo_aislamiento (str): Tipo de aislamiento del transformador, puede ser "Aceite" o "Seco".

    Raises:
        ValueError: Si el tipo de aislamiento no es válido.

    Returns:
        float: Valor de resistencia de aislamiento correspondiente a la temperatura más cercana.
    """
    
    # Listado de temperaturas desde -10 hasta 110 con paso de 5
    temperaturas = list(range(-10, 115, 5))
    
    # Valores por tipo de aislamiento
    valores_aceite = [
        0.125, 0.180, 0.25, 0.36, 0.50, 0.75, 1.00, 1.40, 1.98, 2.80, 
        3.95, 5.60, 7.85, 11.20, 15.85, 22.40, 31.75, 44.70, 63.50, 
        89.789, 127.00, 180.00, 254.00, 359.15, 509.00
    ]
    
    valores_seco = [
        0.25, 0.32, 0.40, 0.50, 0.63, 0.81, 1.00, 1.25, 1.58, 2.00, 
        2.50, 3.15, 3.98, 5.00, 6.30, 7.90, 10.00, 12.60, 15.80, 
        20.00, 25.20, 31.60, 40.00, 50.40, 63.20
    ]
    
    # Seleccionar lista de valores según tipo de aislamiento
    if tipo_aislamiento.lower() == "aceite":
        valores = valores_aceite
    elif tipo_aislamiento.lower() == "seco":
        valores = valores_seco
    else:
        raise ValueError("Tipo de aislamiento no válido. Use 'Aceite' o 'Seco'.")
    
    # Encontrar índice de la temperatura más cercana
    indice_cercano = min(range(len(temperaturas)), key=lambda i: abs(temperaturas[i] - temperatura_prueba))
    
    return valores[indice_cercano]


# Inicialización de estado
if 'step' not in st.session_state:
    st.session_state.step = 1
    st.session_state.data = {}

st.title("Formulario Transformadores - Word Automatizado")

# Funciones de navegación
def next_step():
    missing = [k for k, v in st.session_state.data.items() if v is None or v == ""]
    if missing:
        st.error("Por favor completa todos los campos antes de continuar.")
    else:
        st.session_state.step += 1
        st.rerun()

def prev_step():
    if st.session_state.step > 1:
        st.session_state.step -= 1
        st.rerun()

# Paso 1: Información General
if st.session_state.step == 1:
    st.header("Paso 1: Información General")
    
    st.session_state.data['nombreProyecto'] = st.text_input("Nombre del Proyecto", key='nombreProyecto')
    st.session_state.data['nombreCiudadoMunicipio'] = st.text_input("Ciudad o Municipio", key='ciudad')
    st.session_state.data['nombreDepartamento'] = st.text_input("Departamento", key='departamento')
    st.session_state.data['tipoCoordenada'] = st.selectbox(f"Tipo de Imagen para las Coordenadas", ["Urbano", "Rural"], key=f'tipo_coordenada')
    st.session_state.data['nombreCompleto'] = st.text_input("Nombre Completo", key='nombre')
    st.session_state.data['nroConteoTarjeta'] = st.text_input("Número de CONTE o Tarjeta Profesional", key='conte_tarjeta')
    st.session_state.data['nombreCargo'] = st.text_input("Nombre del Cargo", key='cargo')
    st.session_state.data['fechaCreacionSinFormato'] = st.date_input("Fecha de Creación", key='fecha_creacion', value=datetime.now())
    st.session_state.data['fechaCreacion'] = st.session_state.data['fechaCreacionSinFormato'].strftime("%Y-%m-%d")
    st.session_state.data['direccion'] = st.text_input("Dirección", key='direccion')

    cols = st.columns([1,1])
    if cols[1].button("Siguiente"):
        next_step()

# Paso 2: Datos Técnicos
elif st.session_state.step == 2:
    st.header("Paso 2: Datos Técnicos")
    
    st.session_state.data['nroTransformador'] = st.text_input("Número del Transformador", key='nro_transformador')
    st.session_state.data['capacidadTransformador'] = st.text_input("Capacidad del Transformador [kVA]", key='capacidad_transformador')
    st.session_state.data['tipoTransformador'] = st.selectbox("Tipo de Transformador", ["Trifásico", "Monofásico"], key='tipotrafo')
    st.session_state.data['tipoAislamiento'] = st.selectbox("Tipo de Aislamiento", ["Aceite", "Seco"], key='tipoaislamiento')
    st.session_state.data['voltajePrimario'] = st.text_input("Voltaje Primario [V]", key='voltprimario')
    st.session_state.data['voltajeSecundario'] = st.text_input("Voltaje Secundario [V]", key='voltsecundario')
    st.session_state.data['latitud'] = st.text_input("Latitud", key='latitud')
    st.session_state.data['longitud'] = st.text_input("Longitud", key='longitud')
    st.session_state.data['fechaCalibracionSinFormato'] = st.date_input("Fecha de Calibración", key='fecha_calibracion')
    st.session_state.data['fechaCalibracion'] = st.session_state.data['fechaCalibracionSinFormato'].strftime("%Y-%m-%d")

    cols = st.columns([1,1,1])
    if cols[0].button("Anterior"):
        prev_step()
    if cols[1].button("Siguiente"):
        
        if st.session_state.data['tipoTransformador'] == "Trifásico":
            
        
            template_path = 'templates/templateTRAFO3FS.docx'
            
        else:
            
            template_path = 'templates/templateTRAFO1FS.docx'
        
        # Cargar la plantilla en el estado de sesión
        try:
            st.session_state.doc = DocxTemplate(template_path)
            next_step()
        except FileNotFoundError:
            st.error(f"No se encontró la plantilla: {template_path}")

# Paso 3: Formulario de Verificación
elif st.session_state.step == 3:
    st.header("Paso 3: Formulario de Desarrollo y Resultados de la Prueba")
        
    st.session_state.data['carTrafo_Marca'] = st.text_input("Marca del Transformador", key='cartrafomarca')
    st.session_state.data['carTrafo_Serie'] = st.text_input("Serie del Transformador", key='cartrafoserie')
    st.session_state.data['carTrafo_Tipo'] = st.text_input("Tipo del Transformador", key='cartrafotipo')
    st.session_state.data['carTrafo_FechaFabricacionSinFormato'] = st.date_input("Fecha de Fabricación del Transformador", key='cartrafofechafabrisf')
    st.session_state.data['carTrafo_FechaFabricacion'] = st.session_state.data['carTrafo_FechaFabricacionSinFormato'].strftime("%Y-%m-%d")
    st.session_state.data['carTrafo_Frecuencia'] = st.text_input("Frecuencia del Transformador [Hz]", key='cartrafofrec')
    st.session_state.data['carTrafo_NroFases'] = 3 if st.session_state.data['tipoTransformador'] == "Trifásico" else 1
    st.session_state.data['carTrafo_Conexion'] = st.text_input("Conexión del Transformador", key='cartrafoconexion')
    st.session_state.data['carTrafo_MedioAislamiento'] = st.text_input("Medio de Aislamiento del Transformador", key='cartrafomedioaisl')
    st.session_state.data['carTrafo_FechaMediciones'] = st.session_state.data['fechaCreacion']
    st.session_state.data['temperaturaPrueba'] = st.text_input("Temperatura de la Prueba [°C]", key='temperaturaprueba')

    cols = st.columns([1,1,1])
    if cols[0].button("Anterior"):
        prev_step()
    if cols[1].button("Siguiente"):
        next_step()

# Paso 4: Detalles por Tramo
elif st.session_state.step == 4:
    st.header("Paso 4: Detalles de la tabla de Resistencia de Aislamiento")
    
    if st.session_state.data['carTrafo_NroFases'] == 3:
    
        st.session_state.data['resMedida_AVST'] = st.number_input("Resistencia Medida - Alta VS. Tierra [GΩ]", key='res_medida_avst', min_value=0.0, format="%.2f")
        st.session_state.data['resReferida_AVST'] = st.session_state.data['resMedida_AVST']  * obtener_valor_por_temperatura(temperatura_prueba=float(st.session_state.data.get('temperaturaPrueba', 0)), tipo_aislamiento=st.session_state.data.get('tipoAislamiento', 'NA'))
        st.session_state.data['resMedida_AVSB'] = st.number_input("Resistencia Medida - Alta VS. Baja [GΩ]", key='res_medida_avsb', min_value=0.0, format="%.2f")
        st.session_state.data['resReferida_AVSB'] = st.session_state.data['resMedida_AVSB']  * obtener_valor_por_temperatura(temperatura_prueba=float(st.session_state.data.get('temperaturaPrueba', 0)), tipo_aislamiento=st.session_state.data.get('tipoAislamiento', 'NA'))
        st.session_state.data['resMedida_BVST'] = st.number_input("Resistencia Medida - Baja VS. Tierra [GΩ]", key='res_medida_bvst', min_value=0.0, format="%.2f")
        st.session_state.data['resReferida_BVST'] = st.session_state.data['resMedida_BVST']  * obtener_valor_por_temperatura(temperatura_prueba=float(st.session_state.data.get('temperaturaPrueba', 0)), tipo_aislamiento=st.session_state.data.get('tipoAislamiento', 'NA'))
        
        st.session_state.data['resEsp_AVST'] = 5 if st.session_state.data['tipoAislamiento'] == "Aceite" else 25
        st.session_state.data['resEsp_AVSB'] = 5 if st.session_state.data['tipoAislamiento'] == "Aceite" else 25
        st.session_state.data['resEsp_BVST'] = 1 if st.session_state.data['tipoAislamiento'] == "Aceite" else 5
        
        st.session_state.data['resultado_AVST'] = 'Cumple' if st.session_state.data['resReferida_AVST'] >= st.session_state.data['resEsp_AVST'] else 'No Cumple'
        st.session_state.data['resultado_AVSB'] = 'Cumple' if st.session_state.data['resReferida_AVSB'] >= st.session_state.data['resEsp_AVSB'] else 'No Cumple'
        st.session_state.data['resultado_BVST'] = 'Cumple' if st.session_state.data['resReferida_BVST'] >= st.session_state.data['resEsp_BVST'] else 'No Cumple'
        
        st.session_state.data['comentariosPrueba'] = st.text_area("Comentarios de la Prueba", key='comentarios_prueba')
        
    else:
        
        st.session_state.data['resMedida_AVST'] = st.text_input("Resistencia Medida - Alta VS. Tierra [GΩ]", key='res_medida_avst', value='-', disabled=True)
        st.session_state.data['resReferida_AVST'] = '-'
        st.session_state.data['resMedida_AVSB'] = st.number_input("Resistencia Medida - Alta VS. Baja [GΩ]", key='res_medida_avsb', min_value=0.0, format="%.2f")
        st.session_state.data['resReferida_AVSB'] = st.session_state.data['resMedida_AVSB']  * obtener_valor_por_temperatura(temperatura_prueba=float(st.session_state.data.get('temperaturaPrueba', 0)), tipo_aislamiento=st.session_state.data.get('tipoAislamiento', 'NA'))
        st.session_state.data['resMedida_BVST'] = st.number_input("Resistencia Medida - Baja VS. Tierra [GΩ]", key='res_medida_bvst', min_value=0.0, format="%.2f")
        st.session_state.data['resReferida_BVST'] = st.session_state.data['resMedida_BVST']  * obtener_valor_por_temperatura(temperatura_prueba=float(st.session_state.data.get('temperaturaPrueba', 0)), tipo_aislamiento=st.session_state.data.get('tipoAislamiento', 'NA'))
        
        st.session_state.data['resEsp_AVST'] = 5 if st.session_state.data['tipoAislamiento'] == "Aceite" else 25
        st.session_state.data['resEsp_AVSB'] = 5 if st.session_state.data['tipoAislamiento'] == "Aceite" else 25
        st.session_state.data['resEsp_BVST'] = 1 if st.session_state.data['tipoAislamiento'] == "Aceite" else 5
        
        st.session_state.data['resultado_AVST'] = 'Cumple'
        st.session_state.data['resultado_AVSB'] = 'Cumple' if st.session_state.data['resReferida_AVSB'] >= st.session_state.data['resEsp_AVSB'] else 'No Cumple'
        st.session_state.data['resultado_BVST'] = 'Cumple' if st.session_state.data['resReferida_BVST'] >= st.session_state.data['resEsp_BVST'] else 'No Cumple'
        
        st.session_state.data['comentariosPrueba'] = st.text_area("Comentarios de la Prueba", key='comentarios_prueba')

    cols = st.columns([1,1,1])
    if cols[0].button("Anterior"):
        prev_step()
    if cols[1].button("Siguiente"):
        next_step()

# Paso 5: Subida de Imágenes y Generación de Word
elif st.session_state.step == 5:
    
    st.header("Paso 5: Subida de Imágenes de Pruebas y Mapa")
    datos_Sin_Mayuscula = st.session_state.data.copy()
    
    datos = convertir_a_mayusculas(datos_Sin_Mayuscula)


    if st.session_state.data['carTrafo_NroFases'] == 3:

        # Subida de imágenes por tramo
        st.subheader("Imágenes de Pruebas del Transformador")
        
        key_FichaTecTrafo = "imgFichaTecnicaTrafo"
        uploaded_FichaTecTrafo = st.file_uploader(f"Imagen de Ficha Técnica del Trafo", type=['png','jpg','jpeg'], key=key_FichaTecTrafo)
        
        key_ImagenPrueba1 = "imgPruebaMon1"
        uploaded_Prueba1 = st.file_uploader(f"Imagen de Prueba #1 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba1)
        
        key_ImagenPrueba2 = "imgPruebaMon2"
        uploaded_Prueba2 = st.file_uploader(f"Imagen de Prueba #2 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba2)
        
        key_ImagenPrueba3 = "imgPruebaMon3"
        uploaded_Prueba3 = st.file_uploader(f"Imagen de Prueba #3 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba3)
        
        key_ImagenPrueba4 = "imgPruebaMon4"
        uploaded_Prueba4 = st.file_uploader(f"Imagen de Prueba #4 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba4)
        
        
        if uploaded_FichaTecTrafo:
            buf = io.BytesIO(uploaded_FichaTecTrafo.read())
            buf.seek(0)
            datos[key_FichaTecTrafo] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_FichaTecTrafo] = None
            
            
        if uploaded_Prueba1:
            buf = io.BytesIO(uploaded_Prueba1.read())
            buf.seek(0)
            datos[key_ImagenPrueba1] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba1] = None
            
        if uploaded_Prueba2:
            buf = io.BytesIO(uploaded_Prueba2.read())
            buf.seek(0)
            datos[key_ImagenPrueba2] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba2] = None
            
        if uploaded_Prueba3:
            buf = io.BytesIO(uploaded_Prueba3.read())
            buf.seek(0)
            datos[key_ImagenPrueba3] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba3] = None
            
        if uploaded_Prueba4:
            buf = io.BytesIO(uploaded_Prueba4.read())
            buf.seek(0)
            datos[key_ImagenPrueba4] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba4] = None
            
        if st.session_state.data['tipoCoordenada'] == "Urbano":
        
            if st.session_state.data['latitud'] and st.session_state.data['longitud']:
                try:
                    lat = float(str(datos['latitud']).replace(',', '.'))
                    lon = float(str(datos['longitud']).replace(',', '.'))
                    mapa = StaticMap(600, 400)
                    mapa.add_marker(CircleMarker((lon, lat), 'red', 12))
                    img_map = mapa.render()
                    buf_map = io.BytesIO()
                    img_map.save(buf_map, format='PNG')
                    buf_map.seek(0)
                    datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(18))
                except Exception as e:
                    st.error(f"Coordenadas inválidas para el mapa. {e}")
            else:
                st.error("Faltan coordenadas para el mapa.")
                    
        else:
                
            if st.session_state.data['latitud'] and st.session_state.data['longitud']:
                try:
                    lat = float(str(st.session_state.data['latitud']).replace(',', '.'))
                    
                    lon = float(str(st.session_state.data['longitud']).replace(',', '.'))
                    
                    st.warning(f"Prueba de coordenada en modo rural (latitud): {lat}")
                    st.warning(f"Prueba de coordenada en modo rural (longitud): {lon}")
                        
                    png_bytes = get_map_png_bytes(lon, lat, buffer_m=300, zoom=17)
                        
                    buf_map = io.BytesIO(png_bytes)
                    buf_map.seek(0)
                    datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(18))
                except Exception as e:
                    st.error(f"Coordenadas inválidas para el mapa. {e}")
            else:
                st.error("Faltan coordenadas para el mapa.")
            
    else:
        
        # Subida de imágenes por tramo
        st.subheader("Imágenes de Pruebas del Transformador")
        
        key_FichaTecTrafo = "imgFichaTecnicaTrafo"
        uploaded_FichaTecTrafo = st.file_uploader(f"Imagen de Ficha Técnica del Trafo", type=['png','jpg','jpeg'], key=key_FichaTecTrafo)
        
        key_ImagenPrueba1 = "imgPruebaMon1"
        uploaded_Prueba1 = st.file_uploader(f"Imagen de Prueba #1 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba1)
        
        key_ImagenPrueba2 = "imgPruebaMon2"
        uploaded_Prueba2 = st.file_uploader(f"Imagen de Prueba #2 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba2)
        
        key_ImagenPrueba3 = "imgPruebaMon3"
        uploaded_Prueba3 = st.file_uploader(f"Imagen de Prueba #3 del Trafo", type=['png','jpg','jpeg'], key=key_ImagenPrueba3)
        
        
        if uploaded_FichaTecTrafo:
            buf = io.BytesIO(uploaded_FichaTecTrafo.read())
            buf.seek(0)
            datos[key_FichaTecTrafo] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_FichaTecTrafo] = None
            
            
        if uploaded_Prueba1:
            buf = io.BytesIO(uploaded_Prueba1.read())
            buf.seek(0)
            datos[key_ImagenPrueba1] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba1] = None
            
        if uploaded_Prueba2:
            buf = io.BytesIO(uploaded_Prueba2.read())
            buf.seek(0)
            datos[key_ImagenPrueba2] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba2] = None
            
        if uploaded_Prueba3:
            buf = io.BytesIO(uploaded_Prueba3.read())
            buf.seek(0)
            datos[key_ImagenPrueba3] = InlineImage(st.session_state.doc, buf, Cm(14))
        else:
            datos[key_ImagenPrueba3] = None
            
        if st.session_state.data['tipoCoordenada'] == "Urbano":
        
            if st.session_state.data['latitud'] and st.session_state.data['longitud']:
                try:
                    lat = float(str(datos['latitud']).replace(',', '.'))
                    lon = float(str(datos['longitud']).replace(',', '.'))
                    mapa = StaticMap(600, 400)
                    mapa.add_marker(CircleMarker((lon, lat), 'red', 12))
                    img_map = mapa.render()
                    buf_map = io.BytesIO()
                    img_map.save(buf_map, format='PNG')
                    buf_map.seek(0)
                    datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(18))
                except Exception as e:
                    st.error(f"Coordenadas inválidas para el mapa. {e}")
            else:
                st.error("Faltan coordenadas para el mapa.")
                    
        else:
                
            if st.session_state.data['latitud'] and st.session_state.data['longitud']:
                try:
                    lat = float(str(st.session_state.data['latitud']).replace(',', '.'))
                    
                    lon = float(str(st.session_state.data['longitud']).replace(',', '.'))
                    
                    st.warning(f"Prueba de coordenada en modo rural (latitud): {lat}")
                    st.warning(f"Prueba de coordenada en modo rural (longitud): {lon}")
                        
                    png_bytes = get_map_png_bytes(lon, lat, buffer_m=300, zoom=17)
                        
                    buf_map = io.BytesIO(png_bytes)
                    buf_map.seek(0)
                    datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(18))
                except Exception as e:
                    st.error(f"Coordenadas inválidas para el mapa. {e}")
            else:
                st.error("Faltan coordenadas para el mapa.")
    


    if st.button("Generar Word"):
        doc = st.session_state.doc
        # Añadir fecha al contexto
        ahora = datetime.now()
        meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
        datos['dia'] = ahora.day
        datos['mes'] = meses[ahora.month-1]
        datos['anio'] = ahora.year

        doc.render(datos)
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button(
            "Descargar Reporte Word",
            data=output,
            file_name="reporteProtocoloTransformador.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
