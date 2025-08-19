import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
import io
import json
import os
import math
import matplotlib.pyplot as plt
import geopandas as gpd
from datetime import datetime
from shapely.geometry import Point
import contextily as cx
from staticmap import StaticMap, CircleMarker



def obtener_template_path(tipo_tramo: str, cantidad_tramos: int) -> str:
    """
    Retorna el path del template a usar basado en el tipo de tramo y la cantidad.
    Ejemplo: 'Trifásicos' y 3 → 'templateVLF3FS3TR.docx'
             'Monofásicos' y 10 → 'templateVLF1FS10TR.docx'
    """
    fases = "3FS" if tipo_tramo == "Trifásicos" else "1FS"
    nombre_template = f"templateVLF{fases}{cantidad_tramos}TR.docx"
    return os.path.join('templates', nombre_template)

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

def convertir_valores_a_mayusculas(json_data):
    if isinstance(json_data, dict):
        return {clave: convertir_valores_a_mayusculas(valor) for clave, valor in json_data.items()}
    elif isinstance(json_data, list):
        return [convertir_valores_a_mayusculas(elemento) for elemento in json_data]
    elif isinstance(json_data, tuple):
        return tuple(convertir_valores_a_mayusculas(elemento) for elemento in json_data)
    elif isinstance(json_data, str):
        return json_data.upper()
    else:
        return json_data

def pagina_generacion_word():
    st.title("Generación de Word Automatizado - Pruebas VLF")
    st.write("Sube tu archivo JSON con los datos para el reporte:")
    uploaded_json = st.file_uploader("Archivo JSON", type=["json"])

    if not uploaded_json:
        return

    datos = json.load(uploaded_json)
    
    tension_prueba = datos.get("tensionPrueba", "")
    if tension_prueba == "Aceptación":
        datos["valTensionPrueba"] = 21
    elif tension_prueba == "Mantenimiento":
        datos["valTensionPrueba"] = 16
        
    coordLatitud = datos.get("latitud", "")
    coordLongitud = datos.get("longitud", "")
    if coordLatitud and coordLongitud:
        try:
            coordLatitud = float(coordLatitud)
            coordLongitud = float(coordLongitud)
        except ValueError:
            st.error("Las coordenadas deben ser números válidos.")
            return
    else:
        st.error("Por favor, proporciona las coordenadas de latitud y longitud.")
        return

        
    fecha_actual = datetime.now()
    datos["dia"] = fecha_actual.day
    meses_es = [
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ]
    datos["mes"] = meses_es[fecha_actual.month - 1]
    datos["anio"] = fecha_actual.year

    # Selección de imagen según tensionPrueba
    tension = datos.get("tensionPrueba", "")
    if tension == "Aceptación":
        img_path = "images/imgAceptacion.png"
    elif tension == "Mantenimiento":
        img_path = "images/imgMantenimiento.png"
    else:
        img_path = None

    # Determinar fases según tipo de tramos
    tipo_tramos = datos.get("tipoTramos", "")
    cantidad_tramos = int(datos.get("cantidadTramos", 0))
    if tipo_tramos == "Trifásicos":
        fases = ["A", "B", "C"]
    elif tipo_tramos == "Bifásicos":
        fases = ["A", "B"]
    elif tipo_tramos == "Monofásicos":
        fases = [""]
    else:
        fases = [""]
        
    st.write("Selecciona el tipo de coordenada:")
    
    tipo_coordenada = st.selectbox("Tipo de coordenada", ["Urbano", "Rural"], key="tipoCoordenada")
    
    st.session_state.data = datos
    st.session_state.data['tipoCoordenada'] = tipo_coordenada

    # Subida de imágenes de tramos
    st.write("Sube las imágenes de las pruebas de tramos:")
    tramo_imgs = {}
    for i in range(1, cantidad_tramos + 1):
        for f in fases:
            key = f"imgPruebaTramoTrm{i}{f or ''}"
            uploaded_img = st.file_uploader(f"Imagen para Tramo {i} Fase {f or 'Única'}",
                                            type=["png", "jpg", "jpeg"], key=key)
            if uploaded_img:
                buf = io.BytesIO(uploaded_img.read())
                buf.seek(0)
                tramo_imgs[key] = buf
            else:
                tramo_imgs[key] = None

    if st.button("Generar Word"):
        
        try:
            
            template_path = obtener_template_path(tipo_tramos, cantidad_tramos)
            
        except FileNotFoundError:
            
            st.error(f"No se encontró la plantilla: {template_path}")
        
        
        # 1) Creamos el DocxTemplate una sola vez
        doc = DocxTemplate(template_path)
        contexto = datos.copy()
        
        if st.session_state.data['tipoCoordenada'] == "Urbano":
        
            if contexto['latitud'] and contexto['longitud']:
                try:
                    lat = float(str(contexto['latitud']).replace(',', '.'))
                    lon = float(str(contexto['longitud']).replace(',', '.'))
                    mapa = StaticMap(600, 400)
                    mapa.add_marker(CircleMarker((lon, lat), 'red', 12))
                    img_map = mapa.render()
                    buf_map = io.BytesIO()
                    img_map.save(buf_map, format='PNG')
                    buf_map.seek(0)
                    contexto['imgMapsProyecto'] = InlineImage(doc, buf_map, Cm(18))
                except Exception as e:
                    st.error(f"Coordenadas inválidas para el mapa. {e}")
            else:
                st.error("Faltan coordenadas para el mapa.")
                    
        else:
                
            if contexto['latitud'] and contexto['longitud']:
                try:
                    lat = float(str(contexto['latitud']).replace(',', '.'))
                    
                    lon = float(str(contexto['longitud']).replace(',', '.'))
                    
                    st.warning(f"Prueba de coordenada en modo rural (latitud): {lat}")
                    st.warning(f"Prueba de coordenada en modo rural (longitud): {lon}")
                        
                    png_bytes = get_map_png_bytes(lon, lat, buffer_m=300, zoom=17)
                        
                    buf_map = io.BytesIO(png_bytes)
                    buf_map.seek(0)
                    contexto['imgMapsProyecto'] = InlineImage(doc, buf_map, Cm(18))
                except Exception as e:
                    st.error(f"Coordenadas inválidas para el mapa. {e}")
            else:
                st.error("Faltan coordenadas para el mapa.")
        
        #mapa = StaticMap(600, 400)
        #marker = CircleMarker((coordLongitud, coordLatitud,), 'red', 12)  # lon, lat
        #mapa.add_marker(marker)
        
        #imagenMapa = mapa.render()
        
        #buf_mapa = io.BytesIO()
        #imagenMapa.save(buf_mapa, format='PNG')
        #buf_mapa.seek(0)
        #contexto["imgMapsProyecto"] = InlineImage(doc, buf_mapa, Cm(18))

        # 2) Tabla de tensión
        if img_path and os.path.exists(img_path):
            with open(img_path, "rb") as f:
                buf_tension = io.BytesIO(f.read())
            buf_tension.seek(0)
            contexto["imgTablaTensionPrueba"] = InlineImage(doc, buf_tension, Cm(18))

        # 3) Imágenes de cada tramo
        for key, buf in tramo_imgs.items():
            if buf:
                contexto[key] = InlineImage(doc, buf, Cm(14))
            else:
                contexto[key] = ""

        # 4) Renderizar y ofrecer descarga
        contexto_Mayuscula = convertir_valores_a_mayusculas(contexto)
        doc.render(contexto_Mayuscula)
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button(
            label="Descargar Reporte Word",
            data=output,
            file_name="reporteProtocoloVLF.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

def main():
    st.sidebar.title("Menú de Páginas - Aplicación Web")
    pagina_seleccionada = st.sidebar.radio(
        "Selecciona una página",
        ["Generación de Word Automatizado"]
    )
    if pagina_seleccionada == "Generación de Word Automatizado":
        pagina_generacion_word()

if __name__ == "__main__":
    main()
