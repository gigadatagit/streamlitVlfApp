import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
import io
import json
import os
from datetime import datetime
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
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
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
        fases = ["A"]
    else:
        fases = []

    # Subida de imágenes de tramos
    st.write("Sube las imágenes de las pruebas de tramos:")
    tramo_imgs = {}
    for i in range(1, cantidad_tramos + 1):
        for f in fases:
            key = f"imgPruebaTramoTrm{i}{f or ''}"
            uploaded_img = st.file_uploader(f"Imagen para Tramo {i} Fase {f}",
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
        
        mapa = StaticMap(600, 400)
        marker = CircleMarker((coordLongitud, coordLatitud,), 'red', 12)  # lon, lat
        mapa.add_marker(marker)
        
        imagenMapa = mapa.render()
        
        buf_mapa = io.BytesIO()
        imagenMapa.save(buf_mapa, format='PNG')
        buf_mapa.seek(0)
        contexto["imgMapsProyecto"] = InlineImage(doc, buf_mapa, Cm(18))

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
        doc.render(contexto)
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button(
            label="Descargar Reporte Word",
            data=output,
            file_name="reporte_vlf.docx",
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
