"""
================================================================================
SISTEMA INSTITUCIONAL DE GENERACIÓN DE INFORMES TÉCNICOS DE MANTENIMIENTO
Versión: 3.0 (Optimizado para Personal Técnico - Estándar APA 7.ª Edición)
Desarrollado con Python 3.11+, Streamlit, Google GenAI, python-docx y ReportLab.
================================================================================
"""

import io
import os
import re
from datetime import datetime, date
import streamlit as st
from PIL import Image
import pandas as pd
from google import genai

# Librerías profesionales para generación de documentos Word (.docx)
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Librerías profesionales para generación de documentos PDF (ReportLab)
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas

# ==============================================================================
# 1. CONFIGURACIÓN INICIAL Y DE PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Generador de Informes Técnicos de Mantenimiento",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Profesionales para la Interfaz Web
st.markdown("""
    <style>
    .main {
        background-color: #F8FAFC;
    }
    .stAppHeader {
        background-color: rgba(255, 255, 255, 0.95);
    }
    .main-header {
        font-size: 2.1rem;
        color: #0F172A;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.0rem;
        color: #475569;
        margin-bottom: 1.5rem;
    }
    .card-container {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);
        border: 1px solid #E2E8F0;
        margin-bottom: 1.2rem;
    }
    .stButton>button {
        border-radius: 0.5rem;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. GESTIÓN DE CLIENTE GEMINI API
# ==============================================================================
def get_gemini_client():
    """Inicializa y valida el cliente de la API de Google GenAI de forma segura."""
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "gemini_api_key_input" in st.session_state and st.session_state["gemini_api_key_input"]:
        api_key = st.session_state["gemini_api_key_input"]
    
    if not api_key:
        return None
    return genai.Client(api_key=api_key)

# ==============================================================================
# 3. BARRA LATERAL SIMPLIFICADA PARA PERSONAL TÉCNICO
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=65)
    st.markdown("## **Mantenimiento Industrial**")
    st.markdown("*Módulo de Reportes Técnicos*")
    st.markdown("---")
    
    # Autenticación de API Key si no está en secrets.toml
    if "GEMINI_API_KEY" not in st.secrets:
        st.subheader("🔑 Credenciales")
        api_key_in = st.text_input("Gemini API Key:", type="password", key="gemini_api_key_input")
        if not api_key_in:
            st.info("💡 Ingrese su clave API para activar el asistente.")
        st.markdown("---")

    st.subheader("📋 Datos Generales")
    
    empresa_nombre = st.text_input("Empresa / Planta", "ORICA")
    unidad_operativa = st.text_input("Unidad / Área", "Unidad de Detonadores")
    
    tipo_informe = st.selectbox(
        "Tipo de Informe",
        [
            "Informe Ejecutivo de Mantenimiento",
            "Informe de Falla",
            "Informe de Investigación",
            "Informe de Mantenimiento Correctivo",
            "Informe de Mantenimiento Preventivo",
            "Informe de Incidente Operacional"
        ]
    )

    st.markdown("---")
    st.subheader("👤 Responsable")
    elaborado_por = st.text_input("Elaborado por", "Ing. Roger Huamani Torres")
    cargo_elaborador = st.text_input("Cargo", "Superintendente de Mantenimiento")
    fecha_evento = st.date_input("Fecha del Evento", date.today())

# ==============================================================================
# 4. INTERFAZ PRINCIPAL STREAMLIT
# ==============================================================================
st.markdown('<div class="main-header">Generador Inteligente de Informes de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Plataforma técnica simplificada para el registro rápido de eventos, redacción estructurada por IA y exportación corporativa bajo norma APA.</div>', unsafe_allow_html=True)

# Pestañas de navegación directas para el técnico
tab_datos, tab_evidencias, tab_generacion = st.tabs([
    "📝 1. Datos del Activo & Evento", 
    "📷 2. Evidencias Fotográficas", 
    "🚀 3. Generación & Exportación"
])

# ------------------------------------------------------------------------------
# PESTAÑA 1: DATOS DEL ACTIVO Y EVENTO
# ------------------------------------------------------------------------------
with tab_datos:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### Identificación del Activo")
    col_e1, col_e2, col_e3 = st.columns(3)
    with col_e1:
        equipo_tag = st.text_input("Tag del Equipo", "230-1", help="Identificador único del activo.")
    with col_e2:
        equipo_nombre = st.text_input("Nombre del Equipo", "Carga de Retardos (230-1)")
    with col_e3:
        horas_operacion = st.number_input("Horas de Operación (TTSO)", min_value=0, value=14200, step=100)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### Contexto y Descripción del Evento")
    
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        contexto_evento = st.text_area(
            "Contexto del Evento",
            placeholder="Describa el régimen de trabajo, condiciones iniciales o proceso en curso...",
            height=130
        )
        descripcion_incidente = st.text_area(
            "Descripción Inicial del Incidente",
            placeholder="Detalle la manifestación inicial, anomalías detectadas, ruidos o alarmas...",
            height=130
        )
        hallazgos_tecnicos = st.text_area(
            "Hallazgos Encontrados",
            placeholder="Ingrese los hallazgos físicos visuales, desgaste, fugas, daños observados...",
            height=130
        )
    
    with col_t2:
        trabajos_ejecutados = st.text_area(
            "Trabajos Realizados",
            placeholder="Enumere las tareas correctivas, reemplazos de piezas, ajustes o pruebas...",
            height=130
        )
        repuestos_utilizados = st.text_area(
            "Repuestos / Materiales Empleados",
            placeholder="Ej: O-rings de repuesto, sellos, lubricantes, componentes nuevos...",
            height=130
        )
        info_adicional = st.text_area(
            "Observaciones Adicionales (Opcional)",
            placeholder="Cualquier información técnica complementaria relevante...",
            height=130
        )
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 2: EVIDENCIAS FOTOGRÁFICAS (MÚLTIPLES IMÁGENES LIBRES)
# ------------------------------------------------------------------------------
with tab_evidencias:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### Carga de Fotografías y Descripciones")
    st.info("💡 Puede cargar todas las imágenes que necesite. Cada fotografía requiere obligatoriamente su descripción técnica para cumplir con el formato APA.")
    
    uploaded_files = st.file_uploader(
        "Seleccione las imágenes de evidencia (Formatos: JPG, JPEG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    
    image_data_list = []
    if uploaded_files:
        st.success(f"Se han cargado {len(uploaded_files)} imágenes para el registro fotográfico.")
        st.markdown("---")
        
        for idx, file_item in enumerate(uploaded_files):
            col_img, col_desc = st.columns([1, 2])
            with col_img:
                img_obj = Image.open(file_item)
                st.image(img_obj, caption=f"Imagen Original {idx+1}", use_container_width=True)
            with col_desc:
                desc_input = st.text_input(
                    f"Descripción obligatoria para Figura {idx+1}:",
                    key=f"img_desc_key_{idx}",
                    placeholder="Ej: Vista general de la estación evidenciando daños y residuos..."
                )
                image_data_list.append({
                    "file_name": file_item.name,
                    "image_obj": img_obj,
                    "description": desc_input,
                    "index": idx + 1
                })
    else:
        st.warning("No se han cargado imágenes todavía. Es recomendable adjuntar al menos una evidencia fotográfica.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 3: GENERACIÓN Y EXPORTACIÓN AUTOMÁTICA
# ------------------------------------------------------------------------------
with tab_generacion:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### Consolidación y Aprobación del Informe")
    st.markdown("El sistema consolidará la información en un formato resumido, técnico, profesional y directo para aprobación inmediata.")
    
    has_images = len(image_data_list) > 0
    missing_desc = any(not img["description"].strip() for img in image_data_list) if has_images else False
    
    if has_images and missing_desc:
        st.error("⚠️ Atención: Hay imágenes cargadas sin descripción asignada. Por favor complete todas las descripciones.")
    else:
        st.success("✅ Datos listos para la generación automática por Inteligencia Artificial.")
    
    btn_generar = st.button("🚀 Generar Informe Técnico", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_generar and not missing_desc:
        client = get_gemini_client()
        if not client:
            st.error("❌ No se detectó la API Key de Gemini. Configúrela en los secretos o en la barra lateral.")
        else:
            with st.spinner("Generando informe técnico resumido, estructurado y con formato profesional..."):
                
                evidencias_prompt_str = "\n".join([f"- Figura {item['index']}: {item['description']}" for item in image_data_list]) if image_data_list else "Sin imágenes adjuntas."
                
                # Prompt optimizado para respuestas concisas, técnicas y estructuradas con viñetas limpias
                prompt_master = f"""
Actúa como un Ingeniero Senior de Mantenimiento con amplia experiencia en plantas industriales. 
Genera un Informe Técnico de Mantenimiento corporativo basado estrictamente en la información suministrada, manteniendo un tono técnico, conciso, directo y profesional (evitando explicaciones exageradas o redundantes).

Parámetros:
- Empresa: {empresa_nombre}
- Unidad Operativa: {unidad_operativa}
- Tipo de Informe: {tipo_informe}
- Equipo: {equipo_tag} - {equipo_nombre}
- Horas de Operación: {horas_operacion} hrs
- Elaborado por: {elaborado_por} ({cargo_elaborador})

Datos del Evento:
- Contexto: {contexto_evento}
- Descripción del Incidente: {descripcion_incidente}
- Hallazgos: {hallazgos_tecnicos}
- Trabajos Realizados: {trabajos_ejecutados}
- Repuestos: {repuestos_utilizados}
- Información Adicional: {info_adicional}

Evidencias Fotográficas Registradas:
{evidencias_prompt_str}

ESTRUCTURA OBLIGATORIA (Utiliza exactamente estos 19 encabezados numerados en formato Markdown con '##'):
## 1. Resumen Ejecutivo
## 2. Información General
## 3. Antecedentes
## 4. Descripción del Evento o Problema Identificado
## 5. Inspección Realizada
## 6. Hallazgos
## 7. Observaciones Técnicas
## 8. Trabajos Ejecutados
## 9. Evaluación Mecánica
## 10. Evaluación Eléctrica
## 11. Evaluación de Riesgos
## 12. Análisis de Causa Probable
## 13. Acciones Correctivas
## 14. Acciones Preventivas
## 15. Conclusiones
## 16. Recomendaciones
## 17. Estado Final del Equipo
## 18. Evidencias Fotográficas
## 19. Referencias

REQUISITO IMPORTANTE SOBRE VIÑETAS: Cuando utilices listas o enumeraciones de puntos clave dentro de las secciones, utiliza viñetas con asteriscos o guiones (ej. * Elemento o - Elemento) para que el sistema las preserve adecuadamente. Sé directo y claro.
"""
                try:
                    res_genai = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt_master
                    )
                    st.session_state["informe_generado"] = res_genai.text
                    st.session_state["datos_informe"] = {
                        "empresa": empresa_nombre,
                        "unidad": unidad_operativa,
                        "tipo": tipo_informe,
                        "tag": equipo_tag,
                        "equipo": equipo_nombre,
                        "horas": horas_operacion,
                        "elaborador": elaborado_por,
                        "cargo": cargo_elaborador,
                        "fecha": str(fecha_evento),
                        "imagenes": image_data_list
                    }
                    st.success("¡Informe generado con éxito! Revise la vista previa y proceda a descargar.")
                except Exception as ex:
                    st.error(f"Error al conectar con la API de Gemini: {ex}")

# ==============================================================================
# 5. VISUALIZACIÓN Y EXPORTACIÓN PROFESIONAL (WORD Y PDF)
# ==============================================================================
if "informe_generado" in st.session_state:
    st.markdown("---")
    st.markdown("## 📄 Vista Previa del Informe Técnico")
    
    informe_md = st.session_state["informe_generado"]
    meta_inf = st.session_state["datos_informe"]
    
    with st.expander("Expandir para visualizar el contenido completo", expanded=True):
        st.markdown(informe_md)
        if meta_inf["imagenes"]:
            st.markdown("### Registro Fotográfico Asociado")
            for item in meta_inf["imagenes"]:
                st.image(
                    item["image_obj"], 
                    caption=f"Figura {item['index']}. {item['description']} (Fuente: Elaboración propia)", 
                    use_container_width=True
                )

    st.markdown("---")
    st.markdown("### 📥 Opciones de Exportación")
    
    col_w, col_p = st.columns(2)
    
    # --------------------------------------------------------------------------
    # EXPORTACIÓN A WORD (.DOCX) - JUSTIFICADO, INTERLINEADO Y VIÑETAS
    # --------------------------------------------------------------------------
    with col_w:
        if st.button("📥 Descargar Informe en Word (.docx)", use_container_width=True):
            doc = Document()
            
            # Márgenes de 1 pulgada
            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # Portada de encabezado
            p_head = doc.add_paragraph()
            p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_head = p_head.add_run(f"{meta_inf['empresa']} | {meta_inf['unidad']}")
            r_head.font.size = Pt(9)
            r_head.font.color.rgb = RGBColor(100, 116, 139)
            
            doc.add_paragraph()
            
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(f"INFORME TÉCNICO DE MANTENIMIENTO\n{meta_inf['tipo'].upper()}")
            r_title.font.size = Pt(16)
            r_title.font.bold = True
            r_title.font.color.rgb = RGBColor(15, 23, 42)
            
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sub = p_sub.add_run(f"ACTIVO: {meta_inf['tag']} - {meta_inf['equipo']}")
            r_sub.font.size = Pt(11)
            r_sub.font.bold = True
            r_sub.font.color.rgb = RGBColor(29, 78, 216)
            
            doc.add_paragraph()
            
            # Tabla de metadatos en Word
            table_meta = doc.add_table(rows=4, cols=2)
            table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
            meta_rows = [
                ("Fecha de Emisión:", meta_inf["fecha"]),
                ("Horas de Operación:", f"{meta_inf['horas']} hrs"),
                ("Elaborado por:", f"{meta_inf['elaborador']} ({meta_inf['cargo']})"),
                ("Estándar Normativo:", "APA 7.ª Edición / Buenas Prácticas Industriales")
            ]
            for i, (label, val) in enumerate(meta_rows):
                cells = table_meta.rows[i].cells
                cells[0].text = label
                cells[0].paragraphs[0].runs[0].font.bold = True
                cells[0].paragraphs[0].runs[0].font.size = Pt(10)
                cells[1].text = val
                cells[1].paragraphs[0].runs[0].font.size = Pt(10)
            
            doc.add_page_break()
            
            # Procesamiento de líneas del informe para Word (con soporte de viñetas y justificación)
            for line in informe_md.split('\n'):
                line_str = line.strip()
                if line_str.startswith("## "):
                    p_h = doc.add_heading(line_str.replace("## ", ""), level=2)
                    p_h.runs[0].font.color.rgb = RGBColor(15, 23, 42)
                    p_h.paragraph_format.space_before = Pt(12)
                    p_h.paragraph_format.space_after = Pt(4)
                elif line_str.startswith("# "):
                    p_h = doc.add_heading(line_str.replace("# ", ""), level=1)
                    p_h.runs[0].font.color.rgb = RGBColor(30, 58, 138)
                elif line_str.startswith("* ") or line_str.startswith("- "):
                    # Párrafo con viñeta conservada
                    p_bullet = doc.add_paragraph(style='List Bullet')
                    run_b = p_bullet.add_run(line_str[2:].strip())
                    run_b.font.size = Pt(10.5)
                    run_b.font.name = 'Calibri'
                    p_bullet.paragraph_format.line_spacing = 1.15
                    p_bullet.paragraph_format.space_after = Pt(3)
                    p_bullet.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                elif line_str:
                    p = doc.add_paragraph()
                    run = p.add_run(line_str)
                    run.font.size = Pt(10.5)
                    run.font.name = 'Calibri'
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(6)
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Anexo Fotográfico en Word (Centrado con formato APA)
            if meta_inf["imagenes"]:
                doc.add_page_break()
                p_anexo = doc.add_heading("Evidencias Fotográficas", level=2)
                p_anexo.runs[0].font.color.rgb = RGBColor(15, 23, 42)
                
                for item in meta_inf["imagenes"]:
                    img_mem = io.BytesIO()
                    item["image_obj"].save(img_mem, format='PNG')
                    img_mem.seek(0)
                    
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(img_mem, width=Inches(4.5))
                    
                    p_cap = doc.add_paragraph()
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_c1 = p_cap.add_run(f"Figura {item['index']}. ")
                    r_c1.font.bold = True
                    r_c1.font.size = Pt(9.5)
                    r_c2 = p_cap.add_run(f"{item['description']}\n")
                    r_c2.font.italic = True
                    r_c2.font.size = Pt(9.5)
                    r_c3 = p_cap.add_run("Fuente: Elaboración propia.")
                    r_c3.font.size = Pt(9.5)
                    doc.add_paragraph()

            # Guardar archivo Word en búfer
            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .docx",
                data=docx_buffer,
                file_name=f"Informe_Tecnico_{meta_inf['tag']}_{meta_inf['fecha']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # --------------------------------------------------------------------------
    # EXPORTACIÓN A PDF (REPORTLAB - JUSTIFICADO, CENTRADO, VIÑETAS Y APA)
    # --------------------------------------------------------------------------
    with col_p:
        if st.button("📥 Descargar Informe en PDF (.pdf)", use_container_width=True):
            pdf_buffer = io.BytesIO()
            doc_pdf = SimpleDocTemplate(
                pdf_buffer,
                pagesize=letter,
                rightMargin=54,
                leftMargin=54,
                topMargin=54,
                bottomMargin=54
            )
            
            styles = getSampleStyleSheet()
            
            # Estilos personalizados para PDF con justificación e interlineado corporativo
            style_title = ParagraphStyle(
                'DocTitle',
                parent=styles['Heading1'],
                fontSize=18,
                leading=22,
                textColor=colors.HexColor('#0F172A'),
                alignment=TA_CENTER,
                spaceAfter=12
            )
            style_subtitle = ParagraphStyle(
                'DocSubTitle',
                parent=styles['Normal'],
                fontSize=11,
                leading=15,
                textColor=colors.HexColor('#1D4ED8'),
                alignment=TA_CENTER,
                spaceAfter=25
            )
            style_h2 = ParagraphStyle(
                'DocH2',
                parent=styles['Heading2'],
                fontSize=12,
                leading=16,
                textColor=colors.HexColor('#0F172A'),
                spaceBefore=12,
                spaceAfter=6
            )
            style_body = ParagraphStyle(
                'DocBody',
                parent=styles['Normal'],
                fontSize=10,
                leading=14,
                textColor=colors.HexColor('#334155'),
                alignment=TA_JUSTIFY,
                spaceAfter=6
            )
            style_bullet = ParagraphStyle(
                'DocBullet',
                parent=styles['Normal'],
                fontSize=10,
                leading=14,
                textColor=colors.HexColor('#334155'),
                alignment=TA_JUSTIFY,
                leftIndent=15,
                spaceAfter=4
            )
            style_caption = ParagraphStyle(
                'DocCaption',
                parent=styles['Normal'],
                fontSize=9,
                leading=12,
                textColor=colors.HexColor('#475569'),
                alignment=TA_CENTER,
                spaceAfter=15
            )

            story = []
            
            # PORTADA PDF
            story.append(Paragraph(f"<b>{meta_inf['empresa'].upper()} | {meta_inf['unidad']}</b>", ParagraphStyle('SubC', parent=style_body, alignment=TA_RIGHT, fontSize=8.5, textColor=colors.HexColor('#64748B'))))
            story.append(Spacer(1, 40))
            story.append(Paragraph(f"<b>INFORME TÉCNICO DE MANTENIMIENTO</b><br/>{meta_inf['tipo'].upper()}", style_title))
            story.append(Paragraph(f"<b>ACTIVO: {meta_inf['tag']} - {meta_inf['equipo']}</b>", style_subtitle))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1D4ED8'), spaceAfter=20))
            
            # Tabla de metadatos PDF
            meta_table_data = [
                [Paragraph("<b>Fecha de Emisión:</b>", style_body), Paragraph(meta_inf["fecha"], style_body)],
                [Paragraph("<b>Horas de Operación:</b>", style_body), Paragraph(f"{meta_inf['horas']} hrs", style_body)],
                [Paragraph("<b>Elaborado por:</b>", style_body), Paragraph(f"{meta_inf['elaborador']} ({meta_inf['cargo']})", style_body)],
                [Paragraph("<b>Estándar:</b>", style_body), Paragraph("APA 7.ª Edición / Ingeniería de Mantenimiento", style_body)],
            ]
            t_meta = Table(meta_table_data, colWidths=[130, 370])
            t_meta.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(t_meta)
            story.append(PageBreak())

            # CUERPO DEL INFORME EN PDF (Con manejo de viñetas y párrafos justificados)
            for line in informe_md.split('\n'):
                line_clean = line.strip()
                if line_clean.startswith("## "):
                    story.append(Paragraph(f"<b>{line_clean.replace('## ', '')}</b>", style_h2))
                    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#CBD5E1'), spaceAfter=6))
                elif line_clean.startswith("# "):
                    story.append(Paragraph(f"<b>{line_clean.replace('# ', '')}</b>", style_h2))
                elif line_clean.startswith("* ") or line_clean.startswith("- "):
                    bullet_text = f"• {line_clean[2:].strip()}"
                    story.append(Paragraph(bullet_text, style_bullet))
                elif line_clean:
                    story.append(Paragraph(line_clean, style_body))

            # ANEXO FOTOGRÁFICO EN PDF (Centrado y con formato APA)
            if meta_inf["imagenes"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas (Estándar APA 7.ª Ed.)</b>", style_h2))
                story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1D4ED8'), spaceAfter=15))
                
                for item in meta_inf["imagenes"]:
                    img_bytes_io = io.BytesIO()
                    rgb_image = item["image_obj"].convert('RGB')
                    rgb_image.save(img_bytes_io, format='JPEG', quality=90)
                    img_bytes_io.seek(0)
                    
                    story.append(RLImage(img_bytes_io, width=400, height=260))
                    story.append(Spacer(1, 6))
                    caption_text = f"<b>Figura {item['index']}.</b> <i>{item['description']}</i><br/><b>Fuente:</b> Elaboración propia."
                    story.append(Paragraph(caption_text, style_caption))
                    story.append(Spacer(1, 10))

            # Construcción del PDF
            doc_pdf.build(story)
            pdf_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .pdf",
                data=pdf_buffer,
                file_name=f"Informe_Tecnico_{meta_inf['tag']}_{meta_inf['fecha']}.pdf",
                mime="application/pdf"
            )
