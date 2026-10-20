import io
import os
import json
from datetime import datetime, date
import streamlit as st
from PIL import Image
import pandas as pd
from google import genai

# Librerías para generación de documentos Word (.docx)
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Librerías para generación de documentos PDF (ReportLab)
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ==============================================================================
# CONFIGURACIÓN INICIAL Y DE PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Sistema Institucional de Informes de Mantenimiento",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Avanzados para Apariencia Institucional y Corporativa
st.markdown("""
    <style>
    /* Estilos Generales */
    .main {
        background-color: #F8FAFC;
    }
    .stAppHeader {
        background-color: rgba(255, 255, 255, 0.95);
    }
    .main-header {
        font-size: 2.3rem;
        color: #0F172A;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #475569;
        margin-bottom: 1.8rem;
    }
    .card {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
        border: 1px solid #E2E8F0;
        margin-bottom: 1.5rem;
    }
    .status-badge {
        background-color: #EFF6FF;
        color: #1D4ED8;
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }
    .stButton>button {
        border-radius: 0.5rem;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# GESTIÓN DE CLIENTE GEMINI API
# ==============================================================================
def get_gemini_client():
    """Obtiene y valida la instancia del cliente para la API de Gemini."""
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "gemini_api_key_input" in st.session_state and st.session_state["gemini_api_key_input"]:
        api_key = st.session_state["gemini_api_key_input"]
    
    if not api_key:
        return None
    return genai.Client(api_key=api_key)

# ==============================================================================
# BARRA LATERAL - CONFIGURACIÓN Y METADATOS DEL SISTEMA
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=70)
    st.markdown("## **Mantenimiento & Confiabilidad**")
    st.markdown("*Sistema Institucional v2.5*")
    st.markdown("---")
    
    # Autenticación de API Key si no está configurada en secrets.toml
    if "GEMINI_API_KEY" not in st.secrets:
        st.subheader("🔑 Autenticación")
        api_key_in = st.text_input("Gemini API Key:", type="password", key="gemini_api_key_input")
        if not api_key_in:
            st.info("💡 Ingrese su API Key para habilitar la generación por IA.")
        st.markdown("---")

    st.subheader("📋 Parámetros Institucionales")
    
    empresa_nombre = st.text_input("Empresa / Planta Industrial", "Planta Industrial Petroquímica - Complejo Sur")
    unidad_operativa = st.text_input("Unidad Operativa / Área", "Unidad de Hidrotratamiento (U-200)")
    
    tipo_informe = st.selectbox(
        "Tipo de Informe Técnico",
        [
            "Informe Ejecutivo de Mantenimiento",
            "Informe de Falla y Análisis Causa Raíz (RCA)",
            "Informe de Investigación de Incidente",
            "Informe de Integridad Mecánica",
            "Informe de Ensayos No Destructivos (NDT)",
            "Informe de Incidente Operacional",
            "Informe de Mantenimiento Correctivo",
            "Informe de Mantenimiento Preventivo / Overhaul",
            "Informe de Confiabilidad Operacional"
        ]
    )
    
    criticidad_equipo = st.selectbox(
        "Criticidad del Equipo (RCM)",
        ["Alta (Criticidad A - Vital)", "Media (Criticidad B - Esencial)", "Baja (Criticidad C - Secundario)"]
    )

    st.markdown("---")
    st.subheader("👤 Datos del Personal")
    elaborado_por = st.text_input("Elaborado por (Ingeniero)", "Ing. Roger Huamani Torres")
    cargo_elaborador = st.text_input("Cargo", "Ingeniero Senior de Confiabilidad y Mantenimiento")
    revisado_por = st.text_input("Revisado / Aprobado por", "Ing. Gerente de Planta y Operaciones")
    fecha_evento = st.date_input("Fecha de Ocurrencia / Trabajo", date.today())

# ==============================================================================
# VISTA PRINCIPAL Y NAVEGACIÓN
# ==============================================================================
st.markdown('<div class="main-header">Generador Inteligente de Informes de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Plataforma institucional para la documentación técnica, análisis de causa raíz y estandarización normativa (ISO 55001 / ISO 9001 / APA 7.ª Ed.).</div>', unsafe_allow_html=True)

# Pestañas principales de trabajo
tab_datos, tab_criticidad, tab_evidencias, tab_generacion = st.tabs([
    "📝 Contexto & Datos", 
    "📊 Matriz FMEA / RCM", 
    "📷 Evidencias Fotográficas", 
    "🚀 Generación & Exportación"
])

# ------------------------------------------------------------------------------
# TAB 1: CONTEXTO Y DATOS DEL EVENTO
# ------------------------------------------------------------------------------
with tab_datos:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 1. Datos del Equipo e Identificación")
    col_eq1, col_eq2, col_eq3 = st.columns(3)
    with col_eq1:
        equipo_tag = st.text_input("Tag del Equipo", "P-101A", help="Identificador único del activo.")
    with col_eq2:
        equipo_nombre = st.text_input("Nombre / Descripción", "Bomba Centrífuga Multietapa de Alimentación")
    with col_eq3:
        horas_operacion = st.number_input("Horas acumuladas de operación (TTSO)", min_value=0, value=14200, step=100)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 2. Formulario de Contexto del Evento")
    
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        contexto_evento = st.text_area(
            "Contexto Operacional del Evento",
            placeholder="Describa el régimen de trabajo, carga del proceso, temperatura ambiente, fluctuaciones previas...",
            height=140
        )
        descripcion_incidente = st.text_area(
            "Descripción Inicial del Incidente / Falla",
            placeholder="Detalle la manifestación inicial, disparos de alarmas, enclavamientos, fugas, ruidos anómalos o vibraciones...",
            height=140
        )
        hallazgos_tecnicos = st.text_area(
            "Hallazgos e Inspección Preliminar",
            placeholder="Ingrese los hallazgos físicos visuales, mediciones dimensionales, análisis de lubricantes, holguras, desgastes...",
            height=140
        )
    
    with col_t2:
        trabajos_ejecutados = st.text_area(
            "Trabajos y Procedimientos Ejecutados",
            placeholder="Enumere las tareas de desacople, desmontaje, metrología, cambio de repuestos, alineación láser, balanceo...",
            height=140
        )
        repuestos_utilizados = st.text_area(
            "Repuestos / Materiales Empleados",
            placeholder="Ej: 2x Rodamientos SKF 6316 C3, 1x Kit de Sellos Mecánicos Burgmann, 5L de Grasa Sintética Mobilith...",
            height=140
        )
        info_adicional = st.text_area(
            "Observaciones Técnicas e Historial Previo",
            placeholder="Información sobre análisis de vibraciones previos, intervenciones anteriores, modificaciones de diseño...",
            height=140
        )
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 2: MATRIZ DE CRITICIDAD Y FMEA (RCM)
# ------------------------------------------------------------------------------
with tab_criticidad:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Evaluación Preliminar de Modos de Falla (FMEA / RCM)")
    st.info("Esta sección permite estructurar parámetros cuantitativos que la Inteligencia Artificial integrará en la Evaluación Mecánica, Eléctrica y Riesgos.")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        severidad = st.slider("Severidad del Impacto (S: 1-10)", 1, 10, 7, help="1: Insignificante, 10: Catastrófico para seguridad o producción.")
    with col_f2:
        ocurrencia = st.slider("Probabilidad de Ocurrencia (O: 1-10)", 1, 10, 5, help="1: Muy improbable, 10: Frecuencia extremadamente alta.")
    with col_f3:
        deteccion = st.slider("Facilidad de Detección (D: 1-10)", 1, 10, 4, help="1: Certeza de detección temprana, 10: Indetectable previo a falla.")
    
    rpn = severidad * ocurrencia * deteccion
    st.markdown(f"**Número de Prioridad de Riesgo (NPR / RPN Calculado):** `<span class='status-badge'>{rpn} / 1000</span>`", unsafe_allow_html=True)
    
    modo_falla_p = st.text_input("Modo de Falla Principal Identificado / Presunto", "Desgaste Abrasivo y Fatiga Superficial en Rodamiento Radial Lado Libres")
    efecto_falla_p = st.text_input("Efecto Potencial en la Planta", "Pérdida de capacidad de bombeo, riesgo de contaminación por lubricante y paro de unidad de refinación.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 3: EVIDENCIAS FOTOGRÁFICAS (APA 7.ª EDICIÓN)
# ------------------------------------------------------------------------------
with tab_evidencias:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Carga de Registro Fotográfico y Validación Normativa")
    st.caption("Conforme a la norma ISO 9001 y los estándares de documentación técnica APA 7.ª edición, cada evidencia debe contar obligatoriamente con una descripción explicativa para garantizar la trazabilidad.")
    
    uploaded_files = st.file_uploader(
        "Seleccionar imágenes de evidencia (Formatos soportados: JPG, JPEG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    
    image_data_list = []
    if uploaded_files:
        st.success(f"Se han cargado {len(uploaded_files)} imágenes correctamente.")
        st.markdown("---")
        
        for idx, file_item in enumerate(uploaded_files):
            col_i1, col_i2 = st.columns([1, 2])
            with col_i1:
                img_obj = Image.open(file_item)
                st.image(img_obj, caption=f"Evidencia N.° {idx+1}", use_container_width=True)
            with col_i2:
                desc_input = st.text_input(
                    f"Descripción técnica obligatoria para Figura {idx+1}:",
                    key=f"img_desc_key_{idx}",
                    placeholder="Ejemplo: Aspecto de la pista interna del rodamiento evidenciando desconchado (spalling) severo."
                )
                image_data_list.append({
                    "file_name": file_item.name,
                    "image_obj": img_obj,
                    "description": desc_input,
                    "index": idx + 1
                })
    else:
        st.warning("No se han subido archivos de imagen. El sistema requiere al menos una evidencia con descripción técnica para generar el informe final.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 4: GENERACIÓN CON IA Y EXPORTACIÓN
# ------------------------------------------------------------------------------
with tab_generacion:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Generación Automatizada del Informe")
    
    # Validaciones obligatorias antes de procesar
    has_images = len(image_data_list) > 0
    missing_desc = any(not img["description"].strip() for img in image_data_list) if has_images else False
    
    if not has_images:
        st.error("⚠️ Validación pendiente: Debe cargar al menos una imagen de evidencia fotográfica en la pestaña 'Evidencias Fotográficas'.")
    elif missing_desc:
        st.error("⚠️ Validación pendiente: Existen imágenes cargadas que no tienen descripción. Por favor complete todos los campos de descripción.")
    else:
        st.success("✅ Todas las validaciones de datos y evidencias han sido superadas satisfactoriamente.")
    
    btn_generar = st.button("🧠 Generar Informe Técnico de Ingeniería", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_generar and has_images and not missing_desc:
        client = get_gemini_client()
        if not client:
            st.error("❌ No se encontró la clave de API de Gemini. Verifique su archivo secrets.toml o la clave ingresada en la barra lateral.")
        else:
            with st.spinner("Procesando análisis de ingeniería, infiriendo causa raíz y generando el informe completo con IA..."):
                
                # Resumen de imágenes para el prompt
                evidencias_prompt_str = "\n".join([f"- Figura {item['index']}: {item['description']}" for item in image_data_list])
                
                # Construcción del Prompt Técnico Completo
                prompt_master = f"""
Actúa como un Ingeniero Senior de Mantenimiento, Ingeniero de Confiabilidad e Investigador de Incidentes Industriales con más de 25 años de experiencia en plantas de alta disponibilidad.
Analiza minuciosamente la información suministrada y genera un Informe Técnico Corporativo de Mantenimiento de nivel ejecutivo.

Parámetros Generales:
- Empresa: {empresa_nombre}
- Unidad Operativa: {unidad_operativa}
- Tipo de Informe: {tipo_informe}
- Equipo: {equipo_tag} - {equipo_nombre}
- Horas Acumuladas: {horas_operacion} hrs
- Criticidad RCM: {criticidad_equipo}
- RPN Evaluado (FMEA): {rpn} (Severidad: {severidad}, Ocurrencia: {ocurrencia}, Detección: {deteccion})
- Modo de Falla Principal: {modo_falla_p}
- Efecto Potencial: {efecto_falla_p}
- Elaborado por: {elaborado_por} ({cargo_elaborador})

Datos del Evento:
- Contexto Operacional: {contexto_evento}
- Descripción del Incidente: {descripcion_incidente}
- Hallazgos Técnicos: {hallazgos_tecnicos}
- Trabajos Ejecutados: {trabajos_ejecutados}
- Repuestos Empleados: {repuestos_utilizados}
- Observaciones Adicionales: {info_adicional}

Registro de Evidencias Fotográficas:
{evidencias_prompt_str}

REQUISITOS OBLIGATORIOS DE ESTRUCTURA Y REDACCIÓN:
La respuesta DEBE contener exactamente los siguientes 19 encabezados numerados en formato Markdown (utilizando ## para cada título). Desarrolla cada punto con análisis profundo, terminología técnica avanzada de ingeniería, sin texto robótico, sin repeticiones genéricas y con total coherencia técnica.

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

Genera el texto directamente iniciando en la sección 1.
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
                        "revisador": revisado_por,
                        "fecha": str(fecha_evento),
                        "imagenes": image_data_list
                    }
                    st.success("¡El informe técnico ha sido generado satisfactoriamente!")
                except Exception as ex:
                    st.error(f"Error durante la generación con la API de Gemini: {ex}")

# ==============================================================================
# VISUALIZACIÓN Y EXPORTACIÓN DEL INFORME TÉCNICO
# ==============================================================================
if "informe_generado" in st.session_state:
    st.markdown("---")
    st.markdown("## 📄 Vista Previa del Informe Generado")
    
    informe_md = st.session_state["informe_generado"]
    meta_inf = st.session_state["datos_informe"]
    
    with st.expander("Ver informe completo en formato Markdown", expanded=True):
        st.markdown(informe_md)
        
        st.markdown("### 📷 Registro de Evidencias Fotográficas")
        for item in meta_inf["imagenes"]:
            st.image(
                item["image_obj"], 
                caption=f"Figura {item['index']}. {item['description']} (Fuente: Elaboración propia)", 
                use_container_width=True
            )

    st.markdown("---")
    st.markdown("### 📥 Exportación Documental Institucional")
    
    col_exp_docx, col_exp_pdf = st.columns(2)
    
    # --------------------------------------------------------------------------
    # EXPORTACIÓN A MICROSOFT WORD (.DOCX)
    # --------------------------------------------------------------------------
    with col_exp_docx:
        if st.button("📥 Generar Documento Word (.docx)", use_container_width=True):
            doc = Document()
            
            # Margen institucional (1 pulgada por lado)
            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # Portada Corporativa
            p_head = doc.add_paragraph()
            p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_head = p_head.add_run(f"{meta_inf['empresa']}\n{meta_inf['unidad']}")
            r_head.font.size = Pt(9)
            r_head.font.color.rgb = RGBColor(100, 116, 139)
            
            doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(f"INFORME TÉCNICO DE MANTENIMIENTO\n{meta_inf['tipo'].upper()}")
            r_title.font.size = Pt(18)
            r_title.font.bold = True
            r_title.font.color.rgb = RGBColor(15, 23, 42)
            
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sub = p_sub.add_run(f"ACTIVO: {meta_inf['tag']} - {meta_inf['equipo']}")
            r_sub.font.size = Pt(12)
            r_sub.font.bold = True
            r_sub.font.color.rgb = RGBColor(29, 78, 216)
            
            doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Tabla de metadatos en portada Word
            table_meta = doc.add_table(rows=5, cols=2)
            table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
            table_data = [
                ("Fecha de Emisión:", meta_inf["fecha"]),
                ("Horas de Operación (TTSO):", f"{meta_inf['horas']} Horas"),
                ("Elaborado por:", f"{meta_inf['elaborador']} ({meta_inf['cargo']})"),
                ("Revisado / Aprobado por:", meta_inf["revisador"]),
                ("Estándares de Cumplimiento:", "ISO 55001 / ISO 9001 / APA 7.ª Ed.")
            ]
            for i, (label, val) in enumerate(table_data):
                row_cells = table_meta.rows[i].cells
                row_cells[0].text = label
                row_cells[0].paragraphs[0].runs[0].font.bold = True
                row_cells[1].text = val
            
            doc.add_page_break()
            
            # Cuerpo del Documento
            for line in informe_md.split('\n'):
                line_str = line.strip()
                if line_str.startswith("## "):
                    p_h = doc.add_heading(line_str.replace("## ", ""), level=2)
                    p_h.runs[0].font.color.rgb = RGBColor(15, 23, 42)
                elif line_str.startswith("# "):
                    p_h = doc.add_heading(line_str.replace("# ", ""), level=1)
                    p_h.runs[0].font.color.rgb = RGBColor(30, 58, 138)
                elif line_str.startswith("### "):
                    doc.add_heading(line_str.replace("### ", ""), level=3)
                elif line_str:
                    p = doc.add_paragraph(line_str)
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(4)

            # Inserción de Fotografías con Formato APA 7 en Word
            doc.add_heading("Anexo: Evidencias Fotográficas (APA 7.ª Ed.)", level=2)
            for item in meta_inf["imagenes"]:
                img_mem = io.BytesIO()
                item["image_obj"].save(img_mem, format='PNG')
                img_mem.seek(0)
                
                doc.add_picture(img_mem, width=Inches(5.0))
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_cap1 = p_cap.add_run(f"Figura {item['index']}.\n")
                r_cap1.font.bold = True
                r_cap1.font.size = Pt(9.5)
                r_cap2 = p_cap.add_run(f"{item['description']}\nFuente: Elaboración propia.")
                r_cap2.font.italic = True
                r_cap2.font.size = Pt(9.5)
                doc.add_paragraph()

            # Guardar en búfer
            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Documento .docx",
                data=docx_buffer,
                file_name=f"Informe_Tecnico_{meta_inf['tag']}_{meta_inf['fecha']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # --------------------------------------------------------------------------
    # EXPORTACIÓN A PDF (REPORTLAB - SIN ERRORES DE DISCO EN NUBE)
    # --------------------------------------------------------------------------
    with col_exp_pdf:
        if st.button("📥 Generar Documento PDF (.pdf)", use_container_width=True):
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
            
            # Estilos Personalizados para ReportLab
            style_title = ParagraphStyle(
                'DocTitle',
                parent=styles['Heading1'],
                fontSize=20,
                leading=24,
                textColor=colors.HexColor('#0F172A'),
                alignment=1,
                spaceAfter=15
            )
            style_subtitle = ParagraphStyle(
                'DocSubTitle',
                parent=styles['Normal'],
                fontSize=12,
                leading=16,
                textColor=colors.HexColor('#1D4ED8'),
                alignment=1,
                spaceAfter=30
            )
            style_h2 = ParagraphStyle(
                'DocH2',
                parent=styles['Heading2'],
                fontSize=13,
                leading=17,
                textColor=colors.HexColor('#0F172A'),
                spaceBefore=14,
                spaceAfter=6
            )
            style_body = ParagraphStyle(
                'DocBody',
                parent=styles['Normal'],
                fontSize=9.5,
                leading=13.5,
                textColor=colors.HexColor('#334155'),
                spaceAfter=8
            )
            style_caption = ParagraphStyle(
                'DocCaption',
                parent=styles['Normal'],
                fontSize=8.5,
                leading=11,
                textColor=colors.HexColor('#475569'),
                spaceAfter=12
            )

            story = []
            
            # PORTADA PDF
            story.append(Paragraph(f"<b>{meta_inf['empresa'].upper()}</b>", style_caption))
            story.append(Paragraph(f"{meta_inf['unidad']}", style_caption))
            story.append(Spacer(1, 40))
            story.append(Paragraph(f"<b>INFORME TÉCNICO DE MANTENIMIENTO</b><br/>{meta_inf['tipo'].upper()}", style_title))
            story.append(Paragraph(f"<b>ACTIVO: {meta_inf['tag']} - {meta_inf['equipo']}</b>", style_subtitle))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1D4ED8'), spaceAfter=20))
            
            # Tabla Metadatos en PDF
            meta_table_data = [
                [Paragraph("<b>Fecha de Emisión:</b>", style_body), Paragraph(meta_inf["fecha"], style_body)],
                [Paragraph("<b>Horas de Operación:</b>", style_body), Paragraph(f"{meta_inf['horas']} hrs", style_body)],
                [Paragraph("<b>Elaborado por:</b>", style_body), Paragraph(f"{meta_inf['elaborador']}", style_body)],
                [Paragraph("<b>Cargo:</b>", style_body), Paragraph(f"{meta_inf['cargo']}", style_body)],
                [Paragraph("<b>Revisado / Aprobado:</b>", style_body), Paragraph(f"{meta_inf['revisador']}", style_body)],
            ]
            t_meta = Table(meta_table_data, colWidths=[150, 300])
            t_meta.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(t_meta)
            story.append(PageBreak())

            # CUERPO DEL INFORME EN PDF
            for line in informe_md.split('\n'):
                line_clean = line.strip()
                if line_clean.startswith("## "):
                    story.append(Paragraph(f"<b>{line_clean.replace('## ', '')}</b>", style_h2))
                    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#CBD5E1'), spaceAfter=8))
                elif line_clean.startswith("# "):
                    story.append(Paragraph(f"<b>{line_clean.replace('# ', '')}</b>", style_h2))
                elif line_clean:
                    story.append(Paragraph(line_clean, style_body))

            # ANEXO FOTOGRÁFICO EN PDF (PROCESADO 100% EN MEMORIA CON BYTESIO)
            if meta_inf["imagenes"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas (Estándar APA 7.ª Ed.)</b>", style_h2))
                story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1D4ED8'), spaceAfter=12))
                
                for item in meta_inf["imagenes"]:
                    img_bytes_io = io.BytesIO()
                    rgb_image = item["image_obj"].convert('RGB')
                    rgb_image.save(img_bytes_io, format='JPEG', quality=90)
                    img_bytes_io.seek(0)
                    
                    story.append(RLImage(img_bytes_io, width=420, height=280))
                    story.append(Spacer(1, 4))
                    caption_text = f"<b>Figura {item['index']}.</b> <i>{item['description']}</i><br/><b>Fuente:</b> Elaboración propia."
                    story.append(Paragraph(caption_text, style_caption))
                    story.append(Spacer(1, 12))

            # Construcción del documento PDF
            doc_pdf.build(story)
            pdf_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Documento .pdf",
                data=pdf_buffer,
                file_name=f"Informe_Tecnico_{meta_inf['tag']}_{meta_inf['fecha']}.pdf",
                mime="application/pdf"
            )
