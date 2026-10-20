"""
================================================================================
SISTEMA INSTITUCIONAL DE INFORMES TÉCNICOS DE MANTENIMIENTO (SITM)
Versión: 4.0.0 - Edición Corporativa Institucional
Desarrollado para: Plantas Industriales de Alta Confiabilidad y Seguridad
================================================================================
"""

import io
import os
import re
from datetime import datetime, date
import streamlit as st
from PIL import Image
from google import genai

# Librerías para generación de documentos Word (.docx)
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Librerías para generación de documentos PDF (ReportLab)
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

# ==============================================================================
# 1. CONFIGURACIÓN INICIAL DE LA APLICACIÓN STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="SITM - Sistema Institucional de Mantenimiento",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Corporativos Avanzados
st.markdown("""
    <style>
    .main {
        background-color: #F8FAFC;
    }
    .stAppHeader {
        background-color: rgba(255, 255, 255, 0.95);
    }
    .main-header {
        font-size: 2.2rem;
        color: #0F172A;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #475569;
        margin-bottom: 1.8rem;
    }
    .card-container {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
        border: 1px solid #E2E8F0;
        margin-bottom: 1.25rem;
    }
    .badge-institutional {
        background-color: #E0E7FF;
        color: #3730A3;
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. GESTIÓN DE CLIENTE Y CONEXIÓN CON GEMINI API
# ==============================================================================
def inicializar_cliente_gemini():
    """
    Valida y retorna la instancia oficial del cliente Google GenAI utilizando
    el entorno seguro de Streamlit Secrets o el input dinámico del usuario.
    """
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "gemini_api_key_input" in st.session_state and st.session_state["gemini_api_key_input"]:
        api_key = st.session_state["gemini_api_key_input"]
    
    if not api_key:
        return None
    return genai.Client(api_key=api_key)

# ==============================================================================
# 3. CANVAS PERSONALIZADO CON METADATOS Y PIE DE PÁGINA CORPORATIVO
# ==============================================================================
class CanvasInstitucional(canvas.Canvas):
    """
    Canvas avanzado con patrón de doble pasada para numeración dinámica 
    'Página X de Y' e inserción de cabeceras/pies corporativos configurables.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        # Valores por defecto que se sobrescribirán dinámicamente
        self.empresa_meta = "COMPAÑÍA INDUSTRIAL"
        self.codigo_doc = "SITM-REP-01"

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.dibujar_elementos_pagina(num_pages)
            super().showPage()
        super().save()

    def dibujar_elementos_pagina(self, page_count):
        self.saveState()
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(colors.HexColor("#475569"))
        
        # Omitir cabecera y pie en la Portada (Página 1)
        if self._pageNumber > 1:
            # Cabecera superior institucional
            self.drawString(54, 750, f"{self.empresa_meta.upper()} — INFORME TÉCNICO DE MANTENIMIENTO")
            self.drawRightString(558, 750, f"DOC: {self.codigo_doc}")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.75)
            self.line(54, 742, 558, 742)
            
            # Pie de página inferior
            self.setFont("Helvetica", 8)
            page_text = f"Página {self._pageNumber} de {page_count}"
            self.drawRightString(558, 36, page_text)
            self.drawString(54, 36, "DOCUMENTO OFICIAL DE INGENIERÍA — USO INTERNO Y CONFIDENCIAL")
            self.line(54, 48, 558, 48)
            
        self.restoreState()

# ==============================================================================
# 4. BARRA LATERAL - CONFIGURACIÓN INSTITUCIONAL Y ORGANIZACIONAL
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=65)
    st.markdown("### **Panel Institucional SITM**")
    st.markdown("*Configuración de Encabezados y Metadatos*")
    st.markdown("---")
    
    # Autenticación API Key
    if "GEMINI_API_KEY" not in st.secrets:
        st.markdown("#### 🔑 Credencial API Gemini")
        st.text_input("API Key:", type="password", key="gemini_api_key_input")
        st.markdown("---")

    st.markdown("#### 🏢 Datos de la Compañía")
    empresa_nombre = st.text_input("Nombre de la Compañía", "ORICA")
    unidad_operativa = st.text_input("Unidad / Planta / Complejo", "Unidad de Detonadores")
    codigo_documento = st.text_input("Código de Documento / Ref.", "ORICA-MANT-2026-092")
    nivel_confidencialidad = st.selectbox(
        "Nivel de Confidencialidad",
        ["CONFIDENCIAL - USO INTERNO", "RESERVADO - GERENCIA", "ESTÁNDAR OPERATIVO"]
    )
    
    st.markdown("---")
    st.markdown("#### 📋 Clasificación del Informe")
    tipo_informe = st.selectbox(
        "Tipo de Reporte Técnico",
        [
            "Informe Ejecutivo de Mantenimiento",
            "Informe de Falla y Análisis Causa Raíz",
            "Informe de Investigación de Incidente",
            "Informe de Mantenimiento Correctivo",
            "Informe de Mantenimiento Preventivo",
            "Informe de Confiabilidad Operacional"
        ]
    )

    st.markdown("---")
    st.markdown("#### 👤 Responsables del Informe")
    elaborado_por = st.text_input("Elaborado por", "Ing. Roger Huamani Torres")
    cargo_elaborador = st.text_input("Cargo del Elaborador", "Superintendente de Mantenimiento")
    revisado_por = st.text_input("Revisado / Aprobado por", "Gerencia de Operaciones y Planta")
    fecha_evento = st.date_input("Fecha de Ocurrencia / Emisión", date.today())

# ==============================================================================
# 5. INTERFAZ PRINCIPAL DE ENTRADA (DISEÑO PROFESIONAL)
# ==============================================================================
st.markdown('<div class="main-header">Generador Institucional de Informes de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sub-header">Plataforma experta adaptada al modelo corporativo. Redacte el contexto de forma sencilla y la IA generará un documento técnico profundo, redactado en párrafos estructurados y con uso riguroso de viñetas.</div>', unsafe_allow_html=True)

# Pestañas de flujo de trabajo
tab_datos, tab_evidencias, tab_preview = st.tabs([
    "📝 1. Datos e Información del Evento", 
    "📷 2. Evidencias Fotográficas", 
    "🚀 3. Vista Previa y Exportación"
])

# ------------------------------------------------------------------------------
# PESTAÑA 1: DATOS Y CONTEXTO DEL EVENTO
# ------------------------------------------------------------------------------
with tab_datos:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Identificación del Activo Principal")
    col_e1, col_e2, col_e3 = st.columns(3)
    with col_e1:
        equipo_tag = st.text_input("Tag del Equipo / Máquina", "230-1", help="Ej: 230-1 o P-101A")
    with col_e2:
        equipo_nombre = st.text_input("Nombre / Descripción del Activo", "Carga de Retardos (Línea Principal)")
    with col_e3:
        horas_operacion = st.number_input("Horas de Operación Acumuladas (TTSO)", min_value=0, value=14200, step=100)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Contexto y Antecedentes Técnicos")
    st.info("💡 **Instrucción:** Ingrese los detalles de forma clara. La IA generará explicaciones técnicas exhaustivas en formato de párrafos profesionales, evitando resúmenes excesivos y utilizando viñetas únicamente para enumerar componentes o secuencias necesarias.")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        contexto_evento = st.text_area(
            "Contexto Operacional del Evento",
            placeholder="Ej: Operación en régimen nominal produciendo detonadores de 500ms en 30mm...",
            height=140
        )
        descripcion_incidente = st.text_area(
            "Descripción Inicial del Incidente o Anomalía",
            placeholder="Ej: Detección de ruido anómalo y presencia de humo en Estación ST23, activación de parada de emergencia...",
            height=140
        )
        hallazgos_raw = st.text_area(
            "Hallazgos Encontrados en Inspección",
            placeholder="Ej: Falla en sellos O-ring en ST21 y ST22, desgaste y acumulación de polución en escobillas, rotura en tanque ST23...",
            height=140
        )
    with col_c2:
        trabajos_realizados = st.text_area(
            "Trabajos y Acciones Ejecutadas",
            placeholder="Ej: Sustitución de O-rings en ST21/ST22, reemplazo de tanque ST23, revisión y limpieza de escobillas, pruebas operativas...",
            height=140
        )
        repuestos_materiales = st.text_area(
            "Repuestos y Materiales Empleados",
            placeholder="Ej: 2x O-rings elastoméricos de alta especificación, 1x Tanque superior ST23...",
            height=140
        )
        info_adicional = st.text_area(
            "Información Adicional u Observaciones",
            placeholder="Ej: Historial de intervenciones previas, recomendaciones operativas y condiciones del entorno...",
            height=140
        )
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 2: EVIDENCIAS FOTOGRÁFICAS
# ------------------------------------------------------------------------------
with tab_evidencias:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📷 Registro de Evidencias Fotográficas")
    st.caption("Cargue las fotografías necesarias. El sistema asignará automáticamente la numeración de figuras y sus descripciones centradas conforme al estándar APA y al modelo corporativo.")
    
    uploaded_files = st.file_uploader(
        "Seleccione o arrastre las imágenes (Formatos: JPG, JPEG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    
    lista_imagenes = []
    if uploaded_files:
        st.success(f"Se han cargado **{len(uploaded_files)}** archivos de imagen.")
        st.markdown("---")
        for idx, archivo in enumerate(uploaded_files):
            col_img, col_txt = st.columns([1, 2])
            with col_img:
                img_obj = Image.open(archivo)
                st.image(img_obj, caption=f"Evidencia N.° {idx+1}", use_container_width=True)
            with col_txt:
                desc_img = st.text_input(
                    f"Descripción técnica obligatoria para Figura {idx+1}:",
                    key=f"desc_img_input_{idx}",
                    placeholder="Ej: Registro fotográfico post-incidente de la Estación ST23 evidenciando daños estructurales y residuos..."
                )
                lista_imagenes.append({
                    "nombre_archivo": archivo.name,
                    "objeto_imagen": img_obj,
                    "descripcion": desc_img,
                    "indice": idx + 1
                })
    else:
        st.warning("⚠️ No se han adjuntado imágenes. El informe requiere al menos una evidencia fotográfica con su descripción.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 3: VISTA PREVIA Y GENERACIÓN
# ------------------------------------------------------------------------------
with tab_preview:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🚀 Consolidación y Generación del Informe Institucional")
    
    tiene_imagenes = len(lista_imagenes) > 0
    falta_descripcion = any(not img["descripcion"].strip() for img in lista_imagenes) if tiene_imagenes else False
    
    if not tiene_imagenes:
        st.error("❌ Verificación pendiente: Debe cargar al menos una imagen en la pestaña anterior.")
    elif falta_descripcion:
        st.error("❌ Verificación pendiente: Todas las imágenes cargadas deben poseer una descripción técnica obligatoria.")
    else:
        st.success("✅ Todas las validaciones operativas han sido superadas con éxito.")
        
    btn_generar = st.button("✨ Generar Informe Técnico Institucional con IA", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_generar and tiene_imagenes and not falta_descripcion:
        cliente_ai = inicializar_cliente_gemini()
        if not cliente_ai:
            st.error("❌ Error de Autenticación: Configure su clave API de Gemini en Streamlit Secrets o en la barra lateral.")
        else:
            with st.spinner("🔄 Procesando análisis forense de ingeniería, estructurando párrafos técnicos y generando el informe completo..."):
                
                evidencias_txt = "\n".join([f"- Figura {item['indice']}: {item['descripcion']}" for item in lista_imagenes])
                
                # Prompt altamente especializado para imitar el modelo ORICA institucional
                prompt_ia = f"""
Actúa como un Ingeniero Senior de Mantenimiento, Ingeniero de Confiabilidad e Investigador de Incidentes Industriales con más de 25 años de experiencia en operaciones críticas (bajo el estándar de rigor de reportes corporativos de {empresa_nombre}).
Genera un informe técnico completo, profundo, riguroso y de nivel institucional.

DIRECTRICES DE REDACCIÓN OBLIGATORIAS:
1. **Formato Principal en Párrafos:** Redacta cada sección principal de manera extensa, analítica, técnica y desarrollada en forma de párrafos fluidos. NO resumas en exceso; el informe debe poseer la profundidad técnica y la longitud de un reporte de ingeniería formal (similar al modelo proporcionado).
2. **Uso Restringido y Preciso de Viñetas:** PROHIBIDO utilizar viñetas de forma masiva en todo el documento. Utiliza viñetas (con asteriscos `*`) UNICAMENTE cuando sea gramaticalmente necesario para enlistar componentes físicos discretos (ej. sellos, escobillas, estaciones), secuencias numeradas de eventos o recomendaciones específicas enumeradas.
3. **Trazabilidad y Coherencia:** Todo el contenido debe basarse estrictamente en la información proporcionada, expandiéndola con criterio experto industrial.

Parámetros del Reporte:
- Empresa: {empresa_nombre}
- Unidad Operativa: {unidad_operativa}
- Código de Documento: {codigo_documento}
- Tipo de Informe: {tipo_informe}
- Activo Principal: {equipo_tag} - {equipo_nombre}
- Horas de Operación Acumuladas: {horas_operacion} hrs
- Elaborado por: {elaborado_por} ({cargo_elaborador})

Datos proporcionados por el equipo técnico:
- Contexto Operacional: {contexto_evento}
- Descripción del Incidente: {descripcion_incidente}
- Hallazgos Encontrados: {hallazgos_raw}
- Trabajos Ejecutados: {trabajos_realizados}
- Repuestos y Materiales: {repuestos_materiales}
- Información Adicional: {info_adicional}
- Evidencias Fotográficas Adjuntas:
{evidencias_txt}

ESTRUCTURA OBLIGATORIA (Desarrolla rigurosamente cada sección bajo los siguientes 19 títulos en formato Markdown `##`):
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
"""
                try:
                    respuesta = cliente_ai.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt_ia
                    )
                    st.session_state["texto_informe"] = respuesta.text
                    st.session_state["meta_informe"] = {
                        "empresa": empresa_nombre,
                        "unidad": unidad_operativa,
                        "codigo": codigo_documento,
                        "confidencialidad": nivel_confidencialidad,
                        "tipo": tipo_informe,
                        "tag": equipo_tag,
                        "equipo": equipo_nombre,
                        "horas": horas_operacion,
                        "elaborador": elaborado_por,
                        "cargo": cargo_elaborador,
                        "revisador": revisado_por,
                        "fecha": str(fecha_evento),
                        "imagenes": lista_imagenes
                    }
                    st.success("¡Informe técnico institucional generado con éxito!")
                except Exception as e:
                    st.error(f"Error al conectar con la API de Gemini: {e}")

# ==============================================================================
# 6. VISUALIZACIÓN Y MOTOR DE EXPORTACIÓN (WORD Y PDF CORPORATIVOS)
# ==============================================================================
if "texto_informe" in st.session_state:
    st.markdown("---")
    st.markdown("## 📄 Vista Previa del Informe Institucional")
    
    contenido_md = st.session_state["texto_informe"]
    metadatos = st.session_state["meta_informe"]
    
    with st.expander("🔍 Ver contenido completo del informe", expanded=True):
        st.markdown(contenido_md)
        st.markdown("### 📷 Registro Fotográfico Adjunto")
        for item in metadatos["imagenes"]:
            st.image(
                item["objeto_imagen"], 
                caption=f"Figura {item['indice']}. {item['descripcion']} (Fuente: Elaboración propia)", 
                use_container_width=True
            )

    st.markdown("---")
    st.markdown("### 📥 Exportación de Documentos Oficiales")
    col_w, col_p = st.columns(2)
    
    # --------------------------------------------------------------------------
    # 6.1 EXPORTACIÓN A MICROSOFT WORD (.DOCX)
    # --------------------------------------------------------------------------
    with col_w:
        if st.button("📥 Descargar Documento Word (.docx)", use_container_width=True):
            doc = Document()
            
            # Márgenes institucionales (1 pulgada / 2.54 cm)
            for sec in doc.sections:
                sec.top_margin = Inches(1)
                sec.bottom_margin = Inches(1)
                sec.left_margin = Inches(1)
                sec.right_margin = Inches(1)

            # Portada Institucional Word
            p_sup = doc.add_paragraph()
            p_sup.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_sup = p_sup.add_run(f"{metadatos['empresa']} | {metadatos['unidad']}\nDOC: {metadatos['codigo']}")
            r_sup.font.size = Pt(9)
            r_sup.font.color.rgb = RGBColor(100, 116, 139)
            
            doc.add_paragraph()
            
            p_tit = doc.add_paragraph()
            p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_tit = p_tit.add_run(f"INFORME TÉCNICO DE MANTENIMIENTO\n{metadatos['tipo'].upper()}")
            r_tit.font.size = Pt(16)
            r_tit.font.bold = True
            r_tit.font.color.rgb = RGBColor(15, 23, 42)
            
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sub = p_sub.add_run(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}")
            r_sub.font.size = Pt(11)
            r_sub.font.bold = True
            r_sub.font.color.rgb = RGBColor(30, 64, 175)
            
            doc.add_paragraph()
            
            # Tabla de metadatos corporativos en portada
            t_meta = doc.add_table(rows=6, cols=2)
            t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
            datos_t = [
                ("Compañía / Planta:", metadatos["empresa"]),
                ("Fecha del Evento:", metadatos["fecha"]),
                ("Horas de Operación (TTSO):", f"{metadatos['horas']} hrs"),
                ("Elaborado por:", f"{metadatos['elaborador']} ({metadatos['cargo']})"),
                ("Aprobado por:", metadatos["revisador"]),
                ("Clasificación:", metadatos["confidencialidad"])
            ]
            for idx_t, (lbl, val) in enumerate(datos_t):
                c_cells = t_meta.rows[idx_t].cells
                c_cells[0].text = lbl
                c_cells[0].paragraphs[0].runs[0].font.bold = True
                c_cells[1].text = val
            
            doc.add_page_break()
            
            # Procesamiento de líneas conservando párrafos justificados y viñetas selectivas
            for line in contenido_md.split('\n'):
                line_text = line.strip()
                if line_text.startswith("## "):
                    h = doc.add_heading(line_text.replace("## ", ""), level=2)
                    h.runs[0].font.color.rgb = RGBColor(15, 23, 42)
                elif line_text.startswith("# "):
                    h = doc.add_heading(line_text.replace("# ", ""), level=1)
                    h.runs[0].font.color.rgb = RGBColor(30, 64, 175)
                elif line_text.startswith("* ") or line_text.startswith("- "):
                    p_bullet = doc.add_paragraph(style='List Bullet')
                    r_b = p_bullet.add_run(line_text[2:].strip())
                    r_b.font.size = Pt(10.5)
                    r_b.font.name = 'Arial'
                    p_bullet.paragraph_format.line_spacing = 1.15
                    p_bullet.paragraph_format.space_after = Pt(3)
                elif line_text:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    r = p.add_run(line_text)
                    r.font.size = Pt(10.5)
                    r.font.name = 'Arial'
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(4)

            # Anexo Fotográfico Centrado con Estándar APA en Word
            doc.add_heading("Anexo: Evidencias Fotográficas (APA 7.ª Ed.)", level=2)
            for item in metadatos["imagenes"]:
                img_io = io.BytesIO()
                item["objeto_imagen"].save(img_io, format='PNG')
                img_io.seek(0)
                
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_io, width=Inches(4.5))
                
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_c1 = p_cap.add_run(f"Figura {item['indice']}. ")
                r_c1.font.bold = True
                r_c1.font.size = Pt(9.5)
                r_c2 = p_cap.add_run(f"{item['descripcion']}\nFuente: Elaboración propia.")
                r_c2.font.italic = True
                r_c2.font.size = Pt(9.5)
                doc.add_paragraph()

            buffer_docx = io.BytesIO()
            doc.save(buffer_docx)
            buffer_docx.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .docx",
                data=buffer_docx,
                file_name=f"Informe_Tecnico_{metadatos['tag']}_{metadatos['fecha']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # --------------------------------------------------------------------------
    # 6.2 EXPORTACIÓN A PDF (REPORTLAB - JUSTIFICADO, VIÑETAS Y CENTRADO)
    # --------------------------------------------------------------------------
    with col_p:
        if st.button("📥 Descargar Documento PDF (.pdf)", use_container_width=True):
            buffer_pdf = io.BytesIO()
            doc_pdf = SimpleDocTemplate(
                buffer_pdf,
                pagesize=letter,
                rightMargin=54,
                leftMargin=54,
                topMargin=54,
                bottomMargin=54
            )
            
            styles = getSampleStyleSheet()
            
            # Estilos tipográficos profesionales
            style_title = ParagraphStyle(
                'PdfTitle',
                parent=styles['Heading1'],
                fontName='Helvetica-Bold',
                fontSize=18,
                leading=22,
                textColor=colors.HexColor('#0F172A'),
                alignment=TA_CENTER,
                spaceAfter=12
            )
            style_subtitle = ParagraphStyle(
                'PdfSubTitle',
                parent=styles['Normal'],
                fontName='Helvetica-Bold',
                fontSize=11,
                leading=15,
                textColor=colors.HexColor('#1D4ED8'),
                alignment=TA_CENTER,
                spaceAfter=20
            )
            style_h2 = ParagraphStyle(
                'PdfH2',
                parent=styles['Heading2'],
                fontName='Helvetica-Bold',
                fontSize=12,
                leading=16,
                textColor=colors.HexColor('#0F172A'),
                spaceBefore=12,
                spaceAfter=4
            )
            style_body = ParagraphStyle(
                'PdfBody',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=9.5,
                leading=13.5,
                textColor=colors.HexColor('#334155'),
                alignment=TA_JUSTIFY,
                spaceAfter=6
            )
            style_bullet = ParagraphStyle(
                'PdfBullet',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=9.5,
                leading=13.5,
                textColor=colors.HexColor('#334155'),
                alignment=TA_JUSTIFY,
                leftIndent=15,
                spaceAfter=4
            )
            style_caption = ParagraphStyle(
                'PdfCaption',
                parent=styles['Normal'],
                fontName='Helvetica-Oblique',
                fontSize=8.5,
                leading=11,
                textColor=colors.HexColor('#475569'),
                alignment=TA_LEFT,
                spaceAfter=12
            )

            story = []
            
            # PORTADA PDF
            story.append(Spacer(1, 30))
            story.append(Paragraph(f"<b>{metadatos['empresa'].upper()} — {metadatos['unidad'].upper()}</b>", style_subtitle))
            story.append(Spacer(1, 10))
            story.append(Paragraph(f"INFORME TÉCNICO DE MANTENIMIENTO<br/>{metadatos['tipo'].upper()}", style_title))
            story.append(Paragraph(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}", style_subtitle))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1D4ED8'), spaceAfter=15))
            
            # Tabla de Metadatos PDF
            t_data = [
                [Paragraph("<b>Compañía / Planta:</b>", style_body), Paragraph(metadatos["empresa"], style_body)],
                [Paragraph("<b>Código de Documento:</b>", style_body), Paragraph(metadatos["codigo"], style_body)],
                [Paragraph("<b>Fecha del Evento:</b>", style_body), Paragraph(metadatos["fecha"], style_body)],
                [Paragraph("<b>Horas de Operación (TTSO):</b>", style_body), Paragraph(f"{metadatos['horas']} hrs", style_body)],
                [Paragraph("<b>Elaborado por:</b>", style_body), Paragraph(f"{metadatos['elaborador']} ({metadatos['cargo']})", style_body)],
                [Paragraph("<b>Aprobado por:</b>", style_body), Paragraph(metadatos["revisador"], style_body)],
                [Paragraph("<b>Clasificación:</b>", style_body), Paragraph(metadatos["confidencialidad"], style_body)]
            ]
            t_pdf = Table(t_data, colWidths=[150, 354])
            t_pdf.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 5),
            ]))
            story.append(t_pdf)
            story.append(PageBreak())

            # CUERPO DEL INFORME (PÁRRAFOS JUSTIFICADOS Y VIÑETAS SELECTIVAS)
            for line in contenido_md.split('\n'):
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

            # ANEXO FOTOGRÁFICO EN PDF (CENTRADOS CON CAPTION APA)
            if metadatos["imagenes"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas (Estándar APA 7.ª Ed.)</b>", style_h2))
                story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1D4ED8'), spaceAfter=10))
                
                for item in metadatos["imagenes"]:
                    img_bytes = io.BytesIO()
                    rgb_img = item["objeto_imagen"].convert('RGB')
                    rgb_img.save(img_bytes, format='JPEG', quality=90)
                    img_bytes.seek(0)
                    
                    img_flow = RLImage(img_bytes, width=380, height=260)
                    img_flow.hAlign = 'CENTER'
                    story.append(img_flow)
                    story.append(Spacer(1, 6))
                    
                    caption_str = f"<b>Figura {item['indice']}.</b> <i>{item['descripcion']}</i><br/><b>Fuente:</b> Elaboración propia."
                    story.append(Paragraph(caption_str, style_caption))
                    story.append(Spacer(1, 10))

            # Compilación final del PDF pasando la empresa y código al canvas institucional
            doc_pdf.build(
                story, 
                canvasmaker=lambda *args, **kwargs: type(
                    'CanvasPersonalizado', 
                    (CanvasInstitucional,), 
                    {'empresa_meta': metadatos["empresa"], 'codigo_doc': metadatos["codigo"]}
                )(*args, **kwargs)
            )
            buffer_pdf.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .pdf",
                data=buffer_pdf,
                file_name=f"Informe_Tecnico_{metadatos['tag']}_{metadatos['fecha']}.pdf",
                mime="application/pdf"
            )
