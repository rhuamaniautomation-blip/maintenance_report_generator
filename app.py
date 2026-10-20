"""
================================================================================
SISTEMA INSTITUCIONAL DE INFORMES TÉCNICOS DE MANTENIMIENTO (SITM)
Versión: 3.0.0 - Edición Técnico-Operativa
Desarrollado para: Entornos Industriales de Alta Confiabilidad
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
    page_title="SITM - Sistema de Informes Técnicos de Mantenimiento",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Corporativos y Modernos
st.markdown("""
    <style>
    .main {
        background-color: #F1F5F9;
    }
    .stAppHeader {
        background-color: rgba(255, 255, 255, 0.95);
    }
    .main-header {
        font-size: 2.1rem;
        color: #0F172A;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #475569;
        margin-bottom: 1.5rem;
    }
    .card-container {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
        border: 1px solid #E2E8F0;
        margin-bottom: 1.25rem;
    }
    .badge-tech {
        background-color: #DBEAFE;
        color: #1E40AF;
        padding: 0.2rem 0.6rem;
        border-radius: 9999px;
        font-size: 0.8rem;
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
# 3. CLASE CANVAS PERSONALIZADA PARA NUMERACIÓN PROFESIONAL DE PÁGINAS (PDF)
# ==============================================================================
class NumeredCanvas(canvas.Canvas):
    """
    Canvas avanzado que implementa un patrón de doble pasada para calcular 
    dinámicamente el número total de páginas e incluir encabezados y pies de página 
    institucionales consistentes.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))
        
        # Omitir en la portada (Página 1)
        if self._pageNumber > 1:
            # Encabezado superior
            self.drawString(54, 750, "SISTEMA INSTITUCIONAL DE MANTENIMIENTO — INFORME TÉCNICO")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(54, 742, 558, 742)
            
            # Pie de página inferior
            page_text = f"Página {self._pageNumber} de {page_count}"
            self.drawRightString(558, 36, page_text)
            self.drawString(54, 36, "DOCUMENTO TÉCNICO OFICIAL — USO INTERNO")
            self.line(54, 48, 558, 48)
            
        self.restoreState()

# ==============================================================================
# 4. BARRA LATERAL SIMPLIFICADA (ENFOQUE TÉCNICO)
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=60)
    st.markdown("### **Panel Técnico SITM**")
    st.markdown("*Generación Rápida de Reportes*")
    st.markdown("---")
    
    # Autenticación rápida si no está en secrets
    if "GEMINI_API_KEY" not in st.secrets:
        st.markdown("#### 🔑 Credenciales API")
        st.text_input("Gemini API Key:", type="password", key="gemini_api_key_input")
        st.markdown("---")

    st.markdown("#### 🏢 Datos Organizacionales")
    empresa_nombre = st.text_input("Empresa / Planta", "ORICA Industrial")
    unidad_operativa = st.text_input("Unidad / Área", "Unidad de Detonadores")
    
    tipo_informe = st.selectbox(
        "Tipo de Informe Técnico",
        [
            "Informe Ejecutivo de Mantenimiento",
            "Informe de Falla y Causa Raíz",
            "Informe de Investigación de Incidente",
            "Informe de Mantenimiento Correctivo",
            "Informe de Mantenimiento Preventivo",
            "Informe de Confiabilidad Operacional"
        ]
    )

    st.markdown("---")
    st.markdown("#### 👤 Responsable del Reporte")
    elaborado_por = st.text_input("Elaborado por", "Ing. Roger Huamani Torres")
    cargo_elaborador = st.text_input("Cargo", "Superintendente de Mantenimiento")
    revisado_por = st.text_input("Aprobado por", "Gerencia de Operaciones")
    fecha_evento = st.date_input("Fecha del Evento / Trabajo", date.today())

# ==============================================================================
# 5. INTERFAZ PRINCIPAL DE ENTRADA (OPTIMIZADA PARA TÉCNICOS)
# ==============================================================================
st.markdown('<div class="main-header">Generador Inteligente de Informes de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Herramienta optimizada para personal técnico: Redacte de forma sencilla, cargue sus evidencias y obtenga un informe ejecutivo estructurado al instante.</div>', unsafe_allow_html=True)

# Pestañas de flujo de trabajo simplificadas
tab_datos, tab_evidencias, tab_preview = st.tabs([
    "📝 1. Datos del Evento", 
    "📷 2. Evidencias Fotográficas", 
    "🚀 3. Vista Previa y Exportación"
])

# ------------------------------------------------------------------------------
# PESTAÑA 1: DATOS Y CONTEXTO DEL EVENTO
# ------------------------------------------------------------------------------
with tab_datos:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Identificación del Activo")
    col_e1, col_e2, col_e3 = st.columns(3)
    with col_e1:
        equipo_tag = st.text_input("Tag del Equipo", "230-1", help="Identificador único del activo en planta.")
    with col_e2:
        equipo_nombre = st.text_input("Nombre del Equipo", "Máquina de Carga de Retardos")
    with col_e3:
        horas_operacion = st.number_input("Horas de Operación (TTSO)", min_value=0, value=14200, step=50)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Redacción Técnica del Suceso")
    st.info("💡 **Consejo:** Ingrese la información de forma directa. La Inteligencia Artificial se encargará de redactarla de manera concisa, ejecutiva y profesional, conservando sus viñetas y datos clave.")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        contexto_evento = st.text_area(
            "Contexto del Evento / Operación",
            placeholder="Ej: Operación normal en línea de detonadores 500ms en 30mm...",
            height=130
        )
        descripcion_incidente = st.text_area(
            "Descripción Inicial del Incidente",
            placeholder="Ej: Detección de ruido anómalo y humo en Estación ST23, parada de emergencia...",
            height=130
        )
        hallazgos_raw = st.text_area(
            "Hallazgos Encontrados",
            placeholder="Ej: Desgaste en O-rings de ST21 y ST22, acumulación de polución en escobillas, rotura en tanque ST23...",
            height=130
        )
    with col_c2:
        trabajos_realizados = st.text_area(
            "Trabajos Realizados",
            placeholder="Ej: Sustitución completa de O-rings, reemplazo de tanque ST23, revisión y limpieza de escobillas, pruebas operativas...",
            height=130
        )
        repuestos_materiales = st.text_area(
            "Repuestos y Materiales Empleados",
            placeholder="Ej: 2x O-rings elastoméricos de alta resistencia, 1x Tanque superior ST23...",
            height=130
        )
        info_adicional = st.text_area(
            "Información Adicional (Opcional)",
            placeholder="Ej: Historial de intervenciones recientes, recomendaciones operativas...",
            height=130
        )
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 2: EVIDENCIAS FOTOGRÁFICAS
# ------------------------------------------------------------------------------
with tab_evidencias:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📷 Carga y Descripción de Evidencias Fotográficas")
    st.caption("Cargue todas las imágenes que necesite. Cada imagen debe tener obligatoriamente su descripción técnica para ser incorporada automáticamente de forma centrada y con formato APA en el informe.")
    
    uploaded_files = st.file_uploader(
        "Seleccione o arrastre sus fotografías (Formatos: JPG, JPEG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    
    lista_imagenes = []
    if uploaded_files:
        st.success(f"Total de imágenes cargadas: **{len(uploaded_files)}**")
        st.markdown("---")
        for idx, archivo in enumerate(uploaded_files):
            col_img, col_txt = st.columns([1, 2])
            with col_img:
                img_obj = Image.open(archivo)
                st.image(img_obj, caption=f"Evidencia N.° {idx+1}", use_container_width=True)
            with col_txt:
                desc_img = st.text_input(
                    f"Descripción técnica para Figura {idx+1}:",
                    key=f"desc_img_input_{idx}",
                    placeholder="Ej: Condición post-incidente de la Estación ST23 evidenciando daños estructurales..."
                )
                lista_imagenes.append({
                    "nombre_archivo": archivo.name,
                    "objeto_imagen": img_obj,
                    "descripcion": desc_img,
                    "indice": idx + 1
                })
    else:
        st.warning("⚠️ No se han seleccionado imágenes. Se requiere al menos una evidencia fotográfica con su descripción.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 3: VISTA PREVIA Y GENERACIÓN
# ------------------------------------------------------------------------------
with tab_preview:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🚀 Generación Automática del Informe Ejecutivo")
    
    tiene_imagenes = len(lista_imagenes) > 0
    falta_descripcion = any(not img["descripcion"].strip() for img in lista_imagenes) if tiene_imagenes else False
    
    if not tiene_imagenes:
        st.error("❌ Falta requisito: Debe cargar al menos una imagen en la pestaña anterior.")
    elif falta_descripcion:
        st.error("❌ Falta requisito: Todas las imágenes cargadas deben poseer una descripción técnica obligatoria.")
    else:
        st.success("✅ Verificaciones completadas. El sistema está listo para generar el informe técnico.")
        
    btn_generar = st.button("✨ Generar Informe Técnico con Inteligencia Artificial", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_generar and tiene_imagenes and not falta_descripcion:
        cliente_ai = inicializar_cliente_gemini()
        if not cliente_ai:
            st.error("❌ Error de Autenticación: Configure su clave API de Gemini en Streamlit Secrets o en la barra lateral.")
        else:
            with st.spinner("🔄 Procesando datos, resumiendo hallazgos y estructurando el informe técnico ejecutivo..."):
                
                evidencias_txt = "\n".join([f"- Figura {item['indice']}: {item['descripcion']}" for item in lista_imagenes])
                
                # Prompt estricto para evitar exageraciones y mantener concisión y viñetas
                prompt_ia = f"""
Actúa como un Ingeniero Senior de Mantenimiento con 25 años de experiencia. Redacta un informe técnico conciso, directo, ejecutivo y profesional (sin texto redundante ni exageraciones retóricas).
Basate estrictamente en los datos suministrados. Si se proporcionan viñetas o listas, conservalas en el resultado.

Parámetros del Reporte:
- Empresa: {empresa_nombre}
- Unidad: {unidad_operativa}
- Tipo: {tipo_informe}
- Activo: {equipo_tag} - {equipo_nombre}
- Horas de Operación: {horas_operacion} hrs
- Elaborado por: {elaborado_por} ({cargo_elaborador})

Datos proporcionados por el técnico:
- Contexto: {contexto_evento}
- Descripción del incidente: {descripcion_incidente}
- Hallazgos: {hallazgos_raw}
- Trabajos ejecutados: {trabajos_realizados}
- Repuestos y materiales: {repuestos_materiales}
- Información adicional: {info_adicional}
- Evidencias fotográficas asociadas:
{evidencias_txt}

ESTRUCTURA OBLIGATORIA (Desarrolla cada sección de forma clara, resumida y técnica, usando viñetas donde corresponda con asteriscos `*`):
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
                    st.success("¡Informe técnico generado y validado con éxito!")
                except Exception as e:
                    st.error(f"Error al conectar con la API de Gemini: {e}")

# ==============================================================================
# 6. VISUALIZACIÓN Y MOTOR DE EXPORTACIÓN (WORD Y PDF PROFESIONALES)
# ==============================================================================
if "texto_informe" in st.session_state:
    st.markdown("---")
    st.markdown("## 📄 Vista Previa del Informe Técnico")
    
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
        if st.button("📥 Descargar Formato Word (.docx)", use_container_width=True):
            doc = Document()
            
            # Márgenes de 1 pulgada (2.54 cm)
            for sec in doc.sections:
                sec.top_margin = Inches(1)
                sec.bottom_margin = Inches(1)
                sec.left_margin = Inches(1)
                sec.right_margin = Inches(1)

            # Portada Institucional Word
            p_sup = doc.add_paragraph()
            p_sup.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_sup = p_sup.add_run(f"{metadatos['empresa']} | {metadatos['unidad']}")
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
            
            # Tabla de metadatos corporativos
            t_meta = doc.add_table(rows=5, cols=2)
            t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
            datos_t = [
                ("Fecha del Evento:", metadatos["fecha"]),
                ("Horas de Operación (TTSO):", f"{metadatos['horas']} hrs"),
                ("Elaborado por:", f"{metadatos['elaborador']} ({metadatos['cargo']})"),
                ("Aprobado por:", metadatos["revisador"]),
                ("Estándar Normativo:", "ISO 55001 / ISO 9001 / APA 7.ª Ed.")
            ]
            for idx_t, (lbl, val) in enumerate(datos_t):
                c_cells = t_meta.rows[idx_t].cells
                c_cells[0].text = lbl
                c_cells[0].paragraphs[0].runs[0].font.bold = True
                c_cells[1].text = val
            
            doc.add_page_break()
            
            # Procesamiento de líneas y preservación de viñetas
            for line in contenido_md.split('\n'):
                line_text = line.strip()
                if line_text.startswith("## "):
                    h = doc.add_heading(line_text.replace("## ", ""), level=2)
                    h.runs[0].font.color.rgb = RGBColor(15, 23, 42)
                elif line_text.startswith("# "):
                    h = doc.add_heading(line_text.replace("# ", ""), level=1)
                    h.runs[0].font.color.rgb = RGBColor(30, 64, 175)
                elif line_text.startswith("* ") or line_text.startswith("- "):
                    # Viñeta preservada
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
        if st.button("📥 Descargar Formato PDF (.pdf)", use_container_width=True):
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
            
            # Estilos tipográficos profesionales y justificados
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
            story.append(Spacer(1, 40))
            story.append(Paragraph(f"<b>{metadatos['empresa'].upper()} — {metadatos['unidad'].upper()}</b>", style_subtitle))
            story.append(Spacer(1, 10))
            story.append(Paragraph(f"INFORME TÉCNICO DE MANTENIMIENTO<br/>{metadatos['tipo'].upper()}", style_title))
            story.append(Paragraph(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}", style_subtitle))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1D4ED8'), spaceAfter=20))
            
            # Tabla de Metadatos PDF
            t_data = [
                [Paragraph("<b>Fecha del Evento:</b>", style_body), Paragraph(metadatos["fecha"], style_body)],
                [Paragraph("<b>Horas de Operación (TTSO):</b>", style_body), Paragraph(f"{metadatos['horas']} hrs", style_body)],
                [Paragraph("<b>Elaborado por:</b>", style_body), Paragraph(f"{metadatos['elaborador']} ({metadatos['cargo']})", style_body)],
                [Paragraph("<b>Aprobado por:</b>", style_body), Paragraph(metadatos["revisador"], style_body)],
                [Paragraph("<b>Normativa:</b>", style_body), Paragraph("ISO 55001 / ISO 9001 / APA 7.ª Ed.", style_body)]
            ]
            t_pdf = Table(t_data, colWidths=[150, 354])
            t_pdf.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(t_pdf)
            story.append(PageBreak())

            # CUERPO DEL INFORME (PROCESAMIENTO DE LÍNEAS Y VIÑETAS)
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

            # ANEXO FOTOGRÁFICO EN PDF (CENTRADOS Y CON CAPTION APA)
            if metadatos["imagenes"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas (Estándar APA 7.ª Ed.)</b>", style_h2))
                story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1D4ED8'), spaceAfter=10))
                
                for item in metadatos["imagenes"]:
                    img_bytes = io.BytesIO()
                    rgb_img = item["objeto_imagen"].convert('RGB')
                    rgb_img.save(img_bytes, format='JPEG', quality=90)
                    img_bytes.seek(0)
                    
                    # Imagen centrada usando un Table o flujo de ReportLab
                    img_flow = RLImage(img_bytes, width=380, height=260)
                    img_flow.hAlign = 'CENTER'
                    story.append(img_flow)
                    story.append(Spacer(1, 6))
                    
                    caption_str = f"<b>Figura {item['indice']}.</b> <i>{item['descripcion']}</i><br/><b>Fuente:</b> Elaboración propia."
                    story.append(Paragraph(caption_str, style_caption))
                    story.append(Spacer(1, 10))

            # Compilación final del PDF con paginación numérica automática
            doc_pdf.build(story, canvasmaker=NumeredCanvas)
            buffer_pdf.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .pdf",
                data=buffer_pdf,
                file_name=f"Informe_Tecnico_{metadatos['tag']}_{metadatos['fecha']}.pdf",
                mime="application/pdf"
            )
