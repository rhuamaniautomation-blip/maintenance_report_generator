import io
import os
from datetime import datetime
import streamlit as st
from PIL import Image
import pandas as pd
from google import genai

# Generación de Word
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Generador Inteligente de Informes de Mantenimiento",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para ingeniería
st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        color: #1E3A8A;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #4B5563;
        margin-bottom: 2rem;
    }
    .stAlert {
        border-radius: 0.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# INICIALIZACIÓN DE API GEMINI
# ==========================================
def get_gemini_client():
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "gemini_api" in st.session_state:
        api_key = st.session_state["gemini_api"]
    
    if not api_key:
        return None
    return genai.Client(api_key=api_key)

# ==========================================
# INTERFAZ - BARRA LATERAL
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=70)
    st.markdown("### Configuración del Sistema")
    
    # Clave API en barra lateral si no está en secrets
    if "GEMINI_API_KEY" not in st.secrets:
        api_input = st.text_input("Clave API de Gemini:", type="password")
        if api_input:
            st.session_state["gemini_api"] = api_input
            st.success("Clave configurada temporalmente.")
    
    st.markdown("---")
    st.markdown("#### Parámetros del Reporte")
    tipo_informe = st.selectbox(
        "Tipo de Informe Ejecutivo",
        [
            "Informe Ejecutivo de Mantenimiento",
            "Informe de Falla",
            "Informe de Investigación",
            "Informe de Integridad Mecánica",
            "Informe NDT",
            "Informe de Causa Raíz",
            "Informe de Incidente Operacional",
            "Informe de Mantenimiento Correctivo",
            "Informe de Mantenimiento Preventivo",
            "Informe de Confiabilidad"
        ]
    )
    
    equipo_tag = st.text_input("Tag / Nombre del Equipo", "P-101A / Bomba Centrífuga Principal")
    sistema_planta = st.text_input("Sistema / Área", "Unidad de Hidrotratamiento (U-200)")
    autor_reporte = st.text_input("Elaborado por (Ingeniero)", "Ing. Mantenimiento Senior")
    fecha_reporte = st.date_input("Fecha del Evento", datetime.now())

# ==========================================
# CUERPO PRINCIPAL DE LA APLICACIÓN
# ==========================================
st.markdown('<div class="main-header">Generador Inteligente de Informes Ejecutivos de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sub-header">Plataforma experta para la documentación, análisis de confiabilidad y generación normativa de reportes técnicos.</div>', unsafe_allow_html=True)

# Pestañas de entrada de datos
tab1, tab2, tab3 = st.tabs(["📝 Datos del Evento", "📷 Evidencias Fotográficas", "📊 Generación & Exportación"])

with tab1:
    st.markdown("### 1. Información Técnica de Entrada")
    col1, col2 = st.columns(2)
    
    with col1:
        contexto_evento = st.text_area(
            "Contexto del Evento",
            placeholder="Ej: Operación en régimen nominal a 85% de carga cuando se detectan vibraciones anómalas en el rodamiento lado libre...",
            height=130
        )
        descripcion_inicial = st.text_area(
            "Descripción Inicial del Incidente",
            placeholder="Ej: Disparo por alta vibración enclavado por sistema interlock a las 03:45 hrs...",
            height=130
        )
        hallazgos_raw = st.text_area(
            "Hallazgos Encontrados",
            placeholder="Ej: Temperatura superficial elevada (88°C), juego axial excesivo, trazas de contaminación por humedad en lubricante...",
            height=130
        )
        
    with col2:
        trabajos_realizados = st.text_area(
            "Trabajos Realizados",
            placeholder="Ej: Desmontaje de acoplamiento, inspección dimensional de flecha, reemplazo de rodamientos SKF 6316 y rellenado de grasa sintética...",
            height=130
        )
        info_adicional = st.text_area(
            "Información Adicional Opcional",
            placeholder="Ej: Historial de mantenimientos previos, horas de operación acumuladas desde el último overhaul (14,200 hrs)...",
            height=130
        )

with tab2:
    st.markdown("### 2. Administración de Fotografías y Evidencias")
    st.info("⚠️ **Requisito obligatorio:** Cada imagen cargada debe contar con su descripción técnica detallada para cumplir con el estándar APA 7.ª Edición.")
    
    uploaded_files = st.file_uploader(
        "Cargar Fotografías (Formatos: PNG, JPG, JPEG)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )
    
    image_data_list = []
    if uploaded_files:
        st.markdown(f"**Total de imágenes cargadas:** {len(uploaded_files)}")
        for idx, uploaded_file in enumerate(uploaded_files):
            col_img, col_desc = st.columns([1, 2])
            with col_img:
                image = Image.open(uploaded_file)
                st.image(image, caption=f"Imagen Original {idx+1}", use_container_width=True)
            with col_desc:
                desc_label = st.text_input(
                    f"Descripción obligatoria para Figura {idx+1}:",
                    key=f"img_desc_{idx}",
                    placeholder="Ej: Evidencia de pitting avanzado en pista interna del rodamiento radial..."
                )
                image_data_list.append({
                    "file": uploaded_file,
                    "image_obj": image,
                    "description": desc_label,
                    "index": idx + 1
                })
    else:
        st.warning("No se han cargado imágenes. El informe requiere al menos una evidencia fotográfica con descripción para proceder.")

with tab3:
    st.markdown("### 3. Consolidación y Generación con Inteligencia Artificial")
    
    # Validaciones previas
    missing_descriptions = any(not item["description"].strip() for item in image_data_list)
    
    if not uploaded_files:
        st.error("❌ Debe cargar al menos una imagen en la pestaña anterior.")
    elif missing_descriptions:
        st.error("❌ Todas las imágenes cargadas deben tener una descripción obligatoria asignada.")
    else:
        st.success("✅ Todas las validaciones de evidencias fotográficas superadas con éxito.")
        
    if st.button("🚀 Generar Informe Ejecutivo con IA", type="primary", use_container_width=True):
        client = get_gemini_client()
        if not client:
            st.error("Configure su Clave API de Gemini en los secretos o en la barra lateral.")
        else:
            with st.spinner("🧠 Analizando parámetros, infiriendo modos de falla y redactando informe ejecutivo de ingeniería..."):
                
                # Construcción del Prompt Experto
                prompt_sistema = f"""
Actúa como un Ingeniero Senior de Mantenimiento, Ingeniero de Confiabilidad e Investigador de Incidentes Industriales con más de 25 años de experiencia.
Analiza el contexto suministrado y genera un Informe Ejecutivo de Mantenimiento de nivel corporativo para el equipo '{equipo_tag}' ubicado en '{sistema_planta}'.
El tipo de informe requerido es: '{tipo_informe}'.

La redacción debe ser técnica, profesional, humanizada, clara, ejecutiva, objetiva y trazable, equivalente a reportes presentados a Gerencia de Planta y Confiabilidad.
Cumple estrictamente con ISO 9001, ISO 55001, mejores prácticas de confiabilidad y formato APA 7ma edición.

Proporciona el informe completo estructurado exactamente con los siguientes 19 títulos en formato Markdown (h2 o h3):
1. Resumen Ejecutivo
2. Información General
3. Antecedentes
4. Descripción del Evento o problema identificado
5. Inspección Realizada
6. Hallazgos
7. Observaciones Técnicas
8. Trabajos Ejecutados
9. Evaluación Mecánica
10. Evaluación Eléctrica
11. Evaluación de Riesgos
12. Análisis de Causa Probable (inferir modos de falla, causas raíz técnicas)
13. Acciones Correctivas
14. Acciones Preventivas
15. Conclusiones
16. Recomendaciones
17. Estado Final del Equipo
18. Evidencias Fotográficas (Placeholder para insertar imágenes)
19. Referencias (Normativas ISO y buenas prácticas aplicadas)

Datos de entrada del usuario:
- Contexto: {contexto_evento}
- Descripción inicial: {descripcion_inicial}
- Hallazgos: {hallazgos_raw}
- Trabajos realizados: {trabajos_realizados}
- Información adicional: {info_adicional}
"""
                try:
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt_sistema
                    )
                    st.session_state["generated_report"] = response.text
                    st.session_state["image_data_list"] = image_data_list
                    st.success("¡Informe generado exitosamente!")
                except Exception as e:
                    st.error(f"Error al conectar con la API de Gemini: {e}")

# ==========================================
# VISUALIZACIÓN Y EXPORTACIÓN DEL INFORME
# ==========================================
if "generated_report" in st.session_state:
    st.markdown("---")
    st.markdown("## 📄 Vista Previa del Informe Técnico Generado")
    
    report_text = st.session_state["generated_report"]
    st.markdown(report_text)
    
    # Mostrar vista previa de imágenes numeradas estilo APA
    if "image_data_list" in st.session_state and st.session_state["image_data_list"]:
        st.markdown("### 📷 Evidencias Fotográficas Incorporadas (Formato APA)")
        for item in st.session_state["image_data_list"]:
            st.image(item["image_obj"], caption=f"Figura {item['index']}. {item['description']}\nFuente: Elaboración propia.", use_container_width=True)

    st.markdown("---")
    st.markdown("### 📥 Opciones de Exportación Profesional")
    
    col_exp1, col_exp2 = st.columns(2)
    
    # EXPORTACIÓN WORD (DOCX)
    with col_exp1:
        if st.button("📥 Exportar Documento Word (.docx)", use_container_width=True):
            doc = Document()
            
            # Estilos de portada y márgenes
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
            
            # Portada formal
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_title = p_title.add_run(f"INFORME TÉCNICO DE MANTENIMIENTO\n{tipo_informe.upper()}")
            run_title.font.size = Pt(18)
            run_title.font.bold = True
            run_title.font.color.rgb = RGBColor(30, 58, 138)
            
            doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_sub.add_run(f"Equipo: {equipo_tag}\nSistema: {sistema_planta}\nFecha: {fecha_reporte}\nElaborado por: {autor_reporte}")
            
            doc.add_page_break()
            
            # Contenido principal del informe
            doc.add_heading("Contenido del Informe", level=1)
            for line in report_text.split('\n'):
                if line.startswith("# "):
                    doc.add_heading(line.replace("# ", ""), level=1)
                elif line.startswith("## "):
                    doc.add_heading(line.replace("## ", ""), level=2)
                elif line.startswith("### "):
                    doc.add_heading(line.replace("### ", ""), level=3)
                else:
                    if line.strip():
                        doc.add_paragraph(line)
            
            # Inserción de imágenes con formato APA en Word
            if "image_data_list" in st.session_state and st.session_state["image_data_list"]:
                doc.add_heading("Evidencias Fotográficas", level=2)
                for item in st.session_state["image_data_list"]:
                    img_path = f"temp_img_{item['index']}.png"
                    item["image_obj"].save(img_path)
                    doc.add_picture(img_path, width=Inches(5.0))
                    p_cap = doc.add_paragraph()
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run_cap = p_cap.add_run(f"Figura {item['index']}. {item['description']}\nFuente: Elaboración propia.")
                    run_cap.font.size = Pt(10)
                    run_cap.font.italic = True
                    if os.path.exists(img_path):
                        os.remove(img_path)
            
            # Guardar archivo en memoria
            docx_io = io.BytesIO()
            doc.save(docx_io)
            docx_io.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .docx",
                data=docx_io,
                file_name=f"Informe_Mantenimiento_{equipo_tag.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # EXPORTACIÓN PDF (REPORTLAB)
    with col_exp2:
        if st.button("📥 Exportar Documento PDF (.pdf)", use_container_width=True):
            pdf_io = io.BytesIO()
            doc_pdf = SimpleDocTemplate(pdf_io, pagesize=letter, rightMargin=54, leftMargin=54, topMargin=54, bottomMargin=54)
            styles = getSampleStyleSheet()
            story = []
            
            # Estilos personalizados PDF
            title_style = ParagraphStyle(
                'CoverTitle',
                parent=styles['Heading1'],
                fontSize=20,
                textColor=colors.HexColor('#1E3A8A'),
                alignment=1,
                spaceAfter=20
            )
            body_style = ParagraphStyle(
                'ReportBody',
                parent=styles['Normal'],
                fontSize=10,
                leading=14,
                textColor=colors.HexColor('#1F2937'),
                spaceAfter=10
            )
            
            # Portada PDF
            story.append(Spacer(1, 100))
            story.append(Paragraph(f"<b>INFORME TÉCNICO DE MANTENIMIENTO</b><br/>{tipo_informe.upper()}", title_style))
            story.append(Spacer(1, 40))
            story.append(Paragraph(f"<b>Equipo:</b> {equipo_tag}", body_style))
            story.append(Paragraph(f"<b>Sistema:</b> {sistema_planta}", body_style))
            story.append(Paragraph(f"<b>Fecha:</b> {fecha_reporte}", body_style))
            story.append(Paragraph(f"<b>Elaborado por:</b> {autor_reporte}", body_style))
            story.append(PageBreak())
            
            # Contenido PDF
            for line in report_text.split('\n'):
                if line.startswith("# ") or line.startswith("## ") or line.startswith("### "):
                    clean_title = line.replace("#", "").strip()
                    story.append(Paragraph(f"<b>{clean_title}</b>", styles['Heading2']))
                    story.append(Spacer(1, 6))
                else:
                    if line.strip():
                        story.append(Paragraph(line, body_style))
            
            # Imágenes en PDF
            if "image_data_list" in st.session_state and st.session_state["image_data_list"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas</b>", styles['Heading2']))
                story.append(Spacer(1, 10))
                for item in st.session_state["image_data_list"]:
                    img_path = f"temp_pdf_img_{item['index']}.png"
                    item["image_obj"].save(img_path)
                    story.append(RLImage(img_path, width=400, height=300))
                    story.append(Spacer(1, 5))
                    story.append(Paragraph(f"<i>Figura {item['index']}. {item['description']}<br/>Fuente: Elaboración propia.</i>", body_style))
                    story.append(Spacer(1, 15))
                    if os.path.exists(img_path):
                        os.remove(img_path)
            
            doc_pdf.build(story)
            pdf_io.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Archivo .pdf",
                data=pdf_io,
                file_name=f"Informe_Mantenimiento_{equipo_tag.replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
