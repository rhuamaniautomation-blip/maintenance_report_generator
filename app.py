"""
================================================================================
SISTEMA INSTITUCIONAL DE INFORMES TÉCNICOS DE MANTENIMIENTO (SITM)
Versión: 5.0.0 - Edición Ejecutiva Gerencial & Corrección Tipográfica Avanzada
Desarrollado para: Informes Gerenciales de Alta Confiabilidad Industrial
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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
    page_title="SITM - Sistema Gerencial de Informes de Mantenimiento",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Corporativos y Modernos
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
    .badge-exec {
        background-color: #EEF2FF;
        color: #4338CA;
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
    """Valida y retorna la instancia oficial del cliente Google GenAI."""
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "gemini_api_key_input" in st.session_state and st.session_state["gemini_api_key_input"]:
        api_key = st.session_state["gemini_api_key_input"]
    
    if not api_key:
        return None
    return genai.Client(api_key=api_key)

# ==============================================================================
# 3. UTILIDADES DE PARSEO Y LIMPIEZA DE MARKDOWN (SOLUCIÓN DE FORMATO)
# ==============================================================================
def limpiar_markdown_para_docx(doc_instance, texto_linea):
    """
    Procesa una línea de texto con marcas Markdown (**negrita**) y la añade 
    correctamente a un párrafo de Word sin mostrar asteriscos literales.
    """
    p = doc_instance.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(4)
    
    # Patrón para encontrar texto entre **
    partes = re.split(r'(\*\*.*?\*\*)', texto_linea)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**'):
            run = p.add_run(parte[2:-2])
            run.font.bold = True
            run.font.name = 'Arial'
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(15, 23, 42)
        else:
            if parte:
                run = p.add_run(parte)
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(51, 65, 85)
    return p

def limpiar_markdown_para_pdf(texto_linea):
    """
    Convierte marcas Markdown (**negrita**) a etiquetas HTML nativas de ReportLab (<b>...</b>)
    y limpia asteriscos sueltos para evitar defectos visuales en el PDF.
    """
    # Reemplazar **texto** por <b>texto</b>
    texto_convertido = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', texto_linea)
    # Limpiar cualquier asterisco sobrante al inicio de líneas que no sean viñetas
    texto_convertido = texto_convertido.replace('*', '')
    return texto_convertido

# ==============================================================================
# 4. CANVAS PERSONALIZADO PARA NUMERACIÓN Y ENCABEZADOS PDF
# ==============================================================================
class CanvasGerencial(canvas.Canvas):
    """Canvas de doble pasada para paginación dinámica 'Página X de Y' y membretes."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self.empresa_meta = "COMPAÑÍA INDUSTRIAL"
        self.codigo_doc = "SITM-GER-01"

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
        
        # Omitir en la Portada (Página 1)
        if self._pageNumber > 1:
            self.drawString(54, 750, f"{self.empresa_meta.upper()} — INFORME EJECUTIVO DE MANTENIMIENTO")
            self.drawRightString(558, 750, f"REF: {self.codigo_doc}")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.75)
            self.line(54, 742, 558, 742)
            
            self.setFont("Helvetica", 8)
            page_text = f"Página {self._pageNumber} de {page_count}"
            self.drawRightString(558, 36, page_text)
            self.drawString(54, 36, "REPORTE GERENCIAL OFICIAL — TOMA DE DECISIONES")
            self.line(54, 48, 558, 48)
            
        self.restoreState()

# ==============================================================================
# 5. BARRA LATERAL - CONFIGURACIÓN GERENCIAL
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=65)
    st.markdown("### **Panel Gerencial SITM**")
    st.markdown("*Configuración del Reporte Ejecutivo*")
    st.markdown("---")
    
    if "GEMINI_API_KEY" not in st.secrets:
        st.markdown("#### 🔑 Credencial API Gemini")
        st.text_input("API Key:", type="password", key="gemini_api_key_input")
        st.markdown("---")

    st.markdown("#### 🏢 Datos Organizacionales")
    empresa_nombre = st.text_input("Compañía / Planta", "ORICA Industrial")
    area_proceso = st.text_input("Área o Proceso", "Unidad de Detonadores - Línea 2")
    codigo_documento = st.text_input("Código / Referencia OT", "OT-98421-2026")
    
    st.markdown("---")
    st.markdown("#### 📋 Datos de Intervención")
    tipo_mantenimiento = st.selectbox(
        "Tipo de Mantenimiento",
        ["Correctivo", "Preventivo", "Predictivo", "Emergencia"]
    )
    fecha_intervencion = st.date_input("Fecha de la Intervención", date.today())
    fecha_emision = st.date_input("Fecha de Emisión del Informe", date.today())

    st.markdown("---")
    st.markdown("#### 👤 Responsables")
    responsables_ejecucion = st.text_area("Responsables de Ejecución (Nombres)", "Ing. Carlos Mendoza, Téc. Mario Ruiz")
    responsable_informe = st.text_input("Responsable del Informe", "Ing. Roger Huamani Torres (Superintendente)")

# ==============================================================================
# 6. INTERFAZ PRINCIPAL DE ENTRADA
# ==============================================================================
st.markdown('<div class="main-header">Generador Ejecutivo de Informes de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Plataforma orientada a Gerencia: Transforme la información técnica de campo en un reporte ejecutivo claro, estructurado, sin exageraciones y enfocado en impactos, riesgos y costos.</div>', unsafe_allow_html=True)

tab_datos, tab_plan, tab_evidencias, tab_preview = st.tabs([
    "📝 1. Datos y Contexto Técnico", 
    "📊 2. Plan de Acción", 
    "📷 3. Evidencias Fotográficas", 
    "🚀 4. Vista Previa y Exportación"
])

# ------------------------------------------------------------------------------
# PESTAÑA 1: DATOS Y CONTEXTO TÉCNICO DE LOS 12 APARTADOS
# ------------------------------------------------------------------------------
with tab_datos:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Identificación del Activo")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        equipo_intervenido = st.text_input("Equipo Intervenido", "Máquina de Carga de Retardos")
    with col_e2:
        codigo_tag = st.text_input("Código o TAG del Equipo", "230-1")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Detalle Técnico para el Reporte Gerencial")
    st.info("💡 **Indicación:** Ingrese la información de campo. La IA generará explicaciones técnicas profesionales estructuradas en párrafos formales y viñetas estrictamente donde corresponda, evitando repeticiones y lenguaje excesivamente técnico sin contexto gerencial.")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        condicion_reportada = st.text_area("Condición Reportada / Síntoma Inicial", "Operador reporta ruido anómalo y presencia de humo en Estación ST23 durante producción de 500ms.", height=110)
        condicion_encontrada = st.text_area("Condición Encontrada (Real)", "Degradación severa de sellos O-ring en ST21/ST22, polución en escobillas y rotura por sobrepresión en tanque ST23.", height=110)
        diagnostico_realizado = st.text_area("Diagnóstico Realizado", "Falla de contención por desgaste de elastómeros, migración de material de retardo y deflagración localizada en ST23.", height=110)
        alcance_trabajos = st.text_area("Alcance de Trabajos Ejecutados", "Desmontaje, reemplazo de O-rings, sustitución de tanque ST23, limpieza de escobillas y pruebas operativas.", height=110)
        
    with col_c2:
        impacto_operativo = st.text_area("Impacto Operacional (Horas indisponibilidad, producción)", "14 horas de indisponibilidad. Afectación de 1,200 unidades de detonadores programadas para el turno.", height=110)
        impactos_varios = st.text_area("Impactos (Seguridad, Ambiental, Económico Estimado)", "Seguridad: Activación de protocolos de evacuación sin lesiones. Económico: Costo estimado de parada y repuestos en $4,500 USD.", height=110)
        causa_raiz_info = st.text_area("Análisis de Causa Raíz (Inmediata, Básica, Raíz y Metodología)", "Inmediata: Rotura de tanque ST23. Básica: Envejecimiento de O-rings. Raíz: Ciclo de reemplazo preventivo superado. Metodología: 5 Porqués.", height=110)
        recursos_costos = st.text_area("Recursos y Costos (Personal, Materiales, Costos USD)", "Personal: 3 técnicos mecánicos y 1 supervisor (6 hrs). Materiales: O-rings y tanque ST23. Costo Total: $4,500 USD.", height=110)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Estado, Riesgos y Conclusiones Gerenciales")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        estado_equipo = st.selectbox("Estado Actual del Equipo", ["Operativo", "Operativo con restricciones", "Fuera de servicio", "En observación o monitoreo"])
        nivel_riesgo = st.selectbox("Nivel de Riesgo Actual", ["Bajo", "Medio", "Alto", "Crítico"])
        restricciones_op = st.text_input("Restricciones Operacionales", "Operación vigilada al 80% de capacidad durante las primeras 48 horas.")
        recomendaciones_inmediatas = st.text_area("Recomendaciones Inmediatas", "Inspección termográfica a las 24 hrs de operación continua y monitoreo de vibraciones.", height=90)
    with col_s2:
        riesgos_remanentes = st.text_area("Riesgos Pendientes y Consecuencias Potenciales", "Posible fatiga en componentes adyacentes si no se ajusta la frecuencia del preventivo de sellos.", height=90)
        conclusion_gerencial = st.text_area("Conclusión Ejecutiva y Beneficio Esperado", "Problema cerrado satisfactoriamente. Se restablece la confiabilidad de la línea y se previene la recurrencia mediante ingeniería de sellos.", height=90)
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 2: PLAN DE ACCIÓN Y SEGUIMIENTO (TABLA DINÁMICA)
# ------------------------------------------------------------------------------
with tab_plan:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📊 Plan de Acción y Seguimiento Gerencial")
    st.info("Configure las acciones correctivas o preventivas derivadas del evento para incluirlas en la tabla del informe.")
    
    if "acciones_plan" not in st.session_state:
        st.session_state["acciones_plan"] = [
            {"accion": "Actualizar pautas de mantenimiento preventivo para sellos O-ring en serie 230.", "responsable": "Ing. Carlos Mendoza", "fecha": "2026-11-15", "estado": "Pendiente"},
            {"accion": "Capacitar a operadores en detección temprana de anomalías acústicas.", "responsable": "Sup. Mario Ruiz", "fecha": "2026-11-20", "estado": "En proceso"}
        ]
        
    col_a1, col_a2, col_a3, col_a4 = st.columns([3, 2, 2, 2])
    with col_a1:
        nueva_accion = st.text_input("Acción Propuesta", placeholder="Ej: Revisar especificación de materiales...")
    with col_a2:
        nuevo_resp = st.text_input("Responsable", placeholder="Ej: Ing. Mantenimiento")
    with col_a3:
        nueva_fecha = st.date_input("Fecha Compromiso", date.today())
    with col_a4:
        nuevo_estado = st.selectbox("Estado", ["Pendiente", "En proceso", "Completado"])
        
    if st.button("➕ Agregar Acción al Plan"):
        if nueva_accion:
            st.session_state["acciones_plan"].append({
                "accion": nueva_accion,
                "responsable": nuevo_resp,
                "fecha": str(nueva_fecha),
                "estado": nuevo_estado
            })
            st.success("Acción agregada correctamente.")
            
    # Mostrar tabla actual
    if st.session_state["acciones_plan"]:
        df_acciones = pd.DataFrame(st.session_state["acciones_plan"]) if 'pd' in globals() else None
        st.table(st.session_state["acciones_plan"])
        if st.button("🗑️ Limpiar Plan de Acción"):
            st.session_state["acciones_plan"] = []
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 3: EVIDENCIAS FOTOGRÁFICAS
# ------------------------------------------------------------------------------
with tab_evidencias:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📷 Carga de Evidencias Fotográficas")
    st.caption("Cargue las fotografías necesarias. Indique el estado de cada una (Antes, Durante, Después) y su respectiva descripción técnica.")
    
    uploaded_files = st.file_uploader(
        "Seleccione o arrastre imágenes (JPG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    
    lista_imagenes = []
    if uploaded_files:
        st.success(f"Imágenes cargadas: **{len(uploaded_files)}**")
        st.markdown("---")
        for idx, archivo in enumerate(uploaded_files):
            col_img, col_txt1, col_txt2 = st.columns([1, 1.5, 1.5])
            with col_img:
                img_obj = Image.open(archivo)
                st.image(img_obj, caption=f"Foto {idx+1}", use_container_width=True)
            with col_txt1:
                desc_img = st.text_input(f"Descripción Foto {idx+1}:", key=f"desc_{idx}", placeholder="Ej: Condición inicial de la Estación ST23...")
            with col_txt2:
                estado_img = st.selectbox(f"Estado Foto {idx+1}:", ["Antes", "Durante", "Después"], key=f"estado_{idx}")
                
            lista_imagenes.append({
                "nombre": archivo.name,
                "objeto": img_obj,
                "descripcion": desc_img,
                "estado": estado_img,
                "indice": idx + 1
            })
    else:
        st.warning("⚠️ No se han cargado imágenes. Si no adjunta ninguna, el informe indicará 'No se adjuntó evidencia fotográfica.'")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 4: VISTA PREVIA Y GENERACIÓN DEL INFORME GERENCIAL
# ------------------------------------------------------------------------------
with tab_preview:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🚀 Consolidación del Informe Gerencial")
    
    btn_generar = st.button("✨ Generar Informe Ejecutivo con IA", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_generar:
        cliente_ai = inicializar_cliente_gemini()
        if not cliente_ai:
            st.error("❌ Error de Autenticación: Configure su clave API de Gemini en Streamlit Secrets o en la barra lateral.")
        else:
            with st.spinner("🔄 Procesando datos y estructurando el informe gerencial en los 12 apartados obligatorios..."):
                
                evidencias_str = "\n".join([f"- Foto {item['indice']} ({item['estado']}): {item['descripcion']}" for item in lista_imagenes]) if lista_imagenes else "No se adjuntó evidencia fotográfica."
                acciones_str = "\n".join([f"- [ ] {a['accion']} | Responsable: {a['responsable']} | Fecha: {a['fecha']} | Estado: {a['estado']}" for a in st.session_state.get("acciones_plan", [])])
                
                # Prompt estricto y profesional para los 12 apartados gerenciales
                prompt_gerencial = f"""
Actúa como un Especialista Senior en Gestión de Mantenimiento Industrial, Confiabilidad y Reportes Ejecutivos para Gerencia.
Genera un INFORME EJECUTIVO DE MANTENIMIENTO profesional, estructurado rigurosamente en los 12 apartados solicitados.

DIRECTRICES DE REDACCIÓN:
- Lenguaje ejecutivo, claro, preciso, objetivo y basado en hechos.
- Evita exageraciones; explicaciones técnicas equilibradas enfocadas en impacto operacional, riesgos, costos y decisiones.
- Redacta principalmente en formato de PÁRRAFOS formales y bien desarrollados. Utiliza viñetas (con asteriscos `*`) UNICAMENTE cuando sea estrictamente necesario para aportar claridad (ej. listado de recursos o impactos). NO satures el texto con viñetas innecesarias.
- Si algún dato no fue proporcionado, intégralo bajo una sección de "Información no disponible".

DATOS DE ENTRADA:
- Compañía: {empresa_nombre}
- Área o Proceso: {area_proceso}
- Equipo Intervenido: {equipo_intervenido}
- TAG: {codigo_tag}
- OT: {codigo_documento}
- Tipo de Mantenimiento: {tipo_mantenimiento}
- Fecha Intervención: {fecha_intervencion}
- Fecha Emisión: {fecha_emision}
- Responsables Ejecución: {responsables_ejecucion}
- Responsable Informe: {responsable_informe}
- Condición Reportada: {condicion_reportada}
- Condición Encontrada: {condicion_encontrada}
- Diagnóstico: {diagnostico_realizado}
- Alcance: {alcance_trabajos}
- Impacto Operativo: {impacto_operativo}
- Impactos Varios: {impactos_varios}
- Causa Raíz: {causa_raiz_info}
- Recursos y Costos: {recursos_costos}
- Estado del Equipo: {estado_equipo}
- Nivel de Riesgo: {nivel_riesgo}
- Restricciones: {restricciones_op}
- Recomendaciones Inmediatas: {recomendaciones_inmediatas}
- Riesgos Remanentes: {riesgos_remanentes}
- Conclusión Gerencial: {conclusion_gerencial}
- Acciones del Plan: {acciones_str}
- Evidencias: {evidencias_str}

ESTRUCTURA OBLIGATORIA (Utiliza exactamente estos 12 títulos en formato Markdown H1 `#`):
# 1. DATOS GENERALES
# 2. RESUMEN EJECUTIVO
# 3. DESCRIPCIÓN DE LA FALLA O INTERVENCIÓN
# 4. IMPACTO OPERACIONAL
# 5. ANÁLISIS DE CAUSA RAÍZ
# 6. ACCIONES EJECUTADAS
# 7. RECURSOS UTILIZADOS
# 8. EVIDENCIAS FOTOGRÁFICAS
# 9. ESTADO ACTUAL DEL EQUIPO
# 10. RIESGOS PENDIENTES Y RECOMENDACIONES
# 11. PLAN DE ACCIÓN Y SEGUIMIENTO
# 12. CONCLUSIÓN EJECUTIVA
"""
                try:
                    respuesta = cliente_ai.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt_gerencial
                    )
                    st.session_state["texto_informe"] = respuesta.text
                    st.session_state["meta_informe"] = {
                        "empresa": empresa_nombre,
                        "area": area_proceso,
                        "tag": codigo_tag,
                        "equipo": equipo_intervenido,
                        "ot": codigo_documento,
                        "tipo": tipo_mantenimiento,
                        "emision": str(fecha_emision),
                        "responsable": responsable_informe,
                        "imagenes": lista_imagenes,
                        "acciones": st.session_state.get("acciones_plan", [])
                    }
                    st.success("¡Informe gerencial generado con éxito y optimizado para exportación!")
                except Exception as e:
                    st.error(f"Error al conectar con la API de Gemini: {e}")

# ==============================================================================
# 7. VISUALIZACIÓN Y EXPORTACIÓN PROFESIONAL (WORD Y PDF)
# ==============================================================================
if "texto_informe" in st.session_state:
    st.markdown("---")
    st.markdown("## 📄 Vista Previa del Informe Ejecutivo Gerencial")
    
    contenido_md = st.session_state["texto_informe"]
    metadatos = st.session_state["meta_informe"]
    
    with st.expander("🔍 Ver reporte completo estructurado", expanded=True):
        st.markdown(contenido_md)
        if metadatos["imagenes"]:
            st.markdown("### 📷 Registro Fotográfico Adjunto")
            for item in metadatos["imagenes"]:
                st.image(item["objeto"], caption=f"Foto {item['indice']} ({item['estado']}): {item['descripcion']}", use_container_width=True)

    st.markdown("---")
    st.markdown("### 📥 Exportación de Documentos Oficiales Gerenciales")
    col_w, col_p = st.columns(2)
    
    # --------------------------------------------------------------------------
    # 7.1 EXPORTACIÓN A MICROSOFT WORD (.DOCX) - SIN ASTERISCOS SUELTOS
    # --------------------------------------------------------------------------
    with col_w:
        if st.button("📥 Descargar Reporte Word (.docx)", use_container_width=True):
            doc = Document()
            
            for sec in doc.sections:
                sec.top_margin = Inches(1)
                sec.bottom_margin = Inches(1)
                sec.left_margin = Inches(1)
                sec.right_margin = Inches(1)

            # Portada Word
            p_sup = doc.add_paragraph()
            p_sup.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_sup = p_sup.add_run(f"{metadatos['empresa']} | {metadatos['area']}\nREF: {metadatos['ot']}")
            r_sup.font.size = Pt(9)
            r_sup.font.color.rgb = RGBColor(100, 116, 139)
            
            doc.add_paragraph()
            
            p_tit = doc.add_paragraph()
            p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_tit = p_tit.add_run(f"INFORME EJECUTIVO DE MANTENIMIENTO\n{metadatos['tipo'].upper()}")
            r_tit.font.size = Pt(16)
            r_tit.font.bold = True
            r_tit.font.color.rgb = RGBColor(15, 23, 42)
            
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sub = p_sub.add_run(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}")
            r_sub.font.size = Pt(11)
            r_sub.font.bold = True
            r_sub.font.color.rgb = RGBColor(67, 56, 202)
            
            doc.add_paragraph()
            
            # Tabla de metadatos en portada Word
            t_meta = doc.add_table(rows=5, cols=2)
            t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
            datos_t = [
                ("Fecha de Emisión:", metadatos["emision"]),
                ("Área / Proceso:", metadatos["area"]),
                ("Tipo de Mantenimiento:", metadatos["tipo"]),
                ("Orden de Trabajo (OT):", metadatos["ot"]),
                ("Responsable del Informe:", metadatos["responsable"])
            ]
            for idx_t, (lbl, val) in enumerate(datos_t):
                c_cells = t_meta.rows[idx_t].cells
                c_cells[0].text = lbl
                c_cells[0].paragraphs[0].runs[0].font.bold = True
                c_cells[1].text = val
            
            doc.add_page_break()
            
            # Procesamiento de líneas para Word con limpieza de Markdown
            for line in contenido_md.split('\n'):
                line_text = line.strip()
                if line_text.startswith("# "):
                    h = doc.add_heading(line_text.replace("# ", "").replace("## ", ""), level=1)
                    h.runs[0].font.color.rgb = RGBColor(15, 23, 42)
                elif line_text.startswith("## "):
                    h = doc.add_heading(line_text.replace("## ", ""), level=2)
                    h.runs[0].font.color.rgb = RGBColor(67, 56, 202)
                elif line_text.startswith("* ") or line_text.startswith("- "):
                    p_bullet = doc.add_paragraph(style='List Bullet')
                    r_b = p_bullet.add_run(line_text[2:].replace('**', ''))
                    r_b.font.size = Pt(10)
                    r_b.font.name = 'Arial'
                    p_bullet.paragraph_format.line_spacing = 1.15
                    p_bullet.paragraph_format.space_after = Pt(3)
                elif line_text:
                    limpiar_markdown_para_docx(doc, line_text)

            # Inserción de tabla del Plan de Acción en Word si existe
            if metadatos["acciones"]:
                doc.add_heading("Plan de Acción y Seguimiento (Detalle)", level=2)
                t_act = doc.add_table(rows=len(metadatos["acciones"]) + 1, cols=4)
                t_act.alignment = WD_TABLE_ALIGNMENT.CENTER
                headers = ["Acción", "Responsable", "Fecha Compromiso", "Estado"]
                for j, h_text in enumerate(headers):
                    cell = t_act.rows[0].cells[j]
                    cell.text = h_text
                    cell.paragraphs[0].runs[0].font.bold = True
                for i, act in enumerate(metadatos["acciones"]):
                    row_cells = t_act.rows[i + 1].cells
                    row_cells[0].text = act["accion"]
                    row_cells[1].text = act["responsable"]
                    row_cells[2].text = act["fecha"]
                    row_cells[3].text = act["estado"]

            # Evidencias Fotográficas en Word
            if metadatos["imagenes"]:
                doc.add_heading("Anexo: Evidencias Fotográficas", level=2)
                for item in metadatos["imagenes"]:
                    img_io = io.BytesIO()
                    item["objeto"].save(img_io, format='PNG')
                    img_io.seek(0)
                    
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(img_io, width=Inches(4.5))
                    
                    p_cap = doc.add_paragraph()
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    r_c1 = p_cap.add_run(f"Foto {item['indice']} [{item['estado']}]. ")
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
                file_name=f"Informe_Gerencial_{metadatos['tag']}_{metadatos['emision']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # --------------------------------------------------------------------------
    # 7.2 EXPORTACIÓN A PDF (REPORTLAB) - LIMPIEZA DE ASTERISCOS Y HTML NATIVO
    # --------------------------------------------------------------------------
    with col_p:
        if st.button("📥 Descargar Reporte PDF (.pdf)", use_container_width=True):
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
            
            style_title = ParagraphStyle(
                'PdfTitle', parent=styles['Heading1'],
                fontName='Helvetica-Bold', fontSize=18, leading=22,
                textColor=colors.HexColor('#0F172A'), alignment=TA_CENTER, spaceAfter=12
            )
            style_subtitle = ParagraphStyle(
                'PdfSubTitle', parent=styles['Normal'],
                fontName='Helvetica-Bold', fontSize=11, leading=15,
                textColor=colors.HexColor('#4338CA'), alignment=TA_CENTER, spaceAfter=20
            )
            style_h1 = ParagraphStyle(
                'PdfH1', parent=styles['Heading1'],
                fontName='Helvetica-Bold', fontSize=13, leading=17,
                textColor=colors.HexColor('#0F172A'), spaceBefore=14, spaceAfter=6
            )
            style_h2 = ParagraphStyle(
                'PdfH2', parent=styles['Heading2'],
                fontName='Helvetica-Bold', fontSize=11, leading=15,
                textColor=colors.HexColor('#4338CA'), spaceBefore=10, spaceAfter=4
            )
            style_body = ParagraphStyle(
                'PdfBody', parent=styles['Normal'],
                fontName='Helvetica', fontSize=9.5, leading=13.5,
                textColor=colors.HexColor('#334155'), alignment=TA_JUSTIFY, spaceAfter=6
            )
            style_bullet = ParagraphStyle(
                'PdfBullet', parent=styles['Normal'],
                fontName='Helvetica', fontSize=9.5, leading=13.5,
                textColor=colors.HexColor('#334155'), alignment=TA_JUSTIFY, leftIndent=15, spaceAfter=4
            )
            style_caption = ParagraphStyle(
                'PdfCaption', parent=styles['Normal'],
                fontName='Helvetica-Oblique', fontSize=8.5, leading=11,
                textColor=colors.HexColor('#475569'), alignment=TA_LEFT, spaceAfter=12
            )

            story = []
            
            # Portada PDF
            story.append(Spacer(1, 30))
            story.append(Paragraph(f"<b>{metadatos['empresa'].upper()} — {metadatos['area'].upper()}</b>", style_subtitle))
            story.append(Spacer(1, 10))
            story.append(Paragraph(f"INFORME EJECUTIVO DE MANTENIMIENTO<br/>{metadatos['tipo'].upper()}", style_title))
            story.append(Paragraph(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}", style_subtitle))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#4338CA'), spaceAfter=15))
            
            t_data = [
                [Paragraph("<b>Fecha de Emisión:</b>", style_body), Paragraph(metadatos["emision"], style_body)],
                [Paragraph("<b>Área / Proceso:</b>", style_body), Paragraph(metadatos["area"], style_body)],
                [Paragraph("<b>Tipo de Mantenimiento:</b>", style_body), Paragraph(metadatos["tipo"], style_body)],
                [Paragraph("<b>Orden de Trabajo (OT):</b>", style_body), Paragraph(metadatos["ot"], style_body)],
                [Paragraph("<b>Responsable del Informe:</b>", style_body), Paragraph(metadatos["responsable"], style_body)]
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

            # Procesamiento de líneas para PDF con conversión limpia de markdown a HTML de ReportLab
            for line in contenido_md.split('\n'):
                line_clean = line.strip()
                if line_clean.startswith("# "):
                    story.append(Paragraph(f"<b>{line_clean.replace('# ', '')}</b>", style_h1))
                    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#CBD5E1'), spaceAfter=6))
                elif line_clean.startswith("## "):
                    story.append(Paragraph(f"<b>{line_clean.replace('## ', '')}</b>", style_h2))
                elif line_clean.startswith("* ") or line_clean.startswith("- "):
                    texto_formateado = limpiar_markdown_para_pdf(line_clean[2:])
                    story.append(Paragraph(f"• {texto_formateado}", style_bullet))
                elif line_clean:
                    texto_formateado = limpiar_markdown_para_pdf(line_clean)
                    story.append(Paragraph(texto_formateado, style_body))

            # Tabla del Plan de Acción en PDF si existe
            if metadatos["acciones"]:
                story.append(Paragraph("<b>Detalle del Plan de Acción</b>", style_h1))
                table_act_data = [[Paragraph("<b>Acción</b>", style_body), Paragraph("<b>Responsable</b>", style_body), Paragraph("<b>Fecha</b>", style_body), Paragraph("<b>Estado</b>", style_body)]]
                for act in metadatos["acciones"]:
                    table_act_data.append([
                        Paragraph(act["accion"], style_body),
                        Paragraph(act["responsable"], style_body),
                        Paragraph(act["fecha"], style_body),
                        Paragraph(act["estado"], style_body)
                    ])
                t_acciones = Table(table_act_data, colWidths=[180, 120, 100, 104])
                t_acciones.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EEF2FF')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('PADDING', (0,0), (-1,-1), 5),
                ]))
                story.append(t_acciones)
                story.append(Spacer(1, 10))

            # Anexo Fotográfico en PDF
            if metadatos["imagenes"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas</b>", style_h1))
                story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#4338CA'), spaceAfter=10))
                
                for item in metadatos["imagenes"]:
                    img_bytes = io.BytesIO()
                    rgb_img = item["objeto"].convert('RGB')
                    rgb_img.save(img_bytes, format='JPEG', quality=90)
                    img_bytes.seek(0)
                    
                    img_flow = RLImage(img_bytes, width=380, height=260)
                    img_flow.hAlign = 'CENTER'
                    story.append(img_flow)
                    story.append(Spacer(1, 6))
                    
                    caption_str = f"<b>Foto {item['indice']} [{item['estado']}].</b> <i>{item['descripcion']}</i><br/><b>Fuente:</b> Elaboración propia."
                    story.append(Paragraph(caption_str, style_caption))
                    story.append(Spacer(1, 10))

            doc_pdf.build(
                story, 
                canvasmaker=lambda *args, **kwargs: type(
                    'CanvasGerencialPersonalizado', 
                    (CanvasGerencial,), 
                    {'empresa_meta': metadatos["empresa"], 'codigo_doc': metadatos["ot"]}
                )(*args, **kwargs)
            )
            buffer_pdf.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Reporte PDF (.pdf)",
                data=buffer_pdf,
                file_name=f"Informe_Gerencial_{metadatos['tag']}_{metadatos['emision']}.pdf",
                mime="application/pdf"
            )
