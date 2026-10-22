"""
================================================================================
SISTEMA INSTITUCIONAL DE INFORMES TÉCNICOS DE MANTENIMIENTO (SITM)
Versión: 6.0.0 - Edición Gerencial Avanzada con Gestión de Logos y Editor Previo
Desarrollado para: Informes Gerenciales de Alta Confiabilidad Industrial
Diseñado por: Roger Huamani
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

# Estilos CSS Corporativos y Modernos con Contraste Adecuado
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
        color: #1E3A8A;
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }
    .watermark-footer {
        font-size: 0.75rem;
        color: #64748B;
        text-align: center;
        margin-top: 2rem;
        border-top: 1px solid #E2E8F0;
        padding-top: 10px;
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
# 3. UTILIDADES DE PARSEO Y LIMPIEZA DE MARKDOWN
# ==============================================================================
def limpiar_markdown_para_docx(doc_instance, texto_linea):
    """Procesa marcas Markdown (**negrita**) y las añade a Word con tipografía negra legible."""
    p = doc_instance.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(4)
    
    partes = re.split(r'(\*\*.*?\*\*)', texto_linea)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**'):
            run = p.add_run(parte[2:-2])
            run.font.bold = True
            run.font.name = 'Arial'
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(15, 23, 42) # Azul oscuro / Negro formal
        else:
            if parte:
                run = p.add_run(parte)
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(30, 41, 59) # Texto negro con contraste adecuado
    return p

def limpiar_markdown_para_pdf(texto_linea):
    """Convierte Markdown a etiquetas HTML nativas de ReportLab."""
    texto_convertido = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', texto_linea)
    texto_convertido = texto_convertido.replace('*', '')
    return texto_convertido

# ==============================================================================
# 4. CANVAS PERSONALIZADO CON LOGOS Y NUMERACIÓN PDF
# ==============================================================================
class CanvasGerencialLogos(canvas.Canvas):
    """Canvas de doble pasada para paginación dinámica y soporte de dos logos institucionales."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self.empresa_meta = "COMPAÑÍA INDUSTRIAL"
        self.codigo_doc = "SITM-GER-01"
        self.logo1_img = None
        self.logo2_img = None

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
        self.setFillColor(colors.HexColor("#1E3A8A"))
        
        # Cabecera superior institucional en páginas subsiguientes (o portada)
        if self._pageNumber > 1:
            self.drawString(54, 755, f"{self.empresa_meta.upper()} — INFORME EJECUTIVO DE MANTENIMIENTO")
            self.drawRightString(558, 755, f"REF: {self.codigo_doc}")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.75)
            self.line(54, 745, 558, 745)
            
            # Pie de página inferior
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#475569"))
            page_text = f"Página {self._pageNumber} de {page_count}"
            self.drawRightString(558, 32, page_text)
            self.drawString(54, 32, "REPORTE GERENCIAL OFICIAL — TOMA DE DECISIONES")
            self.line(54, 44, 558, 44)
            
        self.restoreState()

# ==============================================================================
# 5. BARRA LATERAL - CONFIGURACIÓN GERENCIAL Y LOGOS PERSISTENTES
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/engineering.png", width=60)
    st.markdown("### **Panel Gerencial SITM**")
    st.markdown("*Configuración del Reporte Ejecutivo*")
    st.markdown("---")
    
    if "GEMINI_API_KEY" not in st.secrets:
        st.markdown("#### 🔑 Credencial API Gemini")
        st.text_input("API Key:", type="password", key="gemini_api_key_input")
        st.markdown("---")

    st.markdown("#### 🖼️ Logos Corporativos (Persistentes)")
    st.info("Cargue los dos logos institucionales una sola vez. Se mantendrán para todos sus informes.")
    
    logo_file_1 = st.file_uploader("Logo Principal (Empresa)", type=["png", "jpg", "jpeg"], key="uploader_logo_1")
    if logo_file_1 is not None:
        st.session_state["persistent_logo_1"] = logo_file_1
        
    logo_file_2 = st.file_uploader("Logo Secundario (Grupo / Filial)", type=["png", "jpg", "jpeg"], key="uploader_logo_2")
    if logo_file_2 is not None:
        st.session_state["persistent_logo_2"] = logo_file_2
        
    if "persistent_logo_1" in st.session_state:
        st.success("✔ Logo 1 cargado en sesión.")
    if "persistent_logo_2" in st.session_state:
        st.success("✔ Logo 2 cargado en sesión.")

    st.markdown("---")
    st.markdown("#### 🏢 Datos Organizacionales")
    empresa_nombre = st.text_input("Compañía / Planta", "ORICA")
    asunto_informe = st.text_input("Asunto / Título del Informe", "Mantenimiento Correctivo y Diagnóstico de Falla en Horno LEWCO N°1")
    area_proceso = st.text_input("Área o Proceso", "Detonadores")
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
    responsables_ejecucion = st.text_area("Responsables de Ejecución", "Hector Preciado, Lizandro Paucar, Julio Nunjar")
    responsable_informe = st.text_input("Responsable del Informe", "Ing. Roger Huamani Torres")

    # Marca de agua en lateral solicitada
    st.markdown('<div class="watermark-footer">Software diseñado por Roger Huamani</div>', unsafe_allow_html=True)

# ==============================================================================
# 6. INTERFAZ PRINCIPAL DE ENTRADA
# ==============================================================================
st.markdown('<div class="main-header">Generador Ejecutivo de Informes de Mantenimiento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Plataforma orientada a Gerencia: Ingrese los datos, adjunte imágenes específicas por sección, valide y edite el informe antes de exportar.</div>', unsafe_allow_html=True)

tab_datos, tab_plan, tab_evidencias, tab_preview = st.tabs([
    "📝 1. Datos y Contexto Técnico", 
    "📊 2. Plan de Acción", 
    "📷 3. Evidencias e Imágenes por Sección", 
    "🚀 4. Validación, Edición y Exportación"
])

# ------------------------------------------------------------------------------
# PESTAÑA 1: DATOS Y CONTEXTO TÉCNICO
# ------------------------------------------------------------------------------
with tab_datos:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Identificación del Activo")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        equipo_intervenido = st.text_input("Equipo Intervenido", "Horno LEWCO N°1")
    with col_e2:
        codigo_tag = st.text_input("Código o TAG del Equipo", "H4 (asociado a Horno LEWCO 3, información preliminar)")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Detalle Técnico para el Reporte Gerencial")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        condicion_reportada = st.text_area("Condición Reportada / Síntoma Inicial", "Operador reporta condición de vibración elevada en el sistema de extracción.", height=110)
        condicion_encontrada = st.text_area("Condición Encontrada (Real)", "Desgaste avanzado en rodamientos del ventilador de extracción y desalineación en acoplamiento.", height=110)
        diagnostico_realizado = st.text_area("Diagnóstico Realizado", "Análisis de vibraciones confirma desbalance dinámico y fatiga en soportes mecánicos.", height=110)
        alcance_trabajos = st.text_area("Alcance de Trabajos Ejecutados", "Desmontaje de conjunto rotativo, reemplazo de rodamientos SKF, alineación láser y balanceo dinámico in-situ.", height=110)
        
    with col_c2:
        impacto_operativo = st.text_area("Impacto Operacional (Horas indisponibilidad, producción)", "6 horas de indisponibilidad. Demora menor en lote de proceso de secado sin afectar entrega final.", height=110)
        impactos_varios = st.text_area("Impactos (Seguridad, Ambiental, Económico Estimado)", "Seguridad: Cero incidentes. Económico: $1,200 USD entre repuestos y horas hombre de taller.", height=110)
        causa_raiz_info = st.text_area("Análisis de Causa Raíz (Inmediata, Básica, Raíz y Metodología)", "Inmediata: Alta vibración. Básica: Lubricación deficiente acumulada. Raíz: Intervalo de engrase desactualizado. Metodología: 5 Porqués.", height=110)
        recursos_costos = st.text_area("Recursos y Costos (Personal, Materiales, Costos USD)", "Personal: 3 técnicos (Hector Preciado, Lizandro Paucar, Julio Nunjar). Costo total estimado: $1,200 USD.", height=110)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🔹 Estado, Riesgos y Conclusiones Gerenciales")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        estado_equipo = st.selectbox("Estado Actual del Equipo", ["Operativo", "Operativo con restricciones", "Fuera de servicio", "En observación o monitoreo"])
        nivel_riesgo = st.selectbox("Nivel de Riesgo Actual", ["Bajo", "Medio", "Alto", "Crítico"])
        restricciones_op = st.text_input("Restricciones Operacionales", "Ninguna. Operación normal al 100% de capacidad.")
        recomendaciones_inmediatas = st.text_area("Recomendaciones Inmediatas", "Verificar temperatura de chumaceras a las 24 horas de arranque.", height=90)
    with col_s2:
        riesgos_remanentes = st.text_area("Riesgos Pendientes y Consecuencias Potenciales", "Ninguno de alta criticidad tras el reemplazo y balanceo exitoso.", height=90)
        conclusion_gerencial = st.text_area("Conclusión Ejecutiva y Beneficio Esperado", "Intervención completada satisfactoriamente, restaurando los parámetros nominales de vibración y confiabilidad del horno.", height=90)
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 2: PLAN DE ACCIÓN Y SEGUIMIENTO
# ------------------------------------------------------------------------------
with tab_plan:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📊 Plan de Acción y Seguimiento Gerencial")
    
    if "acciones_plan" not in st.session_state:
        st.session_state["acciones_plan"] = [
            {"accion": "Actualizar frecuencia de lubricación en pauta preventiva del Horno LEWCO.", "responsable": "Hector Preciado", "fecha": "2026-11-01", "estado": "Pendiente"}
        ]
        
    col_a1, col_a2, col_a3, col_a4 = st.columns([3, 2, 2, 2])
    with col_a1:
        nueva_accion = st.text_input("Acción Propuesta", placeholder="Ej: Actualizar pauta preventiva...")
    with col_a2:
        nuevo_resp = st.text_input("Responsable", placeholder="Ej: Nombre técnico")
    with col_a3:
        nueva_fecha = st.date_input("Fecha Compromiso", date.today(), key="date_comp_plan")
    with col_a4:
        nuevo_estado = st.selectbox("Estado", ["Pendiente", "En proceso", "Completado"], key="select_estado_plan")
        
    if st.button("➕ Agregar Acción al Plan"):
        if nueva_accion:
            st.session_state["acciones_plan"].append({
                "accion": nueva_accion,
                "responsable": nuevo_resp,
                "fecha": str(nueva_fecha),
                "estado": nuevo_estado
            })
            st.success("Acción agregada correctamente.")
            
    if st.session_state["acciones_plan"]:
        st.table(st.session_state["acciones_plan"])
        if st.button("🗑️ Limpiar Plan de Acción"):
            st.session_state["acciones_plan"] = []
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 3: EVIDENCIAS E IMÁGENES ESPECÍFICAS POR SECCIÓN
# ------------------------------------------------------------------------------
with tab_evidencias:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📷 Carga de Imágenes con Comentarios por Sección")
    st.info("Puede adjuntar la cantidad de imágenes que necesite, especificando a qué sección corresponden (Descripción de Falla o Acciones Ejecutidas).")
    
    seccion_destino = st.selectbox(
        "Seleccione la sección donde desea incrustar la imagen:",
        ["3. DESCRIPCIÓN DE LA FALLA O INTERVENCIÓN", "6. ACCIONES EJECUTADAS", "8. EVIDENCIAS FOTOGRÁFICAS (ANEXO GENERAL)"]
    )
    
    img_subida_seccion = st.file_uploader("Cargar fotografía para esta sección", type=["jpg", "jpeg", "png"], key="uploader_sec_img")
    comentario_img_sec = st.text_input("Comentario / Descripción técnica de la imagen:")
    estado_img_sec = st.selectbox("Estado de la intervención:", ["Antes", "Durante", "Después"])
    
    if "imagenes_por_seccion" not in st.session_state:
        st.session_state["imagenes_por_seccion"] = []
        
    if st.button("📥 Adjuntar Imagen a la Sección"):
        if img_subida_seccion is not None and comentario_img_sec:
            st.session_state["imagenes_por_seccion"].append({
                "seccion": seccion_destino,
                "objeto": Image.open(img_subida_seccion),
                "nombre": img_subida_seccion.name,
                "descripcion": comentario_img_sec,
                "estado": estado_img_sec
            })
            st.success("¡Imagen agregada exitosamente a la sección correspondiente!")
        else:
            st.error("Por favor, cargue una imagen y escriba su respectivo comentario.")
            
    if st.session_state["imagenes_por_seccion"]:
        st.markdown("#### Imágenes Registradas Actualmente:")
        for idx, item in enumerate(st.session_state["imagenes_por_seccion"]):
            col_info1, col_info2 = st.columns([2, 1])
            with col_info1:
                st.markdown(f"**[{item['seccion']}]** - {item['descripcion']} *(Estado: {item['estado']})*")
            with col_info2:
                if st.button(f"🗑️ Eliminar #{idx+1}", key=f"del_img_{idx}"):
                    st.session_state["imagenes_por_seccion"].pop(idx)
                    st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# PESTAÑA 4: VALIDACIÓN, EDICIÓN Y EXPORTACIÓN
# ------------------------------------------------------------------------------
with tab_preview:
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 🚀 Generación, Validación y Edición del Informe")
    st.info("Haga clic en generar para redactar el informe con IA. Una vez generado, **podrá editar el contenido directamente en el cuadro de texto** antes de descargarlo en Word o PDF.")
    
    btn_generar = st.button("✨ Generar Borrador del Informe con IA", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if btn_generar:
        cliente_ai = inicializar_cliente_gemini()
        if not cliente_ai:
            st.error("❌ Error de Autenticación: Configure su clave API de Gemini en Streamlit Secrets o en la barra lateral.")
        else:
            with st.spinner("🔄 Procesando datos y redactando informe ejecutivo gerencial..."):
                
                imgs_falla = [i for i in st.session_state.get("imagenes_por_seccion", []) if "FALLA" in i["seccion"]]
                imgs_acciones = [i for i in st.session_state.get("imagenes_por_seccion", []) if "ACCIONES" in i["seccion"]]
                imgs_anexo = [i for i in st.session_state.get("imagenes_por_seccion", []) if "EVIDENCIAS" in i["seccion"]]
                
                falla_img_str = "\n".join([f"- [Imagen Falla - {i['estado']}]: {i['descripcion']}" for i in imgs_falla])
                acc_img_str = "\n".join([f"- [Imagen Acción - {i['estado']}]: {i['descripcion']}" for i in imgs_acciones])
                
                acciones_str = "\n".join([f"- [ ] {a['accion']} | Responsable: {a['responsable']} | Fecha: {a['fecha']} | Estado: {a['estado']}" for a in st.session_state.get("acciones_plan", [])])
                
                prompt_gerencial = f"""
Actúa como un Especialista Senior en Gestión de Mantenimiento Industrial, Confiabilidad y Reportes Ejecutivos para Gerencia.
Genera un INFORME EJECUTIVO DE MANTENIMIENTO profesional, estructurado rigurosamente en los 12 apartados solicitados.

REGLAS ESTRICTAS DE REDACCIÓN:
- NO incluyas frases automáticas como "Información no disponible" o explicaciones sobre datos faltantes. Si un campo no tiene detalle extenso, redáctalo de forma concisa y directa basada estrictamente en los datos de entrada.
- Lenguaje ejecutivo, claro, preciso, objetivo y basado en hechos orientados a gerencia.
- Redacta principalmente en formato de PÁRRAFOS formales y bien desarrollados. Utiliza viñetas (con asteriscos `*`) UNICAMENTE cuando sea estrictamente necesario para aportar claridad.

DATOS DE ENTRADA:
- Compañía: {empresa_nombre}
- Asunto: {asunto_informe}
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
- Imágenes en Falla: {falla_img_str}
- Impacto Operativo: {impacto_operativo}
- Impactos Varios: {impactos_varios}
- Causa Raíz: {causa_raiz_info}
- Recursos y Costos: {recursos_costos}
- Imágenes en Acciones: {acc_img_str}
- Estado del Equipo: {estado_equipo}
- Nivel de Riesgo: {nivel_riesgo}
- Restricciones: {restricciones_op}
- Recomendaciones Inmediatas: {recomendaciones_inmediatas}
- Riesgos Remanentes: {riesgos_remanentes}
- Conclusión Gerencial: {conclusion_gerencial}
- Acciones del Plan: {acciones_str}

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
                        "asunto": asunto_informe,
                        "area": area_proceso,
                        "tag": codigo_tag,
                        "equipo": equipo_intervenido,
                        "ot": codigo_documento,
                        "tipo": tipo_mantenimiento,
                        "intervencion": str(fecha_intervencion),
                        "emision": str(fecha_emision),
                        "ejecucion": responsables_ejecucion,
                        "responsable": responsable_informe,
                        "imagenes": st.session_state.get("imagenes_por_seccion", []),
                        "acciones": st.session_state.get("acciones_plan", [])
                    }
                    st.success("¡Borrador generado con éxito! Puede validarlo y editarlo a continuación antes de exportar.")
                except Exception as e:
                    st.error(f"Error al conectar con la API de Gemini: {e}")

# ==============================================================================
# 7. EDITOR DE VALIDACIÓN Y EXPORTACIÓN FINAL (WORD Y PDF)
# ==============================================================================
if "texto_informe" in st.session_state:
    st.markdown("---")
    st.markdown("## ✍️ Validación y Edición del Contenido del Informe")
    st.info("💡 Puede modificar cualquier sección, corregir redacción o ajustar párrafos directamente en este cuadro de texto antes de generar los documentos oficiales Word o PDF.")
    
    # Campo de edición interactivo (Observación 2)
    texto_editado = st.text_area(
        "Editor de Contenido del Informe",
        value=st.session_state["texto_informe"],
        height=450,
        key="editor_texto_informe_activo"
    )
    
    # Actualizar texto en sesión si el usuario edita
    st.session_state["texto_informe"] = texto_editado

    st.markdown("---")
    st.markdown("### 📥 Exportación de Documentos Oficiales Gerenciales")
    col_w, col_p = st.columns(2)
    
    metadatos = st.session_state["meta_informe"]
    
    # --------------------------------------------------------------------------
    # 7.1 EXPORTACIÓN A WORD (.DOCX) - CON TABLA DE DATOS GENERALES
    # --------------------------------------------------------------------------
    with col_w:
        if st.button("📥 Descargar Reporte Word (.docx)", use_container_width=True):
            doc = Document()
            
            for sec in doc.sections:
                sec.top_margin = Inches(1)
                sec.bottom_margin = Inches(1)
                sec.left_margin = Inches(1)
                sec.right_margin = Inches(1)

            # Portada / Cabecera Word
            p_sup = doc.add_paragraph()
            p_sup.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_sup = p_sup.add_run(f"{metadatos['empresa']} | {metadatos['area']}\nREF: {metadatos['ot']}")
            r_sup.font.size = Pt(9)
            r_sup.font.color.rgb = RGBColor(100, 116, 139)
            
            doc.add_paragraph()
            
            p_tit = doc.add_paragraph()
            p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_tit = p_tit.add_run(f"INFORME EJECUTIVO DE MANTENIMIENTO\n{metadatos['asunto'].upper()}")
            r_tit.font.size = Pt(15)
            r_tit.font.bold = True
            r_tit.font.color.rgb = RGBColor(30, 58, 138) # Azul oscuro
            
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sub = p_sub.add_run(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}")
            r_sub.font.size = Pt(11)
            r_sub.font.bold = True
            r_sub.font.color.rgb = RGBColor(30, 41, 59)
            
            doc.add_paragraph()
            
            # Tabla de Datos Generales en Word (Observación 1)
            t_meta = doc.add_table(rows=9, cols=2)
            t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
            datos_t = [
                ("Compañía:", metadatos["empresa"]),
                ("Asunto:", metadatos["asunto"]),
                ("Área o Proceso:", metadatos["area"]),
                ("Equipo Intervenido:", metadatos["equipo"]),
                ("TAG (referencia):", metadatos["tag"]),
                ("OT Principal:", metadatos["ot"]),
                ("Tipo de Mantenimiento:", metadatos["tipo"]),
                ("Fecha Intervención:", metadatos["intervencion"]),
                ("Responsables Ejecución:", metadatos["ejecucion"])
            ]
            for idx_t, (lbl, val) in enumerate(datos_t):
                row_cells = t_meta.rows[idx_t].cells
                row_cells[0].text = lbl
                row_cells[0].paragraphs[0].runs[0].font.bold = True
                row_cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(30, 58, 138)
                row_cells[1].text = val
                row_cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(30, 41, 59)
            
            doc.add_page_break()
            
            # Procesamiento de líneas para Word con limpieza
            for line in st.session_state["texto_informe"].split('\n'):
                line_text = line.strip()
                if line_text.startswith("# "):
                    h = doc.add_heading(line_text.replace("# ", "").replace("## ", ""), level=1)
                    h.runs[0].font.color.rgb = RGBColor(30, 58, 138) # Azul oscuro en títulos
                elif line_text.startswith("## "):
                    h = doc.add_heading(line_text.replace("## ", ""), level=2)
                    h.runs[0].font.color.rgb = RGBColor(30, 58, 138)
                elif line_text.startswith("* ") or line_text.startswith("- "):
                    p_bullet = doc.add_paragraph(style='List Bullet')
                    r_b = p_bullet.add_run(line_text[2:].replace('**', ''))
                    r_b.font.size = Pt(10)
                    r_b.font.name = 'Arial'
                    r_b.font.color.rgb = RGBColor(30, 41, 59)
                    p_bullet.paragraph_format.line_spacing = 1.15
                    p_bullet.paragraph_format.space_after = Pt(3)
                elif line_text:
                    # Insertar imágenes específicas si corresponden a secciones de Falla o Acciones
                    limpiar_markdown_para_docx(doc, line_text)

            # Inserción de tabla del Plan de Acción en Word si existe
            if metadatos["acciones"]:
                doc.add_heading("Plan de Acción y Seguimiento", level=2)
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

            # Evidencias Fotográficas generales o por sección en Word
            if metadatos["imagenes"]:
                doc.add_heading("Anexo: Evidencias Fotográficas y Registros", level=2)
                for item in metadatos["imagenes"]:
                    img_io = io.BytesIO()
                    item["objeto"].save(img_io, format='PNG')
                    img_io.seek(0)
                    
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(img_io, width=Inches(4.5))
                    
                    p_cap = doc.add_paragraph()
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    r_c1 = p_cap.add_run(f"[{item['seccion']}] - Foto ({item['estado']}): ")
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
                label="⬇️ Descargar Reporte Word (.docx)",
                data=buffer_docx,
                file_name=f"Informe_Gerencial_{metadatos['tag']}_{metadatos['emision']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # --------------------------------------------------------------------------
    # 7.2 EXPORTACIÓN A PDF (REPORTLAB) - CON TABLA DE DATOS GENERALES Y LOGOS
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
            
            # Estilos con títulos en azul oscuro y textos en negro/gris de alto contraste (Observación 5)
            style_title = ParagraphStyle(
                'PdfTitle', parent=styles['Heading1'],
                fontName='Helvetica-Bold', fontSize=16, leading=20,
                textColor=colors.HexColor('#1E3A8A'), alignment=TA_CENTER, spaceAfter=10
            )
            style_subtitle = ParagraphStyle(
                'PdfSubTitle', parent=styles['Normal'],
                fontName='Helvetica-Bold', fontSize=10, leading=14,
                textColor=colors.HexColor('#334155'), alignment=TA_CENTER, spaceAfter=15
            )
            style_h1 = ParagraphStyle(
                'PdfH1', parent=styles['Heading1'],
                fontName='Helvetica-Bold', fontSize=12, leading=16,
                textColor=colors.HexColor('#1E3A8A'), spaceBefore=12, spaceAfter=4
            )
            style_h2 = ParagraphStyle(
                'PdfH2', parent=styles['Heading2'],
                fontName='Helvetica-Bold', fontSize=10, leading=14,
                textColor=colors.HexColor('#1E3A8A'), spaceBefore=8, spaceAfter=3
            )
            style_body = ParagraphStyle(
                'PdfBody', parent=styles['Normal'],
                fontName='Helvetica', fontSize=9.5, leading=13.5,
                textColor=colors.HexColor('#1F2937'), alignment=TA_JUSTIFY, spaceAfter=6
            )
            style_bullet = ParagraphStyle(
                'PdfBullet', parent=styles['Normal'],
                fontName='Helvetica', fontSize=9.5, leading=13.5,
                textColor=colors.HexColor('#1F2937'), alignment=TA_JUSTIFY, leftIndent=15, spaceAfter=4
            )
            style_caption = ParagraphStyle(
                'PdfCaption', parent=styles['Normal'],
                fontName='Helvetica-Oblique', fontSize=8.5, leading=11,
                textColor=colors.HexColor('#475569'), alignment=TA_LEFT, spaceAfter=10
            )

            story = []
            
            # Renderizar Logos si están cargados en sesión (Observación 6)
            logo_flowables = []
            if "persistent_logo_1" in st.session_state:
                img_io_1 = io.BytesIO(st.session_state["persistent_logo_1"].getvalue())
                r_img1 = RLImage(img_io_1, width=110, height=45)
                r_img1.hAlign = 'LEFT'
                logo_flowables.append(r_img1)
            if "persistent_logo_2" in st.session_state:
                img_io_2 = io.BytesIO(st.session_state["persistent_logo_2"].getvalue())
                r_img2 = RLImage(img_io_2, width=90, height=40)
                r_img2.hAlign = 'RIGHT'
                logo_flowables.append(r_img2)
                
            if logo_flowables:
                # Tabla de cabecera con logos
                t_logos_data = [[logo_flowables[0] if len(logo_flowables)>0 else "", logo_flowables[1] if len(logo_flowables)>1 else ""]]
                t_logo_table = Table(t_logos_data, colWidths=[250, 254])
                t_logo_table.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('ALIGN', (0,0), (0,0), 'LEFT'),
                    ('ALIGN', (1,0), (1,0), 'RIGHT'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 10)
                ]))
                story.append(t_logo_table)

            story.append(Spacer(1, 10))
            story.append(Paragraph(f"<b>INFORME EJECUTIVO DE MANTENIMIENTO</b><br/>{metadatos['asunto'].upper()}", style_title))
            story.append(Paragraph(f"ACTIVO: {metadatos['tag']} - {metadatos['equipo']}", style_subtitle))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1E3A8A'), spaceAfter=15))
            
            # Tabla de Datos Generales Formal en PDF (Observación 1)
            t_data = [
                [Paragraph("<b>Campo</b>", style_body), Paragraph("<b>Detalle</b>", style_body)],
                [Paragraph("Compañía:", style_body), Paragraph(metadatos["empresa"], style_body)],
                [Paragraph("Asunto:", style_body), Paragraph(metadatos["asunto"], style_body)],
                [Paragraph("Área o Proceso:", style_body), Paragraph(metadatos["area"], style_body)],
                [Paragraph("Equipo Intervenido:", style_body), Paragraph(metadatos["equipo"], style_body)],
                [Paragraph("TAG (referencia):", style_body), Paragraph(metadatos["tag"], style_body)],
                [Paragraph("OT Principal:", style_body), Paragraph(metadatos["ot"], style_body)],
                [Paragraph("Tipo de Mantenimiento:", style_body), Paragraph(metadatos["tipo"], style_body)],
                [Paragraph("Fecha Intervención:", style_body), Paragraph(metadatos["intervencion"], style_body)],
                [Paragraph("Responsables Ejecución:", style_body), Paragraph(metadatos["ejecucion"], style_body)]
            ]
            t_pdf = Table(t_data, colWidths=[150, 354])
            t_pdf.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F1F5F9')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 5),
            ]))
            story.append(t_pdf)
            story.append(PageBreak())

            # Procesamiento de líneas editadas para PDF
            for line in st.session_state["texto_informe"].split('\n'):
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
                story.append(Paragraph("<b>Plan de Acción y Seguimiento</b>", style_h1))
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

            # Anexo Fotográfico y por secciones en PDF
            if metadatos["imagenes"]:
                story.append(PageBreak())
                story.append(Paragraph("<b>Evidencias Fotográficas y Registros</b>", style_h1))
                story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1E3A8A'), spaceAfter=10))
                
                for item in metadatos["imagenes"]:
                    img_bytes = io.BytesIO()
                    item["objeto"].convert('RGB').save(img_bytes, format='JPEG', quality=90)
                    img_bytes.seek(0)
                    
                    img_flow = RLImage(img_bytes, width=380, height=260)
                    img_flow.hAlign = 'CENTER'
                    story.append(img_flow)
                    story.append(Spacer(1, 6))
                    
                    caption_str = f"<b>[{item['seccion']}] - Foto ({item['estado']}).</b> <i>{item['descripcion']}</i><br/><b>Fuente:</b> Elaboración propia."
                    story.append(Paragraph(caption_str, style_caption))
                    story.append(Spacer(1, 10))

            doc_pdf.build(
                story, 
                canvasmaker=lambda *args, **kwargs: type(
                    'CanvasGerencialLogosCustom', 
                    (CanvasGerencialLogos,), 
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
