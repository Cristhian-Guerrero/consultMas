"""
Configuración central — URLs, selectores, timeouts, mensajes.
"""

import os

# ─────── URLs y selectores Express ───────
DIAN_URL_BASICA = os.getenv(
    "DIAN_URL",
    "https://muisca.dian.gov.co/WebGestionmasiva/DefSelPublicacionesExterna.faces"
)
SEL_NIT_ID_BASICA  = "vistaSelPublicacionesExterna:formSelPublicacionesExterna:numNit"
SEL_DV_ID_BASICA   = "vistaSelPublicacionesExterna:formSelPublicacionesExterna:dv"
BTN_BUSCAR_ID_BASICA = "vistaSelPublicacionesExterna:formSelPublicacionesExterna:btnBuscar"
FIELDS_BASICA = {
    "razonSocial":    "vistaSelPublicacionesExterna:formSelPublicacionesExterna:razonSocial",
    "primerApellido": "vistaSelPublicacionesExterna:formSelPublicacionesExterna:primerApellido",
    "segundoApellido":"vistaSelPublicacionesExterna:formSelPublicacionesExterna:segundoApellido",
    "primerNombre":   "vistaSelPublicacionesExterna:formSelPublicacionesExterna:primerNombre",
    "otrosNombres":   "vistaSelPublicacionesExterna:formSelPublicacionesExterna:otrosNombres",
}

# ─────── URLs y selectores RUT Detallado ───────
DIAN_URL_RUT    = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces"
SEL_NIT_ID_RUT  = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit"
BTN_BUSCAR_ID_RUT = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar"
FIELDS_RUT = {
    "dv":             "vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv",
    "primerApellido": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido",
    "segundoApellido":"vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido",
    "primerNombre":   "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre",
    "otrosNombres":   "vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres",
    "estado":         "vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado",
}
ERROR_CSS = ".ui-messages-error-detail"

# ─────── Timeouts ───────
class TimeoutConfig:
    INITIAL_WAIT      = 0.7
    POST_NIT_WAIT     = 0.5
    RESULTS_WAIT      = 2.0
    CF_BYPASS_WAIT    = 0.5
    INITIAL_WAIT_RUT  = 1
    POST_NIT_WAIT_RUT = 2
    RESULTS_WAIT_RUT  = 3
    CAPTCHA_WAIT      = 0
    MAX_RETRIES       = 2
    BACKOFF_FACTOR    = 2
    BROWSER_REST_TIME = 0.2

# ─────── Pool de navegadores ───────
MAX_POOL_SIZE = 4
BROWSER_CONFIGS = [
    {"user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",          "resolution": "1920x1080", "name": "Chrome_Win_FHD"},
    {"user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0",                                          "resolution": "1366x768",  "name": "Firefox_Win_HD"},
    {"user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",     "resolution": "1440x900",  "name": "Chrome_Mac_Retina"},
    {"user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0", "resolution": "1920x1200", "name": "Edge_Win_WUXGA"},
]

# ─────── Mensajes y tips ───────
MENSAJES_PROFESIONALES = [
    "📋 Verificando cumplimiento normativo tributario...",
    "🏛️ Consultando bases de datos oficiales DIAN...",
    "📊 Actualizando registros para auditoría fiscal...",
    "⚖️ Validando información para revisión fiscal...",
    "📈 Procesando datos para análisis contable...",
    "🔍 Verificando estados tributarios actualizados...",
    "📝 Recopilando información para declaraciones...",
    "💼 Consolidando datos para reportes gerenciales...",
    "🎯 Optimizando procesos de consulta masiva...",
    "🏢 Fortaleciendo control interno empresarial...",
    "📋 Actualizando expedientes de terceros...",
    "⚡ Agilizando procesos de debida diligencia...",
    "🔒 Garantizando integridad de la información...",
    "📊 Preparando insumos para conciliaciones...",
    "💡 Mejorando eficiencia en procesos contables...",
    "🔍 Ejecutando consulta detallada de estado RUT...",
    "📋 Verificando registro activo en base DIAN...",
]

TIPS_CONTABLES = [
    "💡 Tip: Mantenga actualizada la información de terceros para evitar inconsistencias en declaraciones",
    "⚖️ Normativa: Los contadores deben verificar la vigencia del RUT de sus clientes mensualmente",
    "📋 Recordatorio: La revisión fiscal requiere evidencia documental de todas las transacciones",
    "🎯 Buena práctica: Concilie regularmente la información tributaria con las bases de datos oficiales",
    "📊 Consejo: Documente todos los procesos de verificación para futuras auditorías",
    "🏛️ Actualización: DIAN requiere información veraz y oportuna según Decreto 2041/2023",
    "💼 Estrategia: Implemente controles internos robustos para la gestión de terceros",
    "⚡ Eficiencia: Use herramientas automatizadas para optimizar tiempo en consultas masivas",
    "🔍 Control: Verifique periódicamente cambios en estado tributario de proveedores",
    "📈 Análisis: Correlacione información tributaria con movimientos contables",
    "🎓 Formación: Manténgase actualizado en normatividad tributaria vigente",
    "🔒 Seguridad: Proteja la información tributaria bajo principios de confidencialidad",
    "📝 Documentación: Registre todas las consultas para trazabilidad de procesos",
    "💡 Innovación: Adopte tecnologías que mejoren la precisión de sus análisis",
    "⭐ Excelencia: La calidad en la información es clave para decisiones acertadas",
    "🎯 Estado RUT: La consulta detallada proporciona información de registro más precisa",
]

# Índices rotativos para mensajes
_mensaje_index = 0
_tip_index = 0

def obtener_mensaje_profesional():
    global _mensaje_index
    msg = MENSAJES_PROFESIONALES[_mensaje_index % len(MENSAJES_PROFESIONALES)]
    _mensaje_index += 1
    return msg

def obtener_tip_contable():
    global _tip_index
    tip = TIPS_CONTABLES[_tip_index % len(TIPS_CONTABLES)]
    _tip_index += 1
    return tip
