import time
import threading
from datetime import timedelta, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import webbrowser
import os
import sys
from DrissionPage import ChromiumPage, ChromiumOptions
from PIL import Image, ImageTk

# Importación condicional para compatibilidad Linux/Windows
import platform
SISTEMA_OPERATIVO = platform.system()

# pygetwindow solo funciona en Windows
if SISTEMA_OPERATIVO == "Windows":
    try:
        import pygetwindow as gw
        PYGETWINDOW_DISPONIBLE = True
    except ImportError:
        PYGETWINDOW_DISPONIBLE = False
        print("⚠️ pygetwindow no disponible en Windows")
else:
    # En Linux/Mac no usamos pygetwindow
    PYGETWINDOW_DISPONIBLE = False
    gw = None  # Para evitar errores de referencia
    print(f"ℹ️ Sistema operativo: {SISTEMA_OPERATIVO} - pygetwindow no necesario")


# ─────── Configuración DUAL ────────
# ═══ CONSULTA Express (ORIGINAL) ═══
DIAN_URL_BASICA = os.getenv(
    "DIAN_URL",
    "https://muisca.dian.gov.co/WebGestionmasiva/DefSelPublicacionesExterna.faces"
)

SEL_NIT_ID_BASICA = "vistaSelPublicacionesExterna:formSelPublicacionesExterna:numNit"
SEL_DV_ID_BASICA = "vistaSelPublicacionesExterna:formSelPublicacionesExterna:dv"
BTN_BUSCAR_ID_BASICA = "vistaSelPublicacionesExterna:formSelPublicacionesExterna:btnBuscar"

FIELDS_BASICA = {
    "razonSocial": "vistaSelPublicacionesExterna:formSelPublicacionesExterna:razonSocial",
    "primerApellido": "vistaSelPublicacionesExterna:formSelPublicacionesExterna:primerApellido",
    "segundoApellido": "vistaSelPublicacionesExterna:formSelPublicacionesExterna:segundoApellido", 
    "primerNombre": "vistaSelPublicacionesExterna:formSelPublicacionesExterna:primerNombre",
    "otrosNombres": "vistaSelPublicacionesExterna:formSelPublicacionesExterna:otrosNombres"
}

# ═══ CONSULTA DETALLADA RUT (NUEVA) ═══
DIAN_URL_RUT = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces"

SEL_NIT_ID_RUT = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit"
BTN_BUSCAR_ID_RUT = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar"

FIELDS_RUT = {
    "dv": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv",
    "primerApellido": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido",
    "segundoApellido": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido",
    "primerNombre": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre",
    "otrosNombres": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres",
    "estado": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado"
}

ERROR_CSS = ".ui-messages-error-detail"

# ─────────── Helpers JS (DUAL COMPATIBLE) ───────────
def set_field_js(driver, element_id, value):
    try:
        js = f"""
        const el = document.getElementById('{element_id}');
        if (!el) return false;
        el.value = '{value}';
        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        return el.value;
        """
        result = driver.run_js(js)
        return result == value
    except Exception as e:
        print(f"Error en set_field_js para {element_id}: {str(e)}")
        return False

def click_js(driver, element_id):
    try:
        js = f"""
        const el = document.getElementById('{element_id}');
        if (!el) return false;
        el.click();
        return true;
        """
        return driver.run_js(js)
    except Exception as e:
        print(f"Error en click_js para {element_id}: {str(e)}")
        return False

def calcular_dv(nit: str) -> str:
    try:
        coef = [71, 67, 59, 53, 47, 43, 41, 37, 29, 23, 19, 17, 13, 7, 3]
        n15 = nit.zfill(15)
        total = sum(int(n15[i]) * coef[i] for i in range(15))
        residuo = total % 11
        dv_val = 11 - residuo
        if dv_val == 11:
            return "0"
        if dv_val == 10:
            return "1"
        return str(dv_val)
    except Exception as e:
        print(f"Error calculando DV para NIT {nit}: {str(e)}")
        return "0"

# ═══ CONFIGURACIÓN DUAL - OPTIMIZADA SOLO PARA Express ═══
class TimeoutConfig:
    # Timeouts OPTIMIZADOS para consulta Express
    INITIAL_WAIT = 0.7        # OPTIMIZADO: 2 → 1
    POST_NIT_WAIT = 0.5       # OPTIMIZADO: 3 → 0.8 
    RESULTS_WAIT = 2.0        # OPTIMIZADO: 4 → 2.5
    CF_BYPASS_WAIT = 0.5      # OPTIMIZADO: 1.5 → 0.8
    
    # Timeouts ORIGINALES para RUT (sin cambios)
    INITIAL_WAIT_RUT = 1      # MANTENER ORIGINAL
    POST_NIT_WAIT_RUT = 2     # MANTENER ORIGINAL
    RESULTS_WAIT_RUT = 3      # MANTENER ORIGINAL
    
    # Configuración general
    CAPTCHA_WAIT = 0
    MAX_RETRIES = 2
    BACKOFF_FACTOR = 2
    BROWSER_REST_TIME = 0.2   # OPTIMIZADO: 0.5 → 0.2

# Pool de navegadores (MANTIENE SISTEMA EXISTENTE)
MAX_POOL_SIZE = 4
BROWSER_POOL = []
CURRENT_BROWSER_INDEX = 0

BROWSER_CONFIGS = [
    {
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "resolution": "1920x1080",
        "name": "Chrome_Win_FHD"
    },
    {
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0",
        "resolution": "1366x768", 
        "name": "Firefox_Win_HD"
    },
    {
        "user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "resolution": "1440x900",
        "name": "Chrome_Mac_Retina"
    },
    {
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
        "resolution": "1920x1200",
        "name": "Edge_Win_WUXGA"
    }
]

# ═══ MENSAJES PROFESIONALES ACTUALIZADOS ═══
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
    "📋 Verificando registro activo en base DIAN..."
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
    "🎯 Estado RUT: La consulta detallada proporciona información de registro más precisa"
]

mensaje_index = 0
tip_index = 0

def obtener_mensaje_profesional():
    global mensaje_index
    mensaje = MENSAJES_PROFESIONALES[mensaje_index % len(MENSAJES_PROFESIONALES)]
    mensaje_index += 1
    return mensaje

def obtener_tip_contable():
    global tip_index
    tip = TIPS_CONTABLES[tip_index % len(TIPS_CONTABLES)]
    tip_index += 1
    return tip

# ═══ FUNCIONES DE NAVEGADOR CON POSICIONAMIENTO INVISIBLE ═══
def crear_navegador_con_config(config_index=0):
    """Navegador configurado para aparecer FUERA DE LA VISTA"""
    config = BROWSER_CONFIGS[config_index % len(BROWSER_CONFIGS)]
    options = ChromiumOptions().auto_port()
    
    # CONFIGURACIÓN PARA NAVEGADORES INVISIBLES AL USUARIO
    options.set_argument("--disable-dev-shm-usage")
    options.set_argument("--window-size=400,300")
    options.set_argument("--window-position=-2000,-2000")
    options.set_argument("--disable-extensions")
    options.set_argument("--disable-plugins")
    options.set_argument("--no-first-run")
    options.set_argument("--disable-default-apps")
    # NO --headless para evitar detección de CAPTCHA
    
    return ChromiumPage(addr_or_opts=options), config['name']

def get_browser_from_pool():
    """CORREGIDO: Pool simple como en código exitoso"""
    limpiar_navegadores_inactivos()
    if BROWSER_POOL:
        return BROWSER_POOL.pop()
    return crear_navegador_con_config(0)[0]

def return_browser_to_pool(driver):
    """CORREGIDO: Lógica simple como en código exitoso"""
    limpiar_navegadores_inactivos()
    if driver and len(BROWSER_POOL) < MAX_POOL_SIZE:
        try:
            driver.get('about:blank')
            BROWSER_POOL.append(driver)
        except:
            try:
                driver.quit()
            except:
                pass

def limpiar_navegadores_inactivos():
    """CORREGIDO: Limpieza simple como en código exitoso"""
    global BROWSER_POOL
    navegadores_activos = []
    
    for driver in BROWSER_POOL:
        try:
            driver.title
            navegadores_activos.append(driver)
        except:
            try:
                driver.quit()
            except:
                pass
    
    BROWSER_POOL = navegadores_activos
    return len(BROWSER_POOL)

def cerrar_todos_los_navegadores():
    global BROWSER_POOL
    for driver in BROWSER_POOL:
        try:
            driver.quit()
        except Exception as e:
            print(f"Error cerrando navegador: {e}")
    BROWSER_POOL = []

def minimizar_ventanas_chromium():
    """
    Minimiza ventanas de Chromium (solo Windows).
    En Linux/Mac las ventanas ya están ocultas por --window-position=-2000,-2000
    """
    if not PYGETWINDOW_DISPONIBLE:
        # En Linux/Mac no es necesario, ventanas ya están fuera de vista
        return
    
    # Solo en Windows
    try:
        for window in gw.getWindowsWithTitle("Chromium"):
            window.minimize()
        print("✅ Ventanas Chromium minimizadas (Windows)")
    except Exception as e:
        print(f"⚠️ No se pudieron minimizar ventanas: {e}")



# ═══ FUNCIONES DE SCRAPING BÁSICO (OPTIMIZADAS) ═══
def check_for_results_basica(driver):
    try:
        js = f"""
        const nombre = document.getElementById('{FIELDS_BASICA['primerNombre']}');
        const razon = document.getElementById('{FIELDS_BASICA['razonSocial']}');
        if ((nombre && nombre.textContent.trim()) || (razon && razon.textContent.trim())) {{
            return 'success';
        }}
        const err = document.querySelector('{ERROR_CSS}');
        if (err && err.textContent.trim()) {{
            return 'error:' + err.textContent.trim();
        }}
        return 'waiting';
        """
        return driver.run_js(js)
    except Exception as e:
        print(f"Error en check_for_results_basica: {e}")
        return 'waiting'

def extract_data_basica(driver):
    data = {}
    try:
        for key, eid in FIELDS_BASICA.items():
            try:
                js = f"""
                const el = document.getElementById('{eid}');
                return el ? el.textContent.trim() : null;
                """
                result = driver.run_js(js)
                data[key] = result if result else None
            except Exception as e:
                print(f"Error obteniendo {key}: {str(e)}")
                data[key] = None
    except Exception as e:
        print(f"Error general en extract_data_basica: {str(e)}")
    return data

def reset_form_basica(driver):
    try:
        js = f"""
        document.getElementById('{SEL_NIT_ID_BASICA}').value = '';
        document.getElementById('{SEL_DV_ID_BASICA}').value = '';
        return true;
        """
        return driver.run_js(js)
    except Exception as e:
        print(f"Error al resetear formulario básico: {str(e)}")
        return False

# ═══ FUNCIONES RUT (SIN CAMBIOS - MANTENER EXACTAMENTE IGUAL) ═══
def set_nit_rut_js(driver, nit):
    """MANTENER EXACTAMENTE IGUAL"""
    try:
        js_script = f"""
        const input = document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit');
        input.value = '{nit}';
        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
        return input.value;
        """
        result = driver.run_js(js_script)
        return result == nit
    except Exception as e:
        print(f"Error en set_nit_rut_js: {str(e)}")
        return False

def click_button_rut_js(driver):
    """MANTENER EXACTAMENTE IGUAL"""
    try:
        js_script = """
        document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar').click();
        return true;
        """
        return driver.run_js(js_script)
    except Exception as e:
        print(f"Error en click_button_rut_js: {str(e)}")
        return False

def check_for_results_rut(driver):
    """MANTENER EXACTAMENTE IGUAL"""
    try:
        js_script = """
        const dv = document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv');
        const error = document.querySelector('.ui-messages-error-detail');
        if (dv && dv.textContent.trim()) return 'success';
        if (error) return 'error:' + error.textContent.trim();
        return 'waiting';
        """
        return driver.run_js(js_script)
    except Exception as e:
        print(f"Error en check_for_results_rut: {e}")
        return "waiting"

def extract_data_rut(driver):
    """MANTENER EXACTAMENTE IGUAL"""
    data = {}
    fields = {
        "dv": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv",
        "razonSocial": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:razonSocial",
        "primerNombre": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre",
        "otrosNombres": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres",
        "primerApellido": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido",
        "segundoApellido": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido",
        "estado": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado",
    }
    
    for key, selector in fields.items():
        try:
            js_script = f"""
            const element = document.getElementById('{selector}');
            return element ? element.textContent.trim() : null;
            """
            js_result = driver.run_js(js_script)
            data[key] = js_result if js_result else None
        except Exception as e:
            print(f"Error obteniendo {key}: {str(e)}")
            data[key] = None
    
    return data

def reset_form_rut(driver):
    """MANTENER EXACTAMENTE IGUAL"""
    try:
        js_script = """
        document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit').value = '';
        return true;
        """
        return driver.run_js(js_script)
    except Exception as e:
        print(f"Error al resetear formulario RUT: {str(e)}")
        return False

def check_and_close_error_rut(driver):
    """MANTENER EXACTAMENTE IGUAL"""
    try:
        js_script = """
        const errorTable = document.querySelector("table[background*='fondoMensajeError.gif']");
        if (errorTable) {
            const closeButton = errorTable.querySelector("img[src*='botcerrarrerror.gif']");
            if (closeButton) {
                closeButton.click();
                return true;
            }
        }
        return false;
        """
        return driver.run_js(js_script)
    except Exception as e:
        print(f"Error en check_and_close_error_rut: {str(e)}")
        return False

def check_captcha_error_rut(driver):
    """MANTENER EXACTAMENTE IGUAL"""
    try:
        js_script = """
        const captchaError = document.getElementById('g-recaptcha-error');
        if (captchaError && captchaError.innerText.includes('Se requiere validar captcha.')) {
            return true;
        }
        return false;
        """
        return driver.run_js(js_script)
    except Exception as e:
        print(f"Error en check_captcha_error_rut: {e}")
        return False

def check_no_inconsistencias_and_close(driver):
    try:
        js = """
        const dialogElements = document.querySelectorAll('.ui-dialog-content p, .ui-messages-info-detail, .ui-growl-message p');
        for (let elem of dialogElements) {
            const text = elem.textContent.toLowerCase();
            if (text.includes('no se encontraron documentos') || 
                text.includes('no se encontraron') ||
                text.includes('sin inconsistencias')) {
                const closeBtn = document.querySelector('.ui-dialog-titlebar-close, .ui-growl-icon-close');
                if (closeBtn) {
                    closeBtn.click();
                    return true;
                }
                return true;
            }
        }
        return false;
        """
        return driver.run_js(js)
    except Exception as e:
        print(f"Error en check_no_inconsistencias_and_close: {e}")
        return False

# ═══ FUNCIÓN DE CONSULTA Express OPTIMIZADA ═══
def consultar_nit_basica(nit: str, attempt: int = 1):
    """CONSULTA EXPRESS OPTIMIZADA - SOLO TIMEOUTS MEJORADOS"""
    driver = None
    try:
        driver = get_browser_from_pool()
        
        try:
            current_url = driver.url
        except:
            try:
                driver.quit()
            except:
                pass
            driver = crear_navegador_con_config(0)[0]
        
        if not driver.url or DIAN_URL_BASICA not in driver.url:
            driver.get(DIAN_URL_BASICA)
            time.sleep(TimeoutConfig.INITIAL_WAIT)  # OPTIMIZADO: 1s en lugar de 2s
        
        try:
            from CloudflareBypasser import CloudflareBypasser
            cf = CloudflareBypasser(driver, max_retries=2, log=False)
            cf.bypass()
            time.sleep(TimeoutConfig.CF_BYPASS_WAIT)  # OPTIMIZADO: 0.8s en lugar de 1.5s
        except Exception as e:
            print(f"Warning: Error en bypass de Cloudflare: {e}")
        
        for reset_attempt in range(2):  # OPTIMIZADO: 2 en lugar de 3
            try:
                if reset_form_basica(driver):
                    break
                time.sleep(0.3)  # OPTIMIZADO: 0.3s en lugar de 0.5s
            except Exception as e:
                print(f"Intento {reset_attempt + 1} de reset falló: {e}")
                if reset_attempt == 1:
                    return {"status": "error", "data": {}, "error": "No se pudo resetear formulario"}
        
        time.sleep(0.2)  # OPTIMIZADO: 0.2s en lugar de 0.3s
        
        nit_establecido = False
        for retry in range(TimeoutConfig.MAX_RETRIES):
            try:
                if set_field_js(driver, SEL_NIT_ID_BASICA, nit):
                    nit_establecido = True
                    break
                time.sleep(0.3)  # OPTIMIZADO: 0.3s en lugar de 0.5s
            except Exception as e:
                print(f"Error estableciendo NIT en intento {retry + 1}: {e}")
        
        if not nit_establecido:
            return {"status": "error", "data": {}, "error": f"No se pudo establecer el NIT {nit}"}
        
        time.sleep(TimeoutConfig.POST_NIT_WAIT)  # OPTIMIZADO: 0.8s en lugar de 3s
        
        dv = calcular_dv(nit)
        dv_establecido = False
        for retry in range(TimeoutConfig.MAX_RETRIES):
            try:
                if set_field_js(driver, SEL_DV_ID_BASICA, dv):
                    dv_establecido = True
                    break
                time.sleep(0.3)  # OPTIMIZADO: 0.3s en lugar de 0.5s
            except Exception as e:
                print(f"Error estableciendo DV en intento {retry + 1}: {e}")
        
        if not dv_establecido:
            print(f"⚠️ No se pudo establecer DV={dv} para NIT {nit}")
        
        time.sleep(0.2)  # OPTIMIZADO: 0.2s en lugar de 0.3s
        
        buscar_clicked = False
        for retry in range(TimeoutConfig.MAX_RETRIES):
            try:
                if click_js(driver, BTN_BUSCAR_ID_BASICA):
                    buscar_clicked = True
                    break
                time.sleep(0.3)  # OPTIMIZADO: 0.3s en lugar de 0.5s
            except Exception as e:
                print(f"Error haciendo clic en buscar en intento {retry + 1}: {e}")
        
        if not buscar_clicked:
            return {"status": "error", "data": {}, "error": "No se pudo hacer clic en Buscar"}
        
        time.sleep(0.8)  # OPTIMIZADO: 0.8s en lugar de 1.5s
        
        waited = 0
        status = 'waiting'
        no_data_detected = False
        
        while status == 'waiting' and waited < TimeoutConfig.RESULTS_WAIT:  # OPTIMIZADO: 2.5s en lugar de 4s
            try:
                if check_no_inconsistencias_and_close(driver):
                    no_data_detected = True
                    status = 'success_no_data'
                    break
                
                status = check_for_results_basica(driver)
                
                if status == 'waiting':
                    time.sleep(0.5)  # OPTIMIZADO: Verificaciones más frecuentes
                    waited += 0.5
            except Exception as e:
                print(f"Error verificando resultados: {e}")
                time.sleep(0.5)
                waited += 0.5
        
        if status.startswith('error:'):
            error_msg = status.split('error:', 1)[1]
            return {"status": "error", "data": {}, "error": f"NIT {nit}: {error_msg}"}
        
        if status == 'waiting':
            if attempt == 1:
                return {"status": "retry", "data": {}, "error": f"Timeout en primera consulta para NIT {nit}"}
            else:
                return {"status": "error", "data": {}, "error": f"{nit}: Timeout después de reintentos"}
        
        try:
            data = extract_data_basica(driver)
        except Exception as e:
            print(f"Error extrayendo datos: {e}")
            return {"status": "error", "data": {}, "error": f"Error extrayendo datos: {str(e)}"}
            
        data['nit'] = nit
        data['dv'] = dv
        data['datetime'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data['tipo_consulta'] = 'basica'
        
        if attempt == 2:
            data['observacion'] = "Consulta exitosa en segundo intento"
        
        if no_data_detected or (not data.get("razonSocial") and not data.get("primerNombre")):
            data["razonSocial"] = "-"
            data["primerNombre"] = "-"
            data["primerApellido"] = "-" 
            data["segundoApellido"] = "-"
            data["otrosNombres"] = "-"
        
        return {"status": "success", "data": data, "error": None}
        
    except Exception as e:
        print(f"❌ Error consultando NIT básico {nit}: {str(e)}")
        return {"status": "error", "data": {}, "error": str(e)}
    finally:
        if driver:
            return_browser_to_pool(driver)

# ═══ FUNCIÓN RUT SIN CAMBIOS ═══
def consultar_nit_rut_detallado(nit: str, attempt: int = 1):
    """FUNCIÓN RUT SIN CAMBIOS - MANTENER EXACTAMENTE IGUAL"""
    driver = None
    try:
        driver = get_browser_from_pool()
        
        # Verificar conexión
        try:
            current_url = driver.url
        except:
            try:
                driver.quit()
            except:
                pass
            driver = crear_navegador_con_config(0)[0]
        
        if not driver.url or DIAN_URL_RUT not in driver.url:
            driver.get(DIAN_URL_RUT)
            time.sleep(TimeoutConfig.INITIAL_WAIT_RUT)  # USAR TIMEOUTS ORIGINALES

        # Cerrar errores como en código exitoso
        check_and_close_error_rut(driver)
        reset_form_rut(driver)

        # Configurar NIT con lógica exitosa
        retry_count = 0
        nit_set = False
        while retry_count < TimeoutConfig.MAX_RETRIES and not nit_set:
            nit_set = set_nit_rut_js(driver, nit)
            if not nit_set:
                retry_count += 1
                wait_time = TimeoutConfig.INITIAL_WAIT_RUT * (TimeoutConfig.BACKOFF_FACTOR**retry_count)
                time.sleep(wait_time)

        if not nit_set:
            return {"status": "error", "data": {}, "error": f"No se pudo establecer el NIT {nit} en RUT"}

        time.sleep(TimeoutConfig.POST_NIT_WAIT_RUT)  # USAR TIMEOUTS ORIGINALES

        # Hacer clic con lógica exitosa
        retry_count = 0
        button_clicked = False
        while retry_count < TimeoutConfig.MAX_RETRIES and not button_clicked:
            button_clicked = click_button_rut_js(driver)
            if not button_clicked:
                retry_count += 1
                wait_time = TimeoutConfig.INITIAL_WAIT_RUT * (TimeoutConfig.BACKOFF_FACTOR**retry_count)
                time.sleep(wait_time)

        if not button_clicked:
            return {"status": "error", "data": {}, "error": "No se pudo hacer clic en Buscar RUT"}

        time.sleep(1)

        # Manejo de CAPTCHA como en código exitoso
        MAX_CAPTCHA_RETRIES = 2
        if check_captcha_error_rut(driver):
            if attempt > MAX_CAPTCHA_RETRIES:
                print(f"Captcha detectado para NIT {nit}. Cancelando búsqueda.")
                return {"status": "error", "data": {}, "error": f"Captcha detectado en intento {attempt}. No se puede continuar."}
            else:
                print(f"Captcha detectado para NIT {nit}. Reintentando ({attempt}/{MAX_CAPTCHA_RETRIES})...")
                return {"status": "retry", "data": {}, "error": f"Captcha detectado en intento {attempt}. Reintentando..."}

        # Esperar resultados como en código exitoso
        max_wait_time = TimeoutConfig.RESULTS_WAIT_RUT * 0.8  # USAR TIMEOUTS ORIGINALES
        waited = 0
        result_status = "waiting"
        while result_status == "waiting" and waited < max_wait_time:
            result_status = check_for_results_rut(driver)
            if result_status == "waiting":
                time.sleep(1)
                waited += 1

        if result_status.startswith("error:"):
            error_msg = result_status.split("error:", 1)[1]
            return {"status": "error", "data": {}, "error": f"NIT {nit}: {error_msg}"}

        if result_status == "waiting":
            if attempt == 1:
                return {"status": "retry", "data": {}, "error": f"Estado 'waiting' en primera consulta para NIT {nit}"}
            else:
                return {"status": "error", "data": {}, "error": f"{nit}: No está inscrito en el RUT (Validado)"}

        # Extraer datos
        data = extract_data_rut(driver)
        data["nit"] = nit
        data["datetime"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data['tipo_consulta'] = 'rut_detallado'
        
        if attempt == 2:
            data["observacion"] = "Consulta RUT exitosa en segundo intento"

        if not data.get("razonSocial") and not data.get("primerNombre"):
            data["razonSocial"] = "No está inscrito en el RUT"
            data["primerNombre"] = "No está inscrito en el RUT"
            data["estado"] = "SIN INFORMACIÓN"

        return {"status": "success", "data": data, "error": None}
        
    except Exception as e:
        print(f"❌ Error consultando NIT RUT {nit}: {str(e)}")
        return {"status": "error", "data": {}, "error": str(e)}
    finally:
        if driver:
            return_browser_to_pool(driver)

# ═══ FUNCIÓN COORDINADORA DUAL ═══
def consultar_nit_individual(nit: str, tipo_consulta: str = "basica", attempt: int = 1):
    """FUNCIÓN COORDINADORA PARA AMBOS TIPOS DE CONSULTA"""
    if tipo_consulta == "rut_detallado":
        return consultar_nit_rut_detallado(nit, attempt)
    else:
        return consultar_nit_basica(nit, attempt)

# ═══════════════════════════════════════════════════════════════════════════════
# INTERFAZ DUAL PROFESIONAL - SIN CAMBIOS
# ═══════════════════════════════════════════════════════════════════════════════

class ConsultaRUTApp(tk.Tk):
    """Interfaz Dual Profesional - A.S. Contadores & Asesores SAS"""
    
    def __init__(self):
        super().__init__()
        
        # ═══ CONFIGURACIÓN INICIAL ═══
        self.title("Consulta Gestión Masiva DIAN | A.S. Contadores & Asesores SAS")
        self.ejecucion_activa = threading.Event()
        self.detener_proceso = False
        
        # ═══ NUEVA VARIABLE PARA TIPO DE CONSULTA ═══
        self.tipo_consulta = tk.StringVar(value="basica")
        
        # ═══ PALETA DE COLORES (CORREGIDA PARA FONDO BLANCO) ═══
        self.COLORS = {
            'primary': "#166534",         # Verde Bosque Profesional (Más oscuro y elegante)
            'primary_light': "#DCFCE7",   # Fondo menta muy suave (para selecciones)
            'primary_dark': '#14532D',    # Verde muy oscuro para textos importantes
            'secondary': '#64748B',       # Gris Slate (Botones secundarios)
            'accent': "#0F766E",          # Teal profesional
            'background': '#F8FAFC',      # Gris casi blanco (Fondo general)
            'surface': '#FFFFFF',         # Blanco absoluto (Tarjetas y Header)
            'text_primary': '#1E293B',    # Azul grisáceo muy oscuro (Mejor que negro)
            'text_secondary': '#64748B',  # Gris medio suave
            'text_light': '#FFFFFF',      # Texto blanco (solo para botones sólidos)
            'success': "#15803D",         # Verde éxito
            'warning': '#B45309',         # Naranja quemado
            'error': '#B91C1C',           # Rojo rubí
            'border': '#E2E8F0',          # Borde gris muy sutil
        }
        
        # ═══ FUENTES NATIVAS (EVITA EL PIXELADO DE TEXTO) ═══
        # Usamos Segoe UI (Windows) por defecto, es mucho más nítida que Arial
        self.FONTS = {
            'title': ('Segoe UI', 20, 'bold'),
            'subtitle': ('Segoe UI', 11, 'bold'),
            'body': ('Segoe UI', 10),
            'button': ('Segoe UI', 9, 'bold'),
            'small': ('Segoe UI', 9),
        }
        
        # ═══ CONFIGURACIÓN DE VENTANA ═══
        self.geometry("1000x650")
        self.minsize(950, 630)
        self.resizable(True, False)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.configure(bg=self.COLORS['background'])
        
        # ═══ CONFIGURAR ÍCONO CORPORATIVO ═══
        self.setup_improved_icon()
        
        # Centrar ventana
        self.center_window()
        
        # ═══ CONFIGURAR ESTILOS ═══
        self.setup_compact_styles()
        
        # ═══ CREAR INTERFAZ DUAL ═══
        self.create_dual_ui()
        
        # ═══ VARIABLES DE ESTADO ═══
        self.lista_nits = []
        self.generated_file = None
        self.rows_for_excel = []
        self.tiempo_inicio = None
        self.cronometro_activo = False
    
    def center_window(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_improved_icon(self):
        """Configuración BLINDADA del ícono (Windows/Linux)"""
        try:
            # 1. Definir rutas absolutas
            icon_path = os.path.abspath(self.resource_path("dian.ico"))
            logo_path = os.path.abspath(self.resource_path("logo.png"))
            
            icon_set = False

            # 2. Intentar método nativo de Windows (.ico)
            if os.path.exists(icon_path):
                try:
                    # 'default' aplica el ícono a esta ventana y futuras
                    self.iconbitmap(default=icon_path)
                    icon_set = True
                    print(f"✅ Ícono .ico cargado: {icon_path}")
                except Exception as e:
                    print(f"⚠️ No se pudo cargar .ico (intentando PNG): {e}")
            
            # 3. Si falló el .ico o no existe, intentar usar el PNG con iconphoto
            if not icon_set and os.path.exists(logo_path):
                try:
                    img = Image.open(logo_path)
                    photo = ImageTk.PhotoImage(img)
                    # True = aplica a todas las ventanas de la app
                    self.iconphoto(True, photo)
                    self.icon_refs = [photo] # Evitar garbage collection
                    print(f"✅ Logo PNG usado como ícono")
                except Exception as e:
                    print(f"❌ Falló fallback a PNG: {e}")

        except Exception as e:
            print(f"❌ Error crítico configurando ícono: {e}")
    
    def resource_path(self, relative_path):
        """Obtiene la ruta de recursos"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    def load_and_resize_logo(self, max_width=300, max_height=80):
        """Carga y redimensiona el logo"""
        try:
            logo_path = self.resource_path("logo.png")
            if not os.path.exists(logo_path):
                print(f"❌ Logo no encontrado: {logo_path}")
                return None
                
            try:
                pil_image = Image.open(logo_path)
                original_size = pil_image.size
                
                width_ratio = max_width / original_size[0]
                height_ratio = max_height / original_size[1]
                scale_ratio = min(width_ratio, height_ratio)
                
                if scale_ratio < 1:
                    new_size = (int(original_size[0] * scale_ratio), int(original_size[1] * scale_ratio))
                    pil_resized = pil_image.resize(new_size, Image.Resampling.LANCZOS)
                    logo_tk = ImageTk.PhotoImage(pil_resized)
                else:
                    logo_tk = ImageTk.PhotoImage(pil_image)
                
                return logo_tk
                
            except ImportError:
                logo = tk.PhotoImage(file=logo_path)
                return logo
                
        except Exception as e:
            print(f"❌ Error procesando logo: {e}")
            return None
    
    def setup_compact_styles(self):
        """Configura estilos modernos estilo Dashboard (Clean White)"""
        self.style = ttk.Style()
        self.style.theme_use('clam')  # 'clam' permite mejor personalización
        
        # ═══ CONFIGURACIÓN GENERAL ═══
        self.configure(bg=self.COLORS['background'])
        
        # ═══ ESTILOS DE CONTENEDORES (FRAMES) ═══
        # Header: Ahora BLANCO para que el logo resalte perfecto
        self.style.configure('Header.TFrame', 
                           background=self.COLORS['surface'],
                           relief='flat')
        
        # Panel Principal
        self.style.configure('Main.TFrame', 
                           background=self.COLORS['background'],
                           relief='flat')
        
        # Tarjetas (Card): Fondo blanco con borde sutil
        self.style.configure('Card.TLabelframe',
                           background=self.COLORS['surface'],
                           relief='solid',
                           borderwidth=1,
                           bordercolor=self.COLORS['border'])
                           
        self.style.configure('Card.TLabelframe.Label',
                           background=self.COLORS['surface'],
                           foreground=self.COLORS['primary'], # Título verde corporativo
                           font=self.FONTS['subtitle'])

        # ═══ TEXTOS (LABELS) ═══
        # Título Principal (Ahora Verde sobre Blanco)
        self.style.configure('HeaderTitle.TLabel',
                           background=self.COLORS['surface'],
                           foreground=self.COLORS['primary'],
                           font=self.FONTS['title'])
        
        # Subtítulos (Gris sobre Blanco)
        self.style.configure('HeaderSub.TLabel',
                           background=self.COLORS['surface'],
                           foreground=self.COLORS['text_secondary'],
                           font=self.FONTS['body'])
        
        # Textos normales
        self.style.configure('Body.TLabel',
                           background=self.COLORS['surface'],
                           foreground=self.COLORS['text_primary'],
                           font=self.FONTS['body'])
                           
        self.style.configure('Status.TLabel',
                           background=self.COLORS['surface'],
                           foreground=self.COLORS['text_secondary'],
                           font=self.FONTS['small'])

        # ═══ RADIO BUTTONS (ARREGLADO) ═══
        # Fondo blanco para que no se vea el recuadro "feo"
        self.style.configure('Consulta.TRadiobutton',
                           background=self.COLORS['surface'],
                           foreground=self.COLORS['text_primary'],
                           font=self.FONTS['body'],
                           focuscolor=self.COLORS['surface'], 
                           indicatorcolor=self.COLORS['surface'])
        
        self.style.map('Consulta.TRadiobutton',
                      indicatorcolor=[('selected', self.COLORS['primary'])],
                      foreground=[('selected', self.COLORS['primary'])])

        # ═══ BOTONES (MODERNOS) ═══
        self.style.configure('Primary.TButton',
                           background=self.COLORS['primary'],
                           foreground=self.COLORS['text_light'],
                           font=self.FONTS['button'],
                           borderwidth=0,
                           focuscolor='none')
        
        self.style.map('Primary.TButton',
                      background=[('active', self.COLORS['primary_dark']), 
                                ('pressed', self.COLORS['primary_dark'])])
        
        self.style.configure('Secondary.TButton',
                           background=self.COLORS['secondary'],
                           foreground=self.COLORS['text_light'],
                           font=self.FONTS['button'],
                           borderwidth=0,
                           focuscolor='none')
                           
        self.style.map('Secondary.TButton',
                      background=[('active', '#475569')])

        self.style.configure('Success.TButton',
                           background=self.COLORS['success'],
                           foreground=self.COLORS['text_light'],
                           font=self.FONTS['button'],
                           borderwidth=0)

        # ═══ BARRA DE PROGRESO ═══
        self.style.configure('Compact.Horizontal.TProgressbar',
                           background=self.COLORS['primary'],
                           troughcolor=self.COLORS['border'],
                           thickness=15,
                           borderwidth=0)
        # ═══ SCROLLBAR MODERNIZADO (Adiós barra gris vieja) ═══
        self.style.configure('Vertical.TScrollbar',
                           background='#CBD5E1',        # Color normal (Gris suave, no intrusivo)
                           troughcolor='#F8FAFC',       # Color del carril (Casi blanco, invisible)
                           bordercolor='#F8FAFC',       # Sin bordes feos
                           arrowcolor=self.COLORS['primary'], # Flechitas en Verde Corporativo
                           relief='flat')               # Totalmente plano
        
        # Efecto Hover: Se pone Verde cuando pasas el mouse
        self.style.map('Vertical.TScrollbar',
                      background=[('active', self.COLORS['primary']), 
                                ('pressed', self.COLORS['primary_dark'])],
                      arrowcolor=[('active', self.COLORS['primary_dark'])])
        
    def create_dual_ui(self):
        """Crea la interfaz dual con selector de tipo de consulta"""
        
        # ═══ HEADER COMPACTO ═══
        self.create_compact_header()
        
        # ═══ CONTENIDO PRINCIPAL CON SELECTOR DUAL ═══
        self.create_main_dual_grid()
    def create_compact_header(self):
        """Header limpio y moderno (CORREGIDO: Logo con fondo blanco)"""
        header_frame = ttk.Frame(self, style='Header.TFrame', padding="25 18")
        header_frame.pack(fill=tk.X)
        
        header_frame.columnconfigure(1, weight=1)
        header_frame.rowconfigure(0, weight=1)
        
        # ═══ 1. LOGO (Izquierda) ═══
        logo_container = ttk.Frame(header_frame, style='Header.TFrame')
        logo_container.grid(row=0, column=0, sticky="nsw", padx=(0, 30), pady=8)
        
        # Intentamos cargar logo
        logo = self.load_and_resize_logo(max_width=350, max_height=90)
        
        if logo:
            self.logo_image = logo
            # Frame contenedor blanco
            logo_frame = ttk.Frame(logo_container, style='Header.TFrame', padding=0)
            logo_frame.pack()
            
            logo_label = ttk.Label(logo_frame, 
                                 image=self.logo_image,
                                 style='Body.TLabel') 
            logo_label.pack()
        else:
            # Texto de respaldo
            logo_text_frame = ttk.Frame(logo_container, style='Header.TFrame', padding=10)
            logo_text_frame.pack()
            logo_text = ttk.Label(logo_text_frame,
                                text="🏢 A.S. CONTADORES &\nASESORES SAS",
                                style='HeaderTitle.TLabel',
                                justify=tk.CENTER)
            logo_text.pack()
        
        # ═══ 2. TÍTULOS (Derecha) ═══
        info_container = ttk.Frame(header_frame, style='Header.TFrame')
        info_container.grid(row=0, column=1, sticky="ew", pady=10)
        
        main_title = ttk.Label(info_container,
                              text="Gestión Masiva DIAN",
                              style='HeaderTitle.TLabel')
        main_title.pack(anchor="w", pady=(0, 4))
        
        subtitle = ttk.Label(info_container,
                           text="Sistema automatizado de consulta Express",
                           style='HeaderSub.TLabel')
        subtitle.pack(anchor="w", pady=(0, 3))
        
        services_label = ttk.Label(info_container,
                                 text="Asesoría Contable • Financiera • Tributaria • Revisoría Fiscal",
                                 font=self.FONTS['small'],
                                 background=self.COLORS['surface'],
                                 foreground=self.COLORS['accent'])
        services_label.pack(anchor="w")
    
    def create_main_dual_grid(self):
        """Crear contenido principal con selector dual"""
        main_container = ttk.Frame(self, style='Main.TFrame', padding="15")
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 15))
        
        main_container.columnconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=1)
        main_container.rowconfigure(2, weight=1)
        
        # ═══ FILA 1: SELECTOR DE TIPO DE CONSULTA ═══
        selector_frame = ttk.LabelFrame(main_container, 
                                      text="🎯 Modo de Consulta",
                                      style='Card.TLabelframe',
                                      padding="12")
        selector_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        # Contenedor para radio buttons
        radio_container = ttk.Frame(selector_frame, style='Main.TFrame')
        radio_container.pack(fill=tk.X)
        
        # Radio button para consulta Express
        self.radio_basica = ttk.Radiobutton(radio_container,
                                          text="📋 Express",
                                          variable=self.tipo_consulta,
                                          value="basica",
                                          style='Consulta.TRadiobutton',
                                          command=self.on_tipo_consulta_changed)
        self.radio_basica.pack(side=tk.LEFT, padx=(0, 30))
        
        # Radio button para consulta detallada
        """self.radio_detallada = ttk.Radiobutton(radio_container,
                                             text="🔍 RUT Detallada (Completa)",
                                             variable=self.tipo_consulta,
                                             value="rut_detallado",
                                             style='Consulta.TRadiobutton',
                                             command=self.on_tipo_consulta_changed)
        self.radio_detallada.pack(side=tk.LEFT)
        """
        
        # ═══ FILA 2: CONTROLES ═══
        controls_frame = ttk.LabelFrame(main_container, 
                                      text="🔧 Controles", 
                                      style='Card.TLabelframe',
                                      padding="12")
        controls_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        buttons_container = ttk.Frame(controls_frame, style='Main.TFrame')
        buttons_container.pack(fill=tk.X)
        
        self.btn_cargar_excel = ttk.Button(buttons_container,
                                         text="📁 Cargar Excel/CSV",
                                         command=self.on_cargar_excel,
                                         style='Primary.TButton')
        self.btn_cargar_excel.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_detener = ttk.Button(buttons_container,
                                    text="⏹️ Detener",
                                    command=self.detener_consulta,
                                    style='Secondary.TButton',
                                    state=tk.DISABLED)
        self.btn_detener.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_open_excel = ttk.Button(buttons_container,
                                       text="📊 Abrir Excel",
                                       command=self.open_excel,
                                       style='Success.TButton')
        
        # ═══ FILA 3: PROGRESO Y RESULTADOS ═══
        
        # ═══ COLUMNA IZQUIERDA: PROGRESO ═══
        progress_frame = ttk.LabelFrame(main_container,
                                      text="📊 Progreso",
                                      style='Card.TLabelframe',
                                      padding="12")
        progress_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 5))
        
        # Estado actual
        self.label_status = ttk.Label(progress_frame,
                                    text="Listo para comenzar",
                                    style='Body.TLabel')
        self.label_status.pack(anchor="w", pady=(0, 8))
        
        # Indicador de tipo de consulta activa
        self.label_tipo_activo = ttk.Label(progress_frame,
                                         text="📋 Consulta Express seleccionada",
                                         style='Status.TLabel')
        self.label_tipo_activo.pack(anchor="w", pady=(0, 8))
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(progress_frame,
                                          style='Compact.Horizontal.TProgressbar',
                                          mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))
        
        # Cronómetro
        self.label_cronometro = ttk.Label(progress_frame,
                                        text="⏱️ Tiempo: 00:00:00",
                                        style='Status.TLabel')
        self.label_cronometro.pack(anchor="w", pady=(0, 8))
        
        # ═══ ESTADÍSTICAS ═══
        stats_frame = ttk.Frame(progress_frame, style='Main.TFrame')
        stats_frame.pack(fill=tk.X, pady=(8, 0))
        
        self.stats_frame = ttk.Frame(stats_frame, style='Main.TFrame')
        self.stats_frame.pack(fill=tk.X)
        
        self.label_total = ttk.Label(self.stats_frame, text="Total: 0", style='Status.TLabel')
        self.label_total.pack(anchor="w")
        
        self.label_exitosos = ttk.Label(self.stats_frame, text="✅ Exitosos: 0", style='Status.TLabel')
        self.label_exitosos.pack(anchor="w")
        
        self.label_errores = ttk.Label(self.stats_frame, text="❌ Errores: 0", style='Status.TLabel')
        self.label_errores.pack(anchor="w")
        
        # ═══ COLUMNA DERECHA: RESULTADOS ═══
        # ═══ COLUMNA DERECHA: RESULTADOS (MEJORADA) ═══
        results_frame = ttk.LabelFrame(main_container,
                                     text="📋 Resultados del Proceso", 
                                     style='Card.TLabelframe',
                                     padding="1") # Padding fino y elegante
        results_frame.grid(row=2, column=1, sticky="nsew", padx=(5, 0))
        
        # Frame interno para dar margen
        text_container = ttk.Frame(results_frame, style='Main.TFrame', padding=1)
        text_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(text_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 🔴 AQUÍ CAMBIA: Fuente moderna y fondo gris pálido
        self.text_resultados = tk.Text(text_container,
                                     wrap=tk.WORD,
                                     height=9,
                                     yscrollcommand=scrollbar.set,
                                     bg='#F8FAFC',            # Gris casi blanco (muy elegante)
                                     fg=self.COLORS['text_primary'],
                                     font=('Segoe UI', 9),    # FUENTE MODERNA (Adiós letra de consola)
                                     padx=10,                 # Más espacio interno para leer mejor
                                     pady=10,
                                     relief='flat',           # Sin bordes 3D
                                     borderwidth=0)
        self.text_resultados.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.text_resultados.yview)
        
        # Mensaje inicial actualizado
        initial_message = (f"🏢 A.S. CONTADORES & ASESORES SAS\n"
                          f"Sistema de Consulta Automatizada DIAN DUAL\n\n"
                          f"🎯 OPCIONES DISPONIBLES:\n"
                          f"   📋 Consulta Express: Información general rápida\n"
                          
                          f"✅ NAVEGADORES INVISIBLES AL USUARIO\n"
                          f"💼 FIRMA ESPECIALIZADA EN:\n"
                          f"   • Asesoría Contable y Financiera\n"
                          f"   • Consultoría Tributaria Empresarial\n"
                          f"   • Revisoría Fiscal Profesional\n"
                          f"   • Optimización de Procesos Contables\n\n"
                          f"✨ 1. Seleccione tipo de consulta\n"
                          f"📁 2. Cargar archivo Excel/CSV (columna 'NIT')\n"
                          f"🚀 3. Sistema iniciará procesamiento automático\n\n"
                          f"Esperando configuración...\n")
        
        self.text_resultados.insert(tk.END, initial_message)
        self.text_resultados.config(state=tk.DISABLED)
        
        # ═══ FOOTER ═══
        footer_frame = ttk.Frame(self, style='Main.TFrame', padding="8 5")
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        footer_text = ttk.Label(footer_frame,
                              text="© 2026 A.S. Contadores & Asesores SAS | V4.6 CG",
                              style='Status.TLabel')
        footer_text.pack()
    
    def on_tipo_consulta_changed(self):
        """Callback cuando cambia el tipo de consulta"""
        tipo = self.tipo_consulta.get()
        if tipo == "basica":
            self.label_tipo_activo.config(text="📋 Consulta Express seleccionada (Rápida)")
            self.add_result_message("📋 Modo Básico: Consulta rápida de información general", "INFO")
        else:
            self.label_tipo_activo.config(text="🔍 Consulta Detallada RUT seleccionada (Completa)")
            self.add_result_message("🔍 Modo Detallado: Consulta completa con estado del registro", "INFO")
    
    def update_stats(self, total, exitosos, errores):
        """Actualiza las estadísticas en tiempo real"""
        self.label_total.config(text=f"Total: {total}")
        self.label_exitosos.config(text=f"✅ Exitosos: {exitosos}")
        self.label_errores.config(text=f"❌ Errores: {errores}")
    
    def add_result_message(self, message, level="INFO"):
        """Agrega mensaje a resultados"""
        self.text_resultados.config(state=tk.NORMAL)
        
        if level == "SUCCESS":
            prefix = "✅"
        elif level == "ERROR":
            prefix = "❌"
        elif level == "WARNING":
            prefix = "⚠️"
        else:
            prefix = "ℹ️"
        
        formatted_message = f"{prefix} {message}\n"
        self.text_resultados.insert(tk.END, formatted_message)
        self.text_resultados.see(tk.END)
        self.text_resultados.config(state=tk.DISABLED)
        self.update()
    
    # ═══ MÉTODOS DE FUNCIONALIDAD DUAL - SIN CAMBIOS ═══
    
    def on_close(self):
        cerrar_todos_los_navegadores()
        self.destroy()
    
    def reset_sistema(self):
        cerrar_todos_los_navegadores()
        self.lista_nits = []
        self.rows_for_excel = []
        self.progress_bar["value"] = 0
        self.label_status.config(text="Listo para comenzar")
        self.update_stats(0, 0, 0)
        
        self.text_resultados.config(state=tk.NORMAL)
        self.text_resultados.delete("1.0", tk.END)
        self.text_resultados.insert(tk.END, 
                                   "🔄 Sistema reiniciado\n"
                                   "📁 Seleccione nuevo archivo\n\n")
        self.text_resultados.config(state=tk.DISABLED)
        
        self.generated_file = None
        self.btn_open_excel.pack_forget()
        self.btn_detener.config(state=tk.DISABLED)
        self.detener_cronometro()
    
    def on_cargar_excel(self):
        self.reset_sistema()
        
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel o CSV",
            initialdir=os.path.expanduser("~/Documents"),
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Archivos CSV", "*.csv"),
                ("Todos", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        self.add_result_message(f"Archivo: {os.path.basename(filepath)}", "INFO")
        
        try:
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
            
            self.add_result_message(f"Leído: {len(df)} filas", "SUCCESS")
            
        except Exception as e:
            self.add_result_message(f"Error: {str(e)}", "ERROR")
            return
        
        if "NIT" not in df.columns:
            self.add_result_message("Falta columna 'NIT'", "ERROR")
            return
        
        try:
            nits_raw = df["NIT"].dropna().astype(str).tolist()
            nits_procesados = []
            for nit in nits_raw:
                if '.' in nit:
                    nit = str(int(float(nit)))
                nit_limpio = ''.join(c for c in nit if c.isdigit())
                if nit_limpio and len(nit_limpio) >= 1:
                    nits_procesados.append(nit_limpio)
            
            self.lista_nits = list(dict.fromkeys(nits_procesados))
            
            if not self.lista_nits:
                self.add_result_message("No hay NITs válidos", "ERROR")
                return
            
            total = len(self.lista_nits)
            self.progress_bar["maximum"] = total
            self.progress_bar["value"] = 0
            self.update_stats(total, 0, 0)
            
            tipo = self.tipo_consulta.get()
            tipo_texto = "Express" if tipo == "basica" else "Detallada RUT"
            
            self.label_status.config(text=f"Iniciando consulta {tipo_texto}: {total} NITs")
            self.add_result_message(f"Procesando {total} NITs - Tipo: {tipo_texto}", "SUCCESS")
            self.add_result_message("═══ INICIANDO PROCESAMIENTO ═══", "INFO")
            
            self.iniciar_cronometro()
            threading.Thread(target=self.consultar_nits_dual_gui, daemon=True).start()
            
        except Exception as e:
            self.add_result_message(f"Error: {str(e)}", "ERROR")
    
    def iniciar_cronometro(self):
        self.tiempo_inicio = time.time()
        self.cronometro_activo = True
        self.actualizar_cronometro()
    
    def detener_cronometro(self):
        self.cronometro_activo = False
    
    def actualizar_cronometro(self):
        if self.cronometro_activo and self.tiempo_inicio:
            elapsed = time.time() - self.tiempo_inicio
            tiempo_str = str(timedelta(seconds=int(elapsed)))
            self.label_cronometro.config(text=f"⏱️ Tiempo: {tiempo_str}")
            self.after(1000, self.actualizar_cronometro)
    
    def actualizar_progreso(self, count, total, nit=None):
        self.progress_bar["value"] = count
        if nit:
            tipo = self.tipo_consulta.get()
            tipo_texto = "Express" if tipo == "basica" else "RUT"
            self.label_status.config(text=f"Consultando {tipo_texto}: {nit} ({count}/{total})")
        else:
            self.label_status.config(text=f"Procesando: {count}/{total}")
    
    def detener_consulta(self):
        if self.ejecucion_activa.is_set():
            self.detener_proceso = True
            self.ejecucion_activa.clear()
            self.btn_detener.config(state=tk.DISABLED)
            self.label_status.config(text="Deteniendo...")
            self.add_result_message("Proceso detenido", "WARNING")
            self.update()
            self.generar_excel_parcial()
        else:
            messagebox.showinfo("Info", "No hay proceso activo")
    
    def consultar_nits_dual_gui(self):
        """FUNCIÓN PRINCIPAL DUAL - Maneja ambos tipos de consulta"""
        self.ejecucion_activa.set()
        self.detener_proceso = False
        self.btn_detener.config(state=tk.NORMAL)
        self.rows_for_excel = []
        
        try:
            if not self.lista_nits:
                self.add_result_message("No hay NITs", "ERROR")
                return
            
            total_nits = len(self.lista_nits)
            exitosos = 0
            errores = 0
            
            # Obtener tipo de consulta seleccionado
            tipo_seleccionado = self.tipo_consulta.get()
            tipo_texto = "Express" if tipo_seleccionado == "basica" else "Detallada RUT (Completa)"
            
            self.add_result_message(f"🚀 Procesando {total_nits} NITs - Modo: {tipo_texto}", "SUCCESS")
            self.add_result_message(f"🏛️ Inicializando sistema de consulta {tipo_texto.lower()}", "INFO")
            
            for i, nit in enumerate(self.lista_nits, 1):
                if not self.ejecucion_activa.is_set() or self.detener_proceso:
                    self.add_result_message(f"Proceso interrumpido en consulta {i}/{total_nits}", "WARNING")
                    break
                
                # Mensajes profesionales con rotación
                if i % 8 == 0:
                    mensaje_pro = obtener_mensaje_profesional()
                    self.add_result_message(mensaje_pro, "INFO")
                    limpiar_navegadores_inactivos()
                
                if i % 4 == 1 and i > 1:
                    tip_contable = obtener_tip_contable()
                    self.add_result_message(tip_contable, "INFO")
                
                self.actualizar_progreso(i, total_nits, nit)
                
                # ═══ CONSULTA DUAL - USAR FUNCIÓN COORDINADORA ═══
                resultado = consultar_nit_individual(nit, tipo_seleccionado, attempt=1)
                
                # Manejo de reintentos
                if resultado.get("status") == "retry":
                    tipo_msg = "Express" if tipo_seleccionado == "basica" else "RUT detallada"
                    self.add_result_message(f"🔄 Verificando nuevamente información {tipo_msg} de NIT {nit}", "WARNING")
                    time.sleep(1.0 if tipo_seleccionado == "basica" else 2.5)  # OPTIMIZADO PARA Express
                    resultado = consultar_nit_individual(nit, tipo_seleccionado, attempt=2)
                    
                    if resultado.get("status") == "retry":
                        self.add_result_message(f"🔍 Validación adicional {tipo_msg} para NIT {nit}", "WARNING")
                        time.sleep(1.5 if tipo_seleccionado == "basica" else 3.5)  # OPTIMIZADO PARA Express
                        resultado = consultar_nit_individual(nit, tipo_seleccionado, attempt=3)
                
                if resultado.get("status") == "success":
                    data = resultado["data"]
                    
                    def limpiar_campo(valor):
                        if not valor or str(valor).strip() == "" or str(valor).strip() == "Sin inconsistencias registradas" or str(valor).strip() == "None":
                            return "-"
                        return str(valor).strip()
                    
                    # ═══ ESTRUCTURA DUAL DE EXCEL ═══
                    if tipo_seleccionado == "basica":
                        # Estructura para consulta Express (orden original corregido)
                        fila_excel = {
                            "NIT": nit,
                            "DV": data.get("dv", calcular_dv(nit)),
                            "Primer Apellido": limpiar_campo(data.get("primerNombre", "")),
                            "Segundo Apellido": limpiar_campo(data.get("otrosNombres", "")),
                            "Primer Nombre": limpiar_campo(data.get("primerApellido", "")),
                            "Otros Nombres": limpiar_campo(data.get("segundoApellido", "")),
                            "Razón Social": limpiar_campo(data.get("razonSocial", "")),
                            "Fecha Consulta": data.get("datetime", ""),
                            "Estado Consulta": "Exitoso",
                            "Tipo de Consulta": "Express",
                            "Observaciones": data.get("observacion", "Consulta Express exitosa")
                        }
                    else:
                        # Estructura para consulta RUT detallada
                        fila_excel = {
                            "NIT": nit,
                            "DV": limpiar_campo(data.get("dv", calcular_dv(nit))),
                            "Primer Apellido": limpiar_campo(data.get("primerApellido", "")),
                            "Segundo Apellido": limpiar_campo(data.get("segundoApellido", "")),
                            "Primer Nombre": limpiar_campo(data.get("primerNombre", "")),
                            "Otros Nombres": limpiar_campo(data.get("otrosNombres", "")),
                            "Razón Social": limpiar_campo(data.get("razonSocial", "")),
                            "Estado del Registro": limpiar_campo(data.get("estado", "SIN INFORMACIÓN")),
                            "Fecha Consulta": data.get("datetime", ""),
                            "Estado Consulta": "Exitoso",
                            "Tipo de Consulta": "RUT Detallado",
                            "Observaciones": data.get("observacion", "Consulta RUT detallada exitosa")
                        }
                    
                    self.rows_for_excel.append(fila_excel)
                    
                    # Mostrar resultado según tipo de consulta
                    if tipo_seleccionado == "basica":
                        razon_social = data.get('razonSocial', '')
                        nombre_completo = f"{data.get('primerNombre', '')} {data.get('primerApellido', '')}".strip()
                        
                        if razon_social and razon_social not in ["-", "", "Sin inconsistencias registradas"]:
                            display_text = razon_social[:30] + "..." if len(razon_social) > 30 else razon_social
                        elif nombre_completo and nombre_completo not in ["-", "", "Sin inconsistencias registradas"]:
                            display_text = nombre_completo[:30] + "..." if len(nombre_completo) > 30 else nombre_completo
                        else:
                            display_text = "Sin datos registrados"
                    else:
                        # Para RUT detallado, mostrar razón social o nombre y estado
                        razon_social = data.get('razonSocial', '')
                        nombre_completo = f"{data.get('primerNombre', '')} {data.get('primerApellido', '')}".strip()
                        estado = data.get('estado', 'SIN INFORMACIÓN')
                        
                        if razon_social and razon_social not in ["-", "", "Sin inconsistencias registradas", "No está inscrito en el RUT"]:
                            display_text = f"{razon_social[:20]}... | {estado}"
                        elif nombre_completo and nombre_completo not in ["-", "", "Sin inconsistencias registradas", "No está inscrito en el RUT"]:
                            display_text = f"{nombre_completo[:20]}... | {estado}"
                        else:
                            display_text = f"Sin datos | {estado}"
                    
                    self.add_result_message(f"✅ {nit}: {display_text}", "SUCCESS")
                    exitosos += 1
                    
                else:
                    error_msg = resultado.get("error", "Error")
                    
                    # Estructura de error según tipo de consulta
                    if tipo_seleccionado == "basica":
                        fila_excel = {
                            "NIT": nit,
                            "DV": "-",
                            "Primer Apellido": "-",
                            "Segundo Apellido": "-", 
                            "Primer Nombre": "-",
                            "Otros Nombres": "-",
                            "Razón Social": "-",
                            "Fecha Consulta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Estado Consulta": "No Inscrito",
                            "Tipo de Consulta": "Express",
                            "Observaciones": "No Inscrito" 
                            }
                    
                    else:
                        fila_excel = {
                            "NIT": nit,
                            "DV": "-",
                            "Primer Apellido": "-",
                            "Segundo Apellido": "-", 
                            "Primer Nombre": "-",
                            "Otros Nombres": "-",
                            "Razón Social": "-",
                            "Estado del Registro": "-",
                            "Fecha Consulta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Estado Consulta": "No Inscrito",
                            "Tipo de Consulta": "RUT Detallado",
                            "Observaciones": "No Inscrito"
                        }
                    
                    self.rows_for_excel.append(fila_excel)
                    error_short = error_msg[:25] + "..." if len(error_msg) > 25 else error_msg
                    self.add_result_message(f"⚠️ {nit}: Requiere verificación manual - {error_short}", "ERROR")
                    errores += 1
                
                # Actualizar estadísticas
                self.update_stats(total_nits, exitosos, errores)
                
                # Pausa entre consultas OPTIMIZADA
                if i < total_nits:
                    pausa = TimeoutConfig.BROWSER_REST_TIME  # OPTIMIZADO: 0.2s
                    if tipo_seleccionado == "rut_detallado":
                        pausa += 0.3  # Pausa adicional para RUT (sin cambio)
                    time.sleep(pausa)
            
            # Resumen final
            self.add_result_message("═══ PROCESO COMPLETADO ═══", "INFO")
            self.add_result_message(f"📊 Consultas {tipo_texto}: {i} | ✅ Exitosas: {exitosos} | ⚠️ Requieren revisión: {errores}", "INFO")
            self.add_result_message(f"🎯 Sistema profesional de consulta {tipo_texto.lower()} operado exitosamente", "INFO")
            
            if self.ejecucion_activa.is_set() and not self.detener_proceso:
                self.add_result_message("📈 Generando reporte ejecutivo automatizado", "SUCCESS")
                self.generar_excel_completo()
                self.after(1000, self.open_excel)
            
        except Exception as e:
            self.add_result_message(f"Error en el sistema: {str(e)}", "ERROR")
        finally:
            self.ejecucion_activa.clear()
            self.detener_proceso = False
            self.btn_detener.config(state=tk.DISABLED)
            self.detener_cronometro()
            cerrar_todos_los_navegadores()
            self.label_status.config(text="Proceso completado")
    
    def generar_excel_parcial(self):
        """Genera Excel parcial con formato dual"""
        try:
            if not self.rows_for_excel:
                self.add_result_message("No hay datos para Excel", "WARNING")
                return
            
            df_out = pd.DataFrame(self.rows_for_excel)
            
            if 'NIT' in df_out.columns:
                df_out['NIT'] = pd.to_numeric(df_out['NIT'], errors='coerce').fillna(0).astype('int64')
            if 'DV' in df_out.columns:
                df_out['DV'] = pd.to_numeric(df_out['DV'], errors='coerce').fillna(0).astype('int64')
            
            tipo = self.tipo_consulta.get()
            tipo_texto = "basica" if tipo == "basica" else "rut_detallado"
            fname = f"resultado_parcial_{tipo_texto}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            
            with pd.ExcelWriter(fname, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Resultados', startrow=2)
                
                workbook = writer.book
                worksheet = writer.sheets['Resultados']
                
                self.apply_excel_formatting_dual(workbook, worksheet, df_out, "PARCIAL", tipo)
            
            self.generated_file = fname
            self.add_result_message(f"Excel generado: {fname}", "SUCCESS")
            self.show_open_excel_button()
            
        except Exception as e:
            self.add_result_message(f"Error Excel: {str(e)}", "ERROR")
        finally:
            self.label_status.config(text="Detenido")
            cerrar_todos_los_navegadores()
    
    def generar_excel_completo(self):
        """Genera Excel completo con formato dual"""
        try:
            if not self.rows_for_excel:
                self.add_result_message("No hay datos", "WARNING")
                return
            
            df_out = pd.DataFrame(self.rows_for_excel)
            
            if 'NIT' in df_out.columns:
                df_out['NIT'] = pd.to_numeric(df_out['NIT'], errors='coerce').fillna(0).astype('int64')
            if 'DV' in df_out.columns:
                df_out['DV'] = pd.to_numeric(df_out['DV'], errors='coerce').fillna(0).astype('int64')
            
            tipo = self.tipo_consulta.get()
            tipo_texto = "express" if tipo == "basica" else "rut_detallado"
            fname = f"resultado_completo_{tipo_texto}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            
            with pd.ExcelWriter(fname, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Resultados', startrow=2)
                
                workbook = writer.book
                worksheet = writer.sheets['Resultados']
                
                self.apply_excel_formatting_dual(workbook, worksheet, df_out, "COMPLETO", tipo)
            
            self.generated_file = fname
            self.add_result_message(f"Excel completo: {fname}", "SUCCESS")
            self.add_result_message("🚀 Abriendo Excel automáticamente...", "INFO")
            self.show_open_excel_button()
            
        except Exception as e:
            self.add_result_message(f"Error: {str(e)}", "ERROR")
    
    def apply_excel_formatting_dual(self, workbook, worksheet, df_out, tipo="COMPLETO", tipo_consulta="basica"):
        """Aplica formato profesional dual al Excel"""
        
        # Títulos diferenciados por tipo de consulta
        if tipo == "PARCIAL":
            if tipo_consulta == "basica":
                title_format = workbook.add_format({
                    'bold': True, 'font_size': 16, 'font_color': '#D63384',
                    'align': 'center', 'valign': 'vcenter', 'bg_color': '#FDF2F8',
                    'font_name': 'Calibri'
                })
                titulo = f'[PARCIAL - Express] Consulta DIAN Profesional - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
            else:
                title_format = workbook.add_format({
                    'bold': True, 'font_size': 16, 'font_color': '#1565C0',
                    'align': 'center', 'valign': 'vcenter', 'bg_color': '#E3F2FD',
                    'font_name': 'Calibri'
                })
                titulo = f'[PARCIAL - RUT DETALLADO] Consulta DIAN Profesional - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
        else:
            if tipo_consulta == "basica":
                title_format = workbook.add_format({
                    'bold': True, 'font_size': 16, 'font_color': '#1B5E20',
                    'align': 'center', 'valign': 'vcenter', 'bg_color': '#E8F5E9',
                    'font_name': 'Calibri'
                })
                titulo = f'Consulta Gestión Masiva DIAN Express- A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
            else:
                title_format = workbook.add_format({
                    'bold': True, 'font_size': 16, 'font_color': '#0D47A1',
                    'align': 'center', 'valign': 'vcenter', 'bg_color': '#E1F5FE',
                    'font_name': 'Calibri'
                })
                titulo = f'Consulta Gestión Masiva DIAN RUT DETALLADO - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
        
        # Headers diferenciados
        if tipo_consulta == "basica":
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#1B5E20', 'font_color': 'white', 'border': 2,
                'border_color': '#2E7D32', 'font_size': 11, 'font_name': 'Calibri'
            })
        else:
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#0D47A1', 'font_color': 'white', 'border': 2,
                'border_color': '#1976D2', 'font_size': 11, 'font_name': 'Calibri'
            })
        
        # Formatos de datos
        num_format = workbook.add_format({
            'num_format': '0', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'border_color': '#9E9E9E', 'font_size': 10,
            'font_name': 'Calibri', 'bg_color': '#FFFFFF'
        })
        
        text_format = workbook.add_format({
            'valign': 'vcenter', 'border': 1, 'border_color': '#9E9E9E',
            'font_size': 10, 'font_name': 'Calibri', 'bg_color': '#FFFFFF'
        })
        text_format_par = workbook.add_format({
            'valign': 'vcenter', 'border': 1, 'border_color': '#9E9E9E',
            'font_size': 10, 'font_name': 'Calibri', 'bg_color': '#EAF4EA'
        })
        num_format_par = workbook.add_format({
            'num_format': '0', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'border_color': '#9E9E9E', 'font_size': 10,
            'font_name': 'Calibri', 'bg_color': '#EAF4EA'
        })
        
        success_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1,
            'border_color': '#145214', 'bg_color': '#1B5E20',
            'font_color': '#FFFFFF', 'bold': True, 'font_size': 10, 'font_name': 'Calibri'
        })
        
        error_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1,
            'border_color': '#7B1212', 'bg_color': '#B71C1C',
            'font_color': '#FFFFFF', 'bold': True, 'font_size': 10, 'font_name': 'Calibri'
        })
        no_inscrito_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1,
            'border_color': '#BF360C', 'bg_color': '#E65100',
            'font_color': '#FFFFFF', 'bold': True, 'font_size': 10, 'font_name': 'Calibri'
        })
        
        # Formato especial para Estado del Registro (solo RUT)
        if tipo_consulta == "rut_detallado":
            estado_activo_format = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1,
                'border_color': '#E1F5FE', 'bg_color': '#B3E5FC',
                'font_color': '#0D47A1', 'bold': True, 'font_size': 10, 'font_name': 'Calibri'
            })
            
            estado_inactivo_format = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1,
                'border_color': '#E1F5FE', 'bg_color': '#FFECB3',
                'font_color': '#E65100', 'bold': True, 'font_size': 10, 'font_name': 'Calibri'
            })
        
        # Escribir título
        worksheet.merge_range(0, 0, 0, len(df_out.columns) - 1, titulo, title_format)
        worksheet.set_row(0, 35)
        
        # Configurar anchos optimizados según tipo de consulta
        if tipo_consulta == "basica":
            column_widths = {
                'NIT': 15, 'DV': 5, 'Primer Apellido': 20, 'Segundo Apellido': 20,
                'Primer Nombre': 20, 'Otros Nombres': 16, 'Razón Social': 38,
                'Fecha Consulta': 19, 'Estado Consulta': 14, 'Tipo de Consulta': 10,
                'Observaciones': 22
            }
        else:
            column_widths = {
                'NIT': 15, 'DV': 5, 'Primer Apellido': 20, 'Segundo Apellido': 20,
                'Primer Nombre': 20, 'Otros Nombres': 16, 'Razón Social': 32,
                'Estado del Registro': 22, 'Fecha Consulta': 19, 'Estado Consulta': 14,
                'Tipo de Consulta': 10, 'Observaciones': 22
            }
        
        for i, col_name in enumerate(df_out.columns):
            width = column_widths.get(col_name, 15)
            worksheet.set_column(i, i, width)
        
        # Headers
        for col_num, col_name in enumerate(df_out.columns):
            worksheet.write(2, col_num, col_name, header_format)
        
        # Datos con filas alternas
        for row_num in range(len(df_out)):
            es_par = (row_num % 2 == 0)
            tf = text_format_par if es_par else text_format
            nf = num_format_par  if es_par else num_format

            for col_num, col_name in enumerate(df_out.columns):
                cell_value = df_out.iloc[row_num, col_num]
                actual_row = row_num + 3

                if col_name in ['NIT', 'DV']:
                    cell_format = nf
                elif col_name == 'Estado Consulta':
                    val_lower = str(cell_value).lower()
                    if val_lower == 'exitoso':
                        cell_format = success_format
                    else:
                        cell_format = no_inscrito_format
                elif col_name == 'Estado del Registro' and tipo_consulta == "rut_detallado":
                    estado_val = str(cell_value).upper()
                    if 'ACTIVO' in estado_val:
                        cell_format = estado_activo_format
                    elif 'SUSPENDIDO' in estado_val:
                        cell_format = no_inscrito_format
                    elif 'ERROR' in estado_val or 'SIN' in estado_val:
                        cell_format = estado_inactivo_format
                    else:
                        cell_format = tf
                else:
                    cell_format = tf

                worksheet.write(actual_row, col_num, cell_value, cell_format)

        # Configuraciones finales
        worksheet.freeze_panes(3, 0)
        worksheet.autofilter(2, 0, len(df_out) + 2, len(df_out.columns) - 1)
        worksheet.set_default_row(18)
        worksheet.set_row(2, 32)
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)
    
    def show_open_excel_button(self):
        self.btn_open_excel.pack(side=tk.LEFT, padx=(0, 10))
        self.update()
    
    def open_excel(self):
        if self.generated_file and os.path.exists(self.generated_file):
            try:
                abs_path = os.path.abspath(self.generated_file)
                if sys.platform.startswith('win'):
                    os.startfile(abs_path)
                elif sys.platform.startswith('darwin'):
                    os.system(f'open "{abs_path}"')
                else:
                    webbrowser.open(f"file:///{abs_path}")
                
                self.add_result_message(f"Abriendo: {self.generated_file}", "SUCCESS")
                
            except Exception as e:
                self.add_result_message(f"Error abriendo: {str(e)}", "ERROR")
        else:
            self.add_result_message("Archivo no encontrado", "ERROR")

def main():
    try:
        print("🚀 === CONSULTA DIAN - A.S. CONTADORES & ASESORES SAS ===")
        print("🎯 Iniciando sistema dual con consulta Express y detallada RUT...")
        print("✨ Navegadores configurados para ser invisibles al usuario...")
        app = ConsultaRUTApp()
        print("✅ Sistema profesional dual iniciado exitosamente")
        app.mainloop()
    except Exception as e:
        print(f"❌ Error crítico: {e}")
    finally:
        print("🧹 Cerrando sistema...")
        cerrar_todos_los_navegadores()
        print("✅ Sistema cerrado correctamente")

if __name__ == "__main__":
    main()