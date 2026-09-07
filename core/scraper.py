"""
Lógica de scraping DIAN — Consulta Express y RUT Detallado.
"""

import logging
import random
import re
import time
from datetime import datetime

import requests

log = logging.getLogger(__name__)
from config import (
    DIAN_URL_BASICA, SEL_NIT_ID_BASICA, SEL_DV_ID_BASICA, BTN_BUSCAR_ID_BASICA,
    FIELDS_BASICA, DIAN_URL_RUT, FIELDS_RUT, ERROR_CSS, TimeoutConfig
)
from core.browser import get_browser, return_browser, crear_navegador


# ─────── Helpers JS ───────

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
        return driver.run_js(js) == value
    except Exception as e:
        log.warning(f"set_field_js {element_id}: {e}")
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
        log.warning(f"click_js {element_id}: {e}")
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
        log.warning(f"calcular_dv NIT {nit}: {e}")
        return "0"


# ─────── Consulta Express ───────

def _check_results_basica(driver):
    try:
        js = f"""
        const nombre = document.getElementById('{FIELDS_BASICA['primerNombre']}');
        const razon  = document.getElementById('{FIELDS_BASICA['razonSocial']}');
        if ((nombre && nombre.textContent.trim()) || (razon && razon.textContent.trim()))
            return 'success';
        const err = document.querySelector('{ERROR_CSS}');
        if (err && err.textContent.trim())
            return 'error:' + err.textContent.trim();
        return 'waiting';
        """
        return driver.run_js(js)
    except:
        return 'waiting'


def _extract_basica(driver):
    data = {}
    for key, eid in FIELDS_BASICA.items():
        try:
            result = driver.run_js(f"""
            const el = document.getElementById('{eid}');
            return el ? el.textContent.trim() : null;
            """)
            data[key] = result if result else None
        except:
            data[key] = None
    return data


def _reset_basica(driver):
    try:
        return driver.run_js(f"""
        document.getElementById('{SEL_NIT_ID_BASICA}').value = '';
        document.getElementById('{SEL_DV_ID_BASICA}').value = '';
        return true;
        """)
    except:
        return False


def _check_no_inconsistencias(driver):
    try:
        return driver.run_js("""
        const elems = document.querySelectorAll('.ui-dialog-content p, .ui-messages-info-detail, .ui-growl-message p');
        for (let e of elems) {
            const t = e.textContent.toLowerCase();
            if (t.includes('no se encontraron') || t.includes('sin inconsistencias')) {
                const btn = document.querySelector('.ui-dialog-titlebar-close, .ui-growl-icon-close');
                if (btn) btn.click();
                return true;
            }
        }
        return false;
        """)
    except:
        return False


def consultar_nit_basica(nit: str, attempt: int = 1):
    driver = None
    try:
        driver = get_browser()
        try:
            _ = driver.url
        except:
            try:
                driver.quit()
            except:
                pass
            driver = crear_navegador(0)[0]

        if not driver.url or DIAN_URL_BASICA not in driver.url:
            driver.get(DIAN_URL_BASICA)
            time.sleep(TimeoutConfig.INITIAL_WAIT)

        try:
            from CloudflareBypasser import CloudflareBypasser
            CloudflareBypasser(driver, max_retries=2, log=False).bypass()
            time.sleep(TimeoutConfig.CF_BYPASS_WAIT)
        except Exception as e:
            log.debug(f"Cloudflare bypass: {e}")

        for i in range(2):
            if _reset_basica(driver):
                break
            time.sleep(0.3)

        time.sleep(0.2)

        nit_ok = False
        for _ in range(TimeoutConfig.MAX_RETRIES):
            if set_field_js(driver, SEL_NIT_ID_BASICA, nit):
                nit_ok = True
                break
            time.sleep(0.3)
        if not nit_ok:
            return {"status": "error", "data": {}, "error": f"No se pudo establecer NIT {nit}"}

        time.sleep(TimeoutConfig.POST_NIT_WAIT)

        dv = calcular_dv(nit)
        for _ in range(TimeoutConfig.MAX_RETRIES):
            if set_field_js(driver, SEL_DV_ID_BASICA, dv):
                break
            time.sleep(0.3)

        time.sleep(0.2)

        buscar_ok = False
        for _ in range(TimeoutConfig.MAX_RETRIES):
            if click_js(driver, BTN_BUSCAR_ID_BASICA):
                buscar_ok = True
                break
            time.sleep(0.3)
        if not buscar_ok:
            return {"status": "error", "data": {}, "error": "No se pudo hacer clic en Buscar"}

        time.sleep(0.8)

        waited, status, no_data = 0, 'waiting', False
        while status == 'waiting' and waited < TimeoutConfig.RESULTS_WAIT:
            if _check_no_inconsistencias(driver):
                no_data = True
                status = 'success_no_data'
                break
            status = _check_results_basica(driver)
            if status == 'waiting':
                time.sleep(0.5)
                waited += 0.5

        if status.startswith('error:'):
            return {"status": "error", "data": {}, "error": status.split('error:', 1)[1]}
        if status == 'waiting':
            return {"status": "retry" if attempt == 1 else "error", "data": {},
                    "error": f"Timeout para NIT {nit}"}

        data = _extract_basica(driver)
        data['nit'] = nit
        data['dv'] = dv
        data['datetime'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data['tipo_consulta'] = 'basica'
        if attempt == 2:
            data['observacion'] = "Consulta exitosa en segundo intento"
        if no_data or (not data.get("razonSocial") and not data.get("primerNombre")):
            for k in ["razonSocial", "primerNombre", "primerApellido", "segundoApellido", "otrosNombres"]:
                data[k] = "-"

        return {"status": "success", "data": data, "error": None}

    except Exception as e:
        log.warning(f"Error consulta Express {nit}: {e}")
        return {"status": "error", "data": {}, "error": str(e)}
    finally:
        if driver:
            return_browser(driver)


# ─────── Consulta RUT Detallado ───────

def _set_nit_rut(driver, nit):
    try:
        js = f"""
        const input = document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit');
        input.value = '{nit}';
        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
        return input.value;
        """
        return driver.run_js(js) == nit
    except Exception as e:
        log.warning(f"set_nit_rut: {e}")
        return False


def _click_buscar_rut(driver):
    try:
        return driver.run_js("""
        document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar').click();
        return true;
        """)
    except:
        return False


def _check_results_rut(driver):
    try:
        return driver.run_js("""
        const dv = document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv');
        const err = document.querySelector('.ui-messages-error-detail');
        if (dv && dv.textContent.trim()) return 'success';
        if (err) return 'error:' + err.textContent.trim();
        return 'waiting';
        """)
    except:
        return 'waiting'


def _extract_rut(driver):
    fields = {
        "dv":             "vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv",
        "razonSocial":    "vistaConsultaEstadoRUT:formConsultaEstadoRUT:razonSocial",
        "primerNombre":   "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre",
        "otrosNombres":   "vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres",
        "primerApellido": "vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido",
        "segundoApellido":"vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido",
        "estado":         "vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado",
    }
    data = {}
    for key, sel in fields.items():
        try:
            result = driver.run_js(f"""
            const el = document.getElementById('{sel}');
            return el ? el.textContent.trim() : null;
            """)
            data[key] = result if result else None
        except:
            data[key] = None
    return data


def _reset_rut(driver):
    try:
        return driver.run_js("""
        document.getElementById('vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit').value = '';
        return true;
        """)
    except:
        return False


def _close_error_rut(driver):
    try:
        return driver.run_js("""
        const t = document.querySelector("table[background*='fondoMensajeError.gif']");
        if (t) { const b = t.querySelector("img[src*='botcerrarrerror.gif']"); if (b) { b.click(); return true; } }
        return false;
        """)
    except:
        return False


def _check_captcha_rut(driver):
    try:
        return driver.run_js("""
        const c = document.getElementById('g-recaptcha-error');
        return c && c.innerText.includes('Se requiere validar captcha.');
        """)
    except:
        return False


def consultar_nit_rut_detallado(nit: str, attempt: int = 1):
    driver = None
    try:
        driver = get_browser()
        try:
            _ = driver.url
        except:
            try:
                driver.quit()
            except:
                pass
            driver = crear_navegador(0)[0]

        if not driver.url or DIAN_URL_RUT not in driver.url:
            driver.get(DIAN_URL_RUT)
            time.sleep(TimeoutConfig.INITIAL_WAIT_RUT)

        _close_error_rut(driver)
        _reset_rut(driver)

        retry, nit_set = 0, False
        while retry < TimeoutConfig.MAX_RETRIES and not nit_set:
            nit_set = _set_nit_rut(driver, nit)
            if not nit_set:
                retry += 1
                time.sleep(TimeoutConfig.INITIAL_WAIT_RUT * (TimeoutConfig.BACKOFF_FACTOR ** retry))
        if not nit_set:
            return {"status": "error", "data": {}, "error": f"No se pudo establecer NIT {nit} en RUT"}

        time.sleep(TimeoutConfig.POST_NIT_WAIT_RUT)

        retry, clicked = 0, False
        while retry < TimeoutConfig.MAX_RETRIES and not clicked:
            clicked = _click_buscar_rut(driver)
            if not clicked:
                retry += 1
                time.sleep(TimeoutConfig.INITIAL_WAIT_RUT * (TimeoutConfig.BACKOFF_FACTOR ** retry))
        if not clicked:
            return {"status": "error", "data": {}, "error": "No se pudo hacer clic en Buscar RUT"}

        time.sleep(1)

        if _check_captcha_rut(driver):
            if attempt > 2:
                return {"status": "error", "data": {}, "error": f"Captcha detectado (intento {attempt})"}
            return {"status": "retry", "data": {}, "error": f"Captcha detectado, reintentando..."}

        waited, result_status = 0, "waiting"
        max_wait = TimeoutConfig.RESULTS_WAIT_RUT * 0.8
        while result_status == "waiting" and waited < max_wait:
            result_status = _check_results_rut(driver)
            if result_status == "waiting":
                time.sleep(1)
                waited += 1

        if result_status.startswith("error:"):
            return {"status": "error", "data": {}, "error": result_status.split("error:", 1)[1]}
        if result_status == "waiting":
            if attempt == 1:
                return {"status": "retry", "data": {}, "error": f"Timeout RUT NIT {nit}"}
            return {"status": "error", "data": {}, "error": f"{nit}: No inscrito en RUT"}

        data = _extract_rut(driver)
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
        log.warning(f"Error RUT {nit}: {e}")
        return {"status": "error", "data": {}, "error": str(e)}
    finally:
        if driver:
            return_browser(driver)


# ─────── TECNOPOS (sipos.com.co) — acelerador interno para Express ───────
# Fuente NO oficial de un tercero: HTTP puro (sin navegador), ~0.3-0.4s por
# consulta vs varios segundos de DIAN Express. Se usa como intento previo
# silencioso: cualquier caso no confiable retorna None y el coordinador cae
# automáticamente a consultar_nit_basica() (DIAN real). Nunca se expone al
# usuario de dónde vino el dato — mismo Excel, mismos campos, sin marcas.
#
# Riesgo conocido y medido (ver BITÁCORA): la respuesta de TECNOPOS es
# inestable (a veces no trae razonSocial/dv) y no separa nombre/apellido —
# eso se resuelve aquí con la misma heurística validada en consulta-cedulas
# (es_empresa + separación por conteo de palabras), con fix de 'LIMITADA' y
# 'SOCIEDAD' agregado tras medir contra datos reales (tasa de error 3.96%
# antes del fix, 0% después, sobre muestra de 101 razones sociales reales).
# TECNOPOS nunca trae el estado del registro RUT — el modo RUT Detallado no
# usa nada de esto.

TECNOPOS_URL      = 'https://sipos.com.co/api_rut.php'
TECNOPOS_REFERER  = 'https://sipos.com.co/consultarut'
TECNOPOS_TIMEOUT  = 5

# Sufijos societarios que indican EMPRESA, no persona natural. El chequeo es
# por palabra individual, no por frase — por eso 'SOCIEDAD' cubre tanto
# "SOCIEDAD ANONIMA" como "SOCIEDAD POR ACCIONES SIMPLIFICADA" sin tener que
# listar la frase completa (que nunca matchearía palabra por palabra).
SUFIJOS_EMPRESA = {'SAS', 'SA', 'LTDA', 'LIMITADA', 'CIA', 'EU', 'EIRL', 'ESE',
                    'IPS', 'FUNDACION', 'COOPERATIVA', 'ASOCIACION', 'ONG',
                    'SOCIEDAD', 'ESP'}

# Partículas de apellido compuesto: su sola presencia hace incierto dónde
# termina el apellido con una regla de conteo simple, así que se trata como
# dudoso en vez de intentar pegarlas "inteligentemente".
_PARTICULAS = {'DE', 'DEL', 'LA', 'LAS', 'LOS', 'SAN', 'SANTA',
               'VAN', 'VON', 'MAC', 'MC'}

# Nombres genéricos de POS/facturación que no representan a una persona real.
_PLACEHOLDERS = {'VENTA MOSTRADOR', 'CONSUMIDOR FINAL', 'CLIENTE VARIOS',
                  'PUBLICO EN GENERAL', 'CLIENTE OCASIONAL', 'SIN NOMBRE',
                  'NO APLICA', 'VARIOS'}

_PALABRA_RE = re.compile(r'^[A-ZÁÉÍÓÚÑÜ]+$')

_tecnopos_session = None


def es_empresa(razon_social: str) -> bool:
    """True si alguna palabra es un sufijo societario (SAS, LTDA, S.A., ...).
    Normaliza puntos ("S.A." -> "SA") antes de comparar.

    TECNOPOS a veces separa las siglas con espacio en vez de punto (ej.
    "COLOMBIA MOVIL S A ESP" en vez de "S.A. E.S.P" — confirmado en vivo,
    NIT 830114921). Eso deja "S" y "A" como palabras sueltas de una letra
    que nunca matchean 'SA' completo, así que además se unen corridas de
    palabras de una sola letra ("S", "A" -> "SA") antes de comparar."""
    palabras = (razon_social or '').strip().upper().split()
    limpias = [p.replace('.', '') for p in palabras]
    if any(p in SUFIJOS_EMPRESA for p in limpias):
        return True

    siglas, actual = [], ''
    for p in limpias:
        if len(p) == 1 and p.isalpha():
            actual += p
        else:
            if actual:
                siglas.append(actual)
            actual = ''
    if actual:
        siglas.append(actual)
    return any(s in SUFIJOS_EMPRESA for s in siglas)


def _separar_nombre(razon_social: str):
    """Separa un nombre plano en apellidos/nombres por conteo de palabras
    (estándar RUT: apellidos primero). Retorna None si es dudoso — vacío,
    placeholder, <2 palabras, con partícula de apellido compuesto, o con
    caracteres no alfabéticos. Ya se asume que es_empresa() dio False antes
    de llamar esto."""
    texto = (razon_social or '').strip().upper()
    if not texto or texto in _PLACEHOLDERS:
        return None

    palabras = texto.split()
    if len(palabras) < 2:
        return None
    if any(p in _PARTICULAS for p in palabras):
        return None
    if not all(_PALABRA_RE.match(p) for p in palabras):
        return None

    n = len(palabras)
    if n == 2:
        ap1, ap2, no1, otros = palabras[0], '', palabras[1], ''
    elif n == 3:
        ap1, ap2, no1, otros = palabras[0], palabras[1], palabras[2], ''
    elif n == 4:
        ap1, ap2, no1, otros = palabras[0], palabras[1], palabras[2], palabras[3]
    else:
        ap1, ap2, no1, otros = palabras[0], palabras[1], palabras[2], ' '.join(palabras[3:])
    return ap1, ap2, no1, otros


def _get_tecnopos_session():
    global _tecnopos_session
    if _tecnopos_session is None:
        s = requests.Session()
        s.headers.update({
            'Referer': TECNOPOS_REFERER,
            'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                            'AppleWebKit/537.36 (KHTML, like Gecko) '
                            'Chrome/124.0 Safari/537.36'),
        })
        _tecnopos_session = s
    return _tecnopos_session


def consultar_tecnopos(nit: str):
    """
    Intento rápido vía sipos.com.co antes de recurrir a DIAN. Devuelve el
    MISMO envelope que consultar_nit_basica() ({"status","data","error"})
    para que caché/reintentos/Excel en ui/app.py no necesiten saber de dónde
    vino el dato. Devuelve None si el resultado no es confiable (red, JSON
    inestable, nombre de persona ambiguo) — el llamador cae a DIAN.
    """
    try:
        time.sleep(random.uniform(0.05, 0.15))  # civismo con la API de un tercero
        resp = _get_tecnopos_session().get(TECNOPOS_URL, params={'nit': nit},
                                            timeout=TECNOPOS_TIMEOUT)
        data = resp.json()
    except Exception as e:
        log.debug(f"TECNOPOS no disponible para {nit}: {e}")
        return None

    if not data.get('success'):
        return None

    razon_social = str(data.get('razon_social', '')).strip()
    dv = data.get('dv', '')
    if not razon_social or dv == '':
        return None  # respuesta inestable (a veces trae direccion/ciudad/actividad en su lugar)

    salida = {
        'nit': nit,
        'dv': str(dv),
        'datetime': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'tipo_consulta': 'basica',
        'email': str(data.get('email', '') or '').strip(),
    }

    if not es_empresa(razon_social):
        # TECNOPOS NO tiene un orden de palabras consistente para personas
        # naturales — confirmado con 3 NITs reales: unos vienen "apellidos
        # nombres" (79750160, 52156616) y otros "nombres apellidos"
        # (87069568, "ALBERTO EDMUNDO SANCHEZ MARTINEZ"). Ninguna regla
        # posicional fija sirve para los dos casos a la vez, y no hay señal
        # en el JSON para distinguir cuál es cuál. El nombre de una persona
        # natural NUNCA sale de TECNOPOS, bajo ninguna circunstancia —
        # siempre de DIAN, que tiene apellidos/nombres en campos separados
        # de verdad. _separar_nombre() sigue sin usarse aquí a propósito.
        #
        # El email sí es seguro de usar (no tiene ambigüedad de orden), así
        # que se devuelve para que el coordinador lo fusione con el nombre
        # real de DIAN — email rápido + nombre confiable, sin mezclar el
        # riesgo de uno con el otro.
        return {'email': salida['email'], '_es_persona': True}

    salida['razonSocial'] = razon_social
    # sin nombre que partir — solo empresas llegan aquí, igual que hace DIAN

    return {"status": "success", "data": salida, "error": None}


def _merge_tecnopos_email_con_dian(resultado_dian: dict, resultado_tecnopos: dict) -> dict:
    """Inyecta el email de TECNOPOS en el resultado de DIAN para personas
    naturales. El nombre siempre es el de DIAN (confiable) — el email es un
    extra que se agrega solo si DIAN respondió con éxito. IMPORTANTE: el
    email va dentro de resultado_dian['data'], no en la raíz del envelope
    — ui/app.py lee data.get('email'), no resultado.get('email')."""
    email = resultado_tecnopos.get('email')
    if resultado_dian.get('status') == 'success' and email:
        resultado_dian['data']['email'] = email
    return resultado_dian


# ─────── Coordinador ───────

def consultar_nit(nit: str, tipo: str = "basica", attempt: int = 1):
    if tipo == "rut_detallado":
        return consultar_nit_rut_detallado(nit, attempt)
    if attempt == 1:
        rapido = consultar_tecnopos(nit)
        if rapido is not None:
            if rapido.get('_es_persona'):
                resultado_dian = consultar_nit_basica(nit, attempt)
                return _merge_tecnopos_email_con_dian(resultado_dian, rapido)
            return rapido
    return consultar_nit_basica(nit, attempt)
