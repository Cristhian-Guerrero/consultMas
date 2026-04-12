"""
Lógica de scraping DIAN — Consulta Express y RUT Detallado.
"""

import time
from datetime import datetime
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
        print(f"Error set_field_js {element_id}: {e}")
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
        print(f"Error click_js {element_id}: {e}")
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
        print(f"Error calculando DV para NIT {nit}: {e}")
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
            print(f"Warning Cloudflare: {e}")

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
        print(f"❌ Error consulta Express {nit}: {e}")
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
        print(f"Error set_nit_rut: {e}")
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
        print(f"❌ Error RUT {nit}: {e}")
        return {"status": "error", "data": {}, "error": str(e)}
    finally:
        if driver:
            return_browser(driver)


# ─────── Coordinador ───────

def consultar_nit(nit: str, tipo: str = "basica", attempt: int = 1):
    if tipo == "rut_detallado":
        return consultar_nit_rut_detallado(nit, attempt)
    return consultar_nit_basica(nit, attempt)
