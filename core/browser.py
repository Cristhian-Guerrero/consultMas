"""
Gestión del pool de navegadores Chromium.
"""

import platform
from DrissionPage import ChromiumPage, ChromiumOptions
from config import MAX_POOL_SIZE, BROWSER_CONFIGS

SISTEMA_OPERATIVO = platform.system()

# pygetwindow solo en Windows
if SISTEMA_OPERATIVO == "Windows":
    try:
        import pygetwindow as gw
        _PYGETWINDOW = True
    except ImportError:
        _PYGETWINDOW = False
        gw = None
else:
    _PYGETWINDOW = False
    gw = None

_BROWSER_POOL = []


def crear_navegador(config_index=0):
    config = BROWSER_CONFIGS[config_index % len(BROWSER_CONFIGS)]
    options = ChromiumOptions().auto_port()
    options.set_argument("--disable-dev-shm-usage")
    options.set_argument("--window-size=400,300")
    options.set_argument("--window-position=-2000,-2000")
    options.set_argument("--disable-extensions")
    options.set_argument("--disable-plugins")
    options.set_argument("--no-first-run")
    options.set_argument("--disable-default-apps")
    return ChromiumPage(addr_or_opts=options), config['name']


def limpiar_inactivos():
    global _BROWSER_POOL
    activos = []
    for driver in _BROWSER_POOL:
        try:
            driver.title
            activos.append(driver)
        except:
            try:
                driver.quit()
            except:
                pass
    _BROWSER_POOL = activos
    return len(_BROWSER_POOL)


def get_browser():
    limpiar_inactivos()
    if _BROWSER_POOL:
        return _BROWSER_POOL.pop()
    return crear_navegador(0)[0]


def return_browser(driver):
    limpiar_inactivos()
    if driver and len(_BROWSER_POOL) < MAX_POOL_SIZE:
        try:
            driver.get('about:blank')
            _BROWSER_POOL.append(driver)
        except:
            try:
                driver.quit()
            except:
                pass


def cerrar_todos():
    global _BROWSER_POOL
    for driver in _BROWSER_POOL:
        try:
            driver.quit()
        except Exception as e:
            print(f"Error cerrando navegador: {e}")
    _BROWSER_POOL = []


def minimizar_chromium():
    if not _PYGETWINDOW:
        return
    try:
        for window in gw.getWindowsWithTitle("Chromium"):
            window.minimize()
    except Exception as e:
        print(f"⚠️ No se pudieron minimizar ventanas: {e}")
