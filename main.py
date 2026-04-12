"""
Punto de entrada — Consulta Gestión Masiva DIAN
A.S. Contadores & Asesores SAS
"""

import logging
from ui.app import ConsultaRUTApp
from core.browser import cerrar_todos

logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def main():
    app = ConsultaRUTApp()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.getLogger(__name__).critical(f"Error crítico: {e}", exc_info=True)
    finally:
        cerrar_todos()
