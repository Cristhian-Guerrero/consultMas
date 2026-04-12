"""
Punto de entrada — Consulta Gestión Masiva DIAN
A.S. Contadores & Asesores SAS
"""

from ui.app import ConsultaRUTApp
from core.browser import cerrar_todos


def main():
    print("🚀 === CONSULTA DIAN - A.S. CONTADORES & ASESORES SAS ===")
    app = ConsultaRUTApp()
    print("✅ Sistema iniciado")
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Error crítico: {e}")
    finally:
        print("🧹 Cerrando sistema...")
        cerrar_todos()
        print("✅ Sistema cerrado correctamente")
