"""
Script para generar el archivo .exe de crear_tareas_programadas.py
Ejecuta este script para generar el ejecutable
"""

import subprocess
import sys
import os
from datetime import datetime


def verificar_pyinstaller():
    """Verifica si PyInstaller está instalado"""
    try:
        import PyInstaller
        print(f"PyInstaller ya está instalado (versión {PyInstaller.__version__})")
        return True
    except ImportError:
        print("PyInstaller no está instalado.")
        return False


def instalar_pyinstaller():
    """Instala PyInstaller"""
    print("\nInstalando PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller instalado exitosamente")
        return True
    except Exception as e:
        print(f"Error al instalar PyInstaller: {e}")
        return False


def crear_ejecutable():
    """Crea el archivo .exe usando PyInstaller"""
    print("\n" + "="*60)
    print("GENERANDO ARCHIVO .EXE")
    print("="*60)
    
    # Verificar PyInstaller
    if not verificar_pyinstaller():
        if not instalar_pyinstaller():
            print("No se pudo instalar PyInstaller. Saliendo.")
            return False
    
    # Opciones de PyInstaller
    opciones = [
        "--onefile",           # Crear un solo archivo .exe
        "--noconsole",         # No mostrar ventana de consola
        "--noconfirm",         # No preguntar si sobrescribir
        "--name", "Crear_Tareas_Programadas",  # Nombre del ejecutable
        "crear_tareas_programadas.py"  # Script a convertir
    ]
    
    print(f"\nEjecutando PyInstaller con opciones: {' '.join(opciones)}")
    print("Esto puede tomar unos minutos...\n")
    
    try:
        resultado = subprocess.run(
            [sys.executable, "-m", "PyInstaller"] + opciones,
            capture_output=True,
            text=True,
            check=False
        )
        
        if resultado.returncode == 0:
            print("\n" + "="*60)
            print("¡ÉXITO! Archivo .exe creado")
            print("="*60)
            
            # Buscar el archivo .exe generado
            ruta_exe = os.path.join("dist", "Crear_Tareas_Programadas.exe")
            
            if os.path.exists(ruta_exe):
                print(f"\nArchivo generado: {os.path.abspath(ruta_exe)}")
                
                # Mover a la raíz del proyecto
                destino = os.path.join(os.path.dirname(__file__), "Crear_Tareas_Programadas.exe")
                try:
                    import shutil
                    shutil.move(ruta_exe, destino)
                    print(f"Movido a: {destino}")
                except:
                    pass
                
                print("\nPuedes ejecutar este archivo .exe en cualquier PC con Windows")
                print("y creará las tareas programadas en esa máquina.")
                return True
            else:
                print("No se encontró el archivo .exe generado")
                return False
        else:
            print(f"\nError al crear el ejecutable:")
            print(resultado.stderr)
            return False
            
    except Exception as e:
        print(f"Excepción: {e}")
        return False


def main():
    """Función principal"""
    print("\n" + "="*60)
    print("GENERADOR DE .EXE PARA TAREAS PROGRAMADAS")
    print("="*60)
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Python: {sys.executable}")
    
    crear_ejecutable()
    
    input("\nPresiona Enter para salir...")


if __name__ == "__main__":
    main()
