"""
Script para descargar y configurar Python Portable (Embeddable)
Incluye pip y las dependencias necesarias
"""

import subprocess
import sys
import os
import zipfile
import shutil
from datetime import datetime


def obtener_ruta_proyecto():
    """Obtiene la ruta del directorio del proyecto"""
    return os.path.dirname(os.path.abspath(__file__))


def verificar_python_portable():
    """Verifica si Python portable ya está configurado"""
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    if os.path.exists(python_exe):
        # Verificar que funciona
        try:
            resultado = subprocess.run(
                [python_exe, "--version"],
                capture_output=True,
                text=True,
                check=False
            )
            if resultado.returncode == 0:
                print(f"Python portable ya está configurado: {resultado.stdout.strip()}")
                return True
        except:
            pass
    
    return False


def verificar_pip():
    """Verifica si pip está disponible"""
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    try:
        resultado = subprocess.run(
            [python_exe, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            check=False
        )
        if resultado.returncode == 0:
            print(f"pip ya está disponible: {resultado.stdout.strip()}")
            return True
    except:
        pass
    
    return False


def descargar_python_portable():
    """Descarga Python Embeddable"""
    print("\n" + "="*60)
    print("DESCARGANDO PYTHON PORTABLE")
    print("="*60)
    
    # URL de Python 3.11.9 Embeddable (versión estable)
    # Usamos la versión embeddable que es más ligera
    url_python = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
    archivo_zip = "python-portable.zip"
    
    try:
        import urllib.request
        
        print(f"Descargando desde: {url_python}")
        print("Esto puede tomar unos minutos...\n")
        
        # Descargar el archivo
        urllib.request.urlretrieve(url_python, archivo_zip)
        print("Descarga completada")
        
        # Extraer
        print("Extrayendo archivos...")
        ruta_proyecto = obtener_ruta_proyecto()
        carpeta_destino = os.path.join(ruta_proyecto, "python-portable")
        
        with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
            zip_ref.extractall(ruta_proyecto)
        
        # Renombrar carpeta si es necesario
        carpeta_extraccion = os.path.join(ruta_proyecto, "python-3.11.9-embed-amd64")
        if os.path.exists(carpeta_extraccion) and not os.path.exists(carpeta_destino):
            os.rename(carpeta_extraccion, carpeta_destino)
        
        # Limpiar archivo zip
        os.remove(archivo_zip)
        
        print(f"Python extraído en: {carpeta_destino}")
        return True
        
    except Exception as e:
        print(f"Error al descargar: {e}")
        return False


def configurar_pip():
    """Configura pip para Python Embeddable"""
    print("\n" + "="*60)
    print("CONFIGURANDO PIP")
    print("="*60)
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    # Python embeddable no incluye pip por defecto
    # Necesitamos descargarlo
    
    print("Descargando pip...")
    
    try:
        # Descargar get-pip.py
        import urllib.request
        url_get_pip = "https://bootstrap.pypa.io/get-pip.py"
        urllib.request.urlretrieve(url_get_pip, os.path.join(ruta_proyecto, "get-pip.py"))
        
        # Instalar pip
        print("Instalando pip...")
        resultado = subprocess.run(
            [python_exe, os.path.join(ruta_proyecto, "get-pip.py")],
            capture_output=True,
            text=True,
            check=False
        )
        
        # Limpiar
        try:
            os.remove(os.path.join(ruta_proyecto, "get-pip.py"))
        except:
            pass
        
        if resultado.returncode == 0:
            print("pip instalado exitosamente")
            return True
        else:
            print(f"Error al instalar pip: {resultado.stderr}")
            return False
            
    except Exception as e:
        print(f"Error: {e}")
        return False


def instalar_dependencias():
    """Instala las dependencias del proyecto"""
    print("\n" + "="*60)
    print("INSTALANDO DEPENDENCIAS")
    print("="*60)
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    # Leer requirements.txt si existe
    requirements_file = os.path.join(ruta_proyecto, "requirements.txt")
    
    if not os.path.exists(requirements_file):
        print("No se encontró requirements.txt")
        print("Instalando dependencias básicas de todos modos...")
    
    try:
        # Actualizar pip primero
        print("Actualizando pip...")
        subprocess.run(
            [python_exe, "-m", "pip", "install", "--upgrade", "pip"],
            capture_output=True,
            check=False
        )
        
        # Instalar dependencias
        if os.path.exists(requirements_file):
            print(f"Instalando dependencias desde {requirements_file}...")
            resultado = subprocess.run(
                [python_exe, "-m", "pip", "install", "-r", requirements_file],
                capture_output=True,
                text=True,
                check=False
            )
            
            if resultado.returncode == 0:
                print("Dependencias instaladas exitosamente")
                return True
            else:
                print(f"Advertencia: {resultado.stderr}")
                return True  # No fallar por dependencias opcionales
        else:
            print("No hay requirements.txt,saltando...")
            return True
            
    except Exception as e:
        print(f"Error al instalar dependencias: {e}")
        return False


def main():
    """Función principal"""
    print("\n" + "="*60)
    print("CONFIGURADOR DE PYTHON PORTABLE")
    print("="*60)
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Proyecto: {obtener_ruta_proyecto()}")
    
    # Verificar si ya está configurado
    if verificar_python_portable() and verificar_pip():
        print("\n¡Python portable ya está completamente configurado!")
        print("No necesitas hacer nada más.")
    else:
        print("\nConfigurando Python portable...\n")
        
        # Descargar Python
        if not verificar_python_portable():
            if not descargar_python_portable():
                print("Error al descargar Python")
                return
        else:
            print("Python ya descargado")
        
        # Configurar pip
        if not verificar_pip():
            if not configurar_pip():
                print("Error al configurar pip")
                return
        else:
            print("pip ya configurado")
        
        # Instalar dependencias
        instalar_dependencias()
        
        print("\n" + "="*60)
        print("¡CONFIGURACIÓN COMPLETADA!")
        print("="*60)
    
    print("\nAhora puedes:")
    print("1. Ejecutar: python-portable\\python.exe crear_tareas_programadas.py")
    print("2. O ejecutar: crear_tareas_programadas.py (si Python está en PATH)")
    print("\nLas tareas programadas usarán:")
    print("   python-portable\\python.exe main.py")
    print("   python-portable\\python.exe clasificacionABC.py")
    print("   etc.")
    
    input("\nPresiona Enter para salir...")


if __name__ == "__main__":
    main()
