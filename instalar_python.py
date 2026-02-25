"""
Script para instalar Python automáticamente
Designed to work on Windows systems
"""

import subprocess
import sys
import os
from datetime import datetime


def verificar_python_instalado():
    """Verifica si Python ya está instalado"""
    try:
        resultado = subprocess.run(
            ["python", "--version"],
            capture_output=True,
            text=True,
            check=False
        )
        if resultado.returncode == 0:
            print(f"Python ya está instalado: {resultado.stdout.strip()}")
            return True
    except FileNotFoundError:
        pass
    
    # Try python3
    try:
        resultado = subprocess.run(
            ["python3", "--version"],
            capture_output=True,
            text=True,
            check=False
        )
        if resultado.returncode == 0:
            print(f"Python ya está instalado: {resultado.stdout.strip()}")
            return True
    except FileNotFoundError:
        pass
    
    print("Python no está instalado o no está en el PATH")
    return False


def verificar_winget():
    """Verifica si winget está disponible"""
    try:
        resultado = subprocess.run(
            ["winget", "--version"],
            capture_output=True,
            text=True,
            check=False
        )
        if resultado.returncode == 0:
            print(f"winget disponible: {resultado.stdout.strip()}")
            return True
    except FileNotFoundError:
        pass
    
    print("winget no está disponible")
    return False


def instalar_con_winget():
    """Instala Python usando winget (Windows Package Manager)"""
    print("\nIntentando instalar Python con winget...")
    print("Esto requiere permisos de administrador")
    print("Por favor, acepta la ventana de consentimiento que aparecerá\n")
    
    try:
        # Install Python using winget
        resultado = subprocess.run(
            ["winget", "install", "--id", "Python.Python.3.11", "--silent", "--accept-package-agreements", "--accept-source-agreements"],
            capture_output=True,
            text=True,
            check=False
        )
        
        if resultado.returncode == 0:
            print("Python instalado exitosamente con winget")
            print("\nNOTA: Es posible que necesites reiniciar tu PC o")
            print("cerrar y abrir la terminal para que Python esté disponible en el PATH")
            return True
        else:
            print(f"Error con winget: {resultado.stderr}")
            return False
            
    except Exception as e:
        print(f"Excepción: {e}")
        return False


def instalar_descargando():
    """Descarga e instala Python manualmente"""
    print("\nDescargando Python...")
    print("Esto puede tomar unos minutos depending de tu conexión...\n")
    
    # URL del instalador de Python 3.11 (versión estable)
    url_instalador = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    nombre_archivo = "python_installer.exe"
    
    try:
        import urllib.request
        
        # Descargar el instalador
        print("Descargando instalador de Python...")
        urllib.request.urlretrieve(url_instalador, nombre_archivo)
        print("Descarga completada")
        
        # Ejecutar el instalador silenciosamente
        print("\nEjecutando instalador...")
        print("Instalando Python (esto puede tardar unos minutos)...\n")
        
        # Opciones de instalación silenciosa
        # /quiet = instalación silenciosa
        # InstallAllUsers=1 = instalar para todos los usuarios
        # PrependPath=1 = agregar al PATH
        # Include_test=0 = no incluir Python test suite
        comando = [
            nombre_archivo,
            "/quiet",
            "InstallAllUsers=1",
            "PrependPath=1",
            "Include_test=0"
        ]
        
        subprocess.run(comando, check=True)
        
        print("Python instalado exitosamente")
        print("\nNOTA: Es posible que necesites reiniciar tu PC para que")
        print("Python esté disponible en el PATH")
        
        # Limpiar archivo descargado
        try:
            os.remove(nombre_archivo)
        except:
            pass
            
        return True
        
    except Exception as e:
        print(f"Error al descargar/instalar: {e}")
        return False


def main():
    """Función principal"""
    print("\n" + "="*60)
    print("INSTALADOR DE PYTHON")
    print("="*60)
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Verificar si ya está instalado
    if verificar_python_instalado():
        print("\n¡Python ya está instalado! No necesitas hacer nada más.")
        input("\nPresiona Enter para salir...")
        return
    
    print("\nPython no está instalado. Procediendo con la instalación...\n")
    
    # Intentar con winget primero
    if verificar_winget():
        print("\n" + "="*60)
        opcion = input("¿Quieres instalar Python ahora? (s/n): ").strip().lower()
        
        if opcion == "s" or opcion == "si":
            if instalar_con_winget():
                print("\n" + "="*60)
                print("INSTALACIÓN COMPLETADA")
                print("="*60)
            else:
                print("\n winget falló. Intentando método alternativo...")
                instalar_descargando()
        else:
            print("Instalación cancelada")
    else:
        # Si no hay winget, descargar directamente
        print("\nNo se encontró winget. Se descargará Python automáticamente...")
        opcion = input("¿Continuar con la descarga? (s/n): ").strip().lower()
        
        if opcion == "s" or opcion == "si":
            instalar_descargando()
        else:
            print("Instalación cancelada")
    
    print("\n" + "="*60)
    print("INSTRUCCIONES POST-INSTALACIÓN")
    print("="*60)
    print("1. Reinicia tu PC (recomendado)")
    print("2. O cierra la terminal y abre una nueva")
    print("3. Verifica la instalación ejecutando: python --version")
    print("\nSi después de reiniciar Python no funciona, consulta el")
    print("manual de instalación de Python para agregar manualmente")
    print("al PATH (variables de entorno)")
    
    input("\nPresiona Enter para salir...")


if __name__ == "__main__":
    main()
