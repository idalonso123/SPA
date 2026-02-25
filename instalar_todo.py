"""
Script maestro que:
1. Descarga y configura Python Portable
2. Crea todas las tareas programadas

Este es el script principal que debes convertir a .exe
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


def mensaje(titulo, texto):
    """Muestra un mensaje formateado"""
    print("\n" + "="*60)
    print(f"{titulo}")
    print("="*60)
    print(texto)


def verificar_python_portable():
    """Verifica si Python portable ya está configurado"""
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    if os.path.exists(python_exe):
        try:
            resultado = subprocess.run(
                [python_exe, "--version"],
                capture_output=True,
                text=True,
                check=False
            )
            if resultado.returncode == 0:
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
            return True
    except:
        pass
    return False


def descargar_python_portable():
    """Descarga Python Embeddable"""
    mensaje("DESCARGANDO PYTHON PORTABLE", "Esto puede tomar unos minutos...")
    
    url_python = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
    archivo_zip = "python-portable.zip"
    
    try:
        import urllib.request
        
        urllib.request.urlretrieve(url_python, archivo_zip)
        print("Descarga completada. Extrayendo...")
        
        ruta_proyecto = obtener_ruta_proyecto()
        
        with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
            zip_ref.extractall(ruta_proyecto)
        
        # Renombrar carpeta
        carpeta_extraccion = os.path.join(ruta_proyecto, "python-3.11.9-embed-amd64")
        carpeta_destino = os.path.join(ruta_proyecto, "python-portable")
        
        if os.path.exists(carpeta_extraccion) and not os.path.exists(carpeta_destino):
            os.rename(carpeta_extraccion, carpeta_destino)
        
        os.remove(archivo_zip)
        print("Python extraído exitosamente")
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False


def configurar_pip():
    """Configura pip para Python Embeddable"""
    print("\nConfigurando pip...")
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    try:
        import urllib.request
        url_get_pip = "https://bootstrap.pypa.io/get-pip.py"
        urllib.request.urlretrieve(url_get_pip, os.path.join(ruta_proyecto, "get-pip.py"))
        
        subprocess.run(
            [python_exe, os.path.join(ruta_proyecto, "get-pip.py")],
            capture_output=True,
            check=False
        )
        
        try:
            os.remove(os.path.join(ruta_proyecto, "get-pip.py"))
        except:
            pass
        
        print("pip configurado")
        return True
            
    except Exception as e:
        print(f"Error: {e}")
        return False


def instalar_dependencias():
    """Instala las dependencias del proyecto"""
    print("\nInstalando dependencias...")
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    requirements_file = os.path.join(ruta_proyecto, "requirements.txt")
    
    try:
        # Actualizar pip
        subprocess.run(
            [python_exe, "-m", "pip", "install", "--upgrade", "pip"],
            capture_output=True,
            check=False
        )
        
        # Instalar dependencias
        if os.path.exists(requirements_file):
            subprocess.run(
                [python_exe, "-m", "pip", "install", "-r", requirements_file],
                capture_output=True,
                check=False
            )
        
        print("Dependencias instaladas")
        return True
            
    except Exception as e:
        print(f"Error: {e}")
        return False


def obtener_python_ejecutable():
    """Obtiene la ruta al ejecutable de Python"""
    ruta_proyecto = obtener_ruta_proyecto()
    python_portable = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    if os.path.exists(python_portable):
        return python_portable
    return sys.executable


def crear_tareas_programadas():
    """Crea las tareas programadas"""
    mensaje("CREANDO TAREAS PROGRAMADAS", "Configurando tareas en Windows...")
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = obtener_python_ejecutable()
    
    # Importar y ejecutar la lógica de crear_tareas_programadas
    sys.path.insert(0, ruta_proyecto)
    
    # Ejecutar las tareas una por una
    tareas = [
        ("Vivero_Main_Pedidos_Semanales", "main.py", "", "THU", 21, 0),
        ("Vivero_Informe_Compras_Sin_Autorizacion", "informe_compras_sin_autorizacion.py", "", "THU", 21, 10),
        ("Vivero_Informe_Articulos_No_Comprados", "Informe_artículos_no_comprados.py", "", "THU", 21, 10),
        ("Vivero_Analisis_Categoria_CD", "analisis_categoria_cd.py", "", "THU", 21, 10),
    ]
    
    # Crear tarea semanal
    for nombre_tarea, nombre_script, parametros, dia_semana, hora, minuto in tareas:
        print(f"\nCreando: {nombre_tarea}")
        
        if parametros:
            comando = f'cmd /c "cd /d \\"{ruta_proyecto}\\" && \\"{python_exe}\\" \\"{nombre_script}\\" {parametros}"'
        else:
            comando = f'cmd /c "cd /d \\"{ruta_proyecto}\\" && \\"{python_exe}\\" \\"{nombre_script}\\""'
        
        schtasks = [
            "schtasks", "/create", "/tn", nombre_tarea, "/tr", comando,
            "/sc", "weekly", "/d", dia_semana, "/st", f"{hora}:{str(minuto).zfill(2)}", "/f"
        ]
        
        subprocess.run(schtasks, capture_output=True, check=False)
        print(f"  ✓ Creada")
    
    # Tareas mensuales
    tareas_mensuales = [
        ("Vivero_ClasificacionABC_P1", "clasificacionABC.py", "--P1", 1, 1, 12, 0),
        ("Vivero_ClasificacionABC_P2", "clasificacionABC.py", "--P2", 1, 2, 9, 0),
        ("Vivero_ClasificacionABC_P3", "clasificacionABC.py", "--P3", 1, 5, 9, 0),
        ("Vivero_ClasificacionABC_P4", "clasificacionABC.py", "--P4", 1, 8, 9, 0),
        ("Vivero_Presentacion_Enero", "PRESENTACION.py", "", 1, 1, 12, 30),
        ("Vivero_Presentacion_Febrero", "PRESENTACION.py", "", 1, 2, 9, 30),
        ("Vivero_Presentacion_Mayo", "PRESENTACION.py", "", 1, 5, 9, 30),
        ("Vivero_Presentacion_Agosto", "PRESENTACION.py", "", 1, 8, 9, 30),
        ("Vivero_Informe_Enero", "INFORME.py", "", 1, 1, 12, 30),
        ("Vivero_Informe_Febrero", "INFORME.py", "", 1, 2, 9, 30),
        ("Vivero_Informe_Mayo", "INFORME.py", "", 1, 5, 9, 30),
        ("Vivero_Informe_Agosto", "INFORME.py", "", 1, 8, 9, 30),
    ]
    
    for nombre_tarea, nombre_script, parametros, dia, mes, hora, minuto in tareas_mensuales:
        print(f"\nCreando: {nombre_tarea}")
        
        if parametros:
            comando = f'cmd /c "cd /d \\"{ruta_proyecto}\\" && \\"{python_exe}\\" \\"{nombre_script}\\" {parametros}"'
        else:
            comando = f'cmd /c "cd /d \\"{ruta_proyecto}\\" && \\"{python_exe}\\" \\"{nombre_script}\\""'
        
        schtasks = [
            "schtasks", "/create", "/tn", nombre_tarea, "/tr", comando,
            "/sc", "monthly", "/d", str(dia), "/m", str(mes).zfill(2),
            "/st", f"{hora}:{str(minuto).zfill(2)}", "/f"
        ]
        
        subprocess.run(schtasks, capture_output=True, check=False)
        print(f"  ✓ Creada")
    
    print("\n¡Todas las tareas creadas!")


def main():
    """Función principal"""
    print("\n" + "="*60)
    print("CONFIGURADOR COMPLETO - VIVERO ARANJUEZ")
    print("="*60)
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Proyecto: {obtener_ruta_proyecto()}")
    
    # Paso 1: Configurar Python portable
    if not verificar_python_portable():
        print("\nPython portable no encontrado. Descargando...")
        if not descargar_python_portable():
            print("Error al descargar Python")
            return
    else:
        print("\n✓ Python portable ya está")
    
    # Paso 2: Configurar pip
    if not verificar_pip():
        print("\n pip no encontrado. Configurando...")
        configurar_pip()
    else:
        print("✓ pip ya está")
    
    # Paso 3: Instalar dependencias
    print("\nVerificando dependencias...")
    instalar_dependencias()
    
    # Paso 4: Crear tareas
    crear_tareas_programadas()
    
    # Resumen
    print("\n" + "="*60)
    print("¡CONFIGURACIÓN COMPLETADA!")
    print("="*60)
    print("\nSe han creado las siguientes tareas:")
    print("\n  SEMANALES (jueves):")
    print("    - Vivero_Main_Pedidos_Semanales (21:00)")
    print("    - Vivero_Informe_Compras_Sin_Autorizacion (21:10)")
    print("    - Vivero_Informe_Articulos_No_Comprados (21:10)")
    print("    - Vivero_Analisis_Categoria_CD (21:10)")
    print("\n  MENSUALES:")
    print("    - Clasificación ABC (P1, P2, P3, P4)")
    print("    - Presentación (Enero, Febrero, Mayo, Agosto)")
    print("    - Informe (Enero, Febrero, Mayo, Agosto)")
    print("\n" + "="*60)
    
    input("\nPresiona Enter para salir...")


if __name__ == "__main__":
    main()
