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
    # Si se ejecuta como exe, obtener la carpeta donde está el ejecutable
    if getattr(sys, 'frozen', False):
        # Estamos en un exe compilado
        ruta_exe = os.path.dirname(sys.executable)
        return ruta_exe
    else:
        # Estamos en modo script normal
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
    python_exe = os.path.join(ruta_proyecto, "Python", "python.exe")
    
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
    python_exe = os.path.join(ruta_proyecto, "Python", "python.exe")
    
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
    """Descarga Python completo (no embeddable) para uso portable"""
    mensaje("DESCARGANDO PYTHON PORTABLE", "Esto puede tomar varios minutos...")
    
    # Usar el instalador de Python en modo silencioso
    # InstallDir crea una instalación en la carpeta especificada sin afectar el sistema
    url_python = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    archivo_instalador = "python-installer.exe"
    ruta_proyecto = obtener_ruta_proyecto()
    carpeta_python = os.path.join(ruta_proyecto, "Python")
    
    try:
        import urllib.request
        
        print("Descargando instalador de Python 3.11.9...")
        urllib.request.urlretrieve(url_python, archivo_instalador)
        print("Instalando Python en modo silencioso...")
        
        # Instalar Python silenciosamente en la carpeta especificada
        # /quiet = modo silencioso
        # InstallDir=ruta = carpeta de instalación
        # TargetDir=ruta = carpeta de instalación (alternativo)
        # PrependPath=0 = no añadir al PATH del sistema
        # SimpleInstall=1 = instalación simple
        # SimpleInstallDomain=0 = no crear entorno de dominio
        install_args = [
            archivo_instalador,
            "/quiet",
            "InstallAllUsers=0",
            "PrependPath=0",
            f"TargetDir={carpeta_python}",
            "AssociateFiles=0",
            "Shortcuts=0",
            "Include_test=0"
        ]
        
        resultado = subprocess.run(
            install_args,
            capture_output=True,
            text=True,
            timeout=600  # 10 minutos para instalar
        )
        
        # Limpiar instalador
        try:
            os.remove(archivo_instalador)
        except:
            pass
        
        # Verificar instalación
        python_exe = os.path.join(carpeta_python, "python.exe")
        if os.path.exists(python_exe):
            print("Python instalado correctamente")
            return True
        else:
            print("Advertencia: Python puede no haberse instalado correctamente")
            # Intentar con método alternativo
            return instalar_python_alternativo(ruta_proyecto)
        
    except subprocess.TimeoutExpired:
        print("Error: Timeout al instalar Python")
        return instalar_python_alternativo(ruta_proyecto)
    except Exception as e:
        print(f"Error al descargar Python: {e}")
        return instalar_python_alternativo(ruta_proyecto)


def instalar_python_alternativo(ruta_proyecto):
    """Método alternativo: descargar Python Embeddable y configurar manualmente"""
    print("\nUsando método alternativo...")
    
    try:
        import urllib.request
        
        # Descargar Python Embeddable
        url_python = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
        archivo_zip = "python-portable.zip"
        
        urllib.request.urlretrieve(url_python, archivo_zip)
        print("Extrayendo Python...")
        
        with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
            zip_ref.extractall(ruta_proyecto)
        
        carpeta_extraccion = os.path.join(ruta_proyecto, "python-3.11.9-embed-amd64")
        carpeta_destino = os.path.join(ruta_proyecto, "Python")
        
        if os.path.exists(carpeta_extraccion) and not os.path.exists(carpeta_destino):
            os.rename(carpeta_extraccion, carpeta_destino)
        
        # Añadir pip manualmente descargando los archivos necesarios
        print("Configurando pip manualmente...")
        if not configurar_pip_manual(carpeta_destino):
            print("Advertencia: pip no se pudo configurar completamente")
        
        os.remove(archivo_zip)
        print("Python extraído exitosamente")
        return True
        
    except Exception as e:
        print(f"Error en método alternativo: {e}")
        return False


def configurar_pip_manual(carpeta_python):
    """Configura pip manualmente para Python Embeddable"""
    try:
        # Descargar pip y wheel
        import urllib.request
        
        # Crear carpeta Scripts si no existe
        scripts_folder = os.path.join(carpeta_python, "Scripts")
        if not os.path.exists(scripts_folder):
            os.makedirs(scripts_folder)
        
        # Descargar get-pip.py
        url_get_pip = "https://bootstrap.pypa.io/get-pip.py"
        get_pip_path = os.path.join(carpeta_python, "get-pip.py")
        urllib.request.urlretrieve(url_get_pip, get_pip_path)
        
        # Ejecutar get-pip.py
        python_exe = os.path.join(carpeta_python, "python.exe")
        resultado = subprocess.run(
            [python_exe, get-pip_path, "--no-warn-script-location"],
            capture_output=True,
            text=True,
            timeout=300,
            cwd=carpeta_python
        )
        
        # Limpiar
        try:
            os.remove(get_pip_path)
        except:
            pass
        
        # Verificar si pip está disponible
        resultado_pip = subprocess.run(
            [python_exe, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        return resultado_pip.returncode == 0
        
    except Exception as e:
        print(f"Error configurando pip manualmente: {e}")
        return False


def configurar_pip():
    """Configura pip para Python Embeddable"""
    print("\nConfigurando pip...")
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "Python", "python.exe")
    
    try:
        import urllib.request
        
        # Descargar get-pip.py
        url_get_pip = "https://bootstrap.pypa.io/get-pip.py"
        get_pip_path = os.path.join(ruta_proyecto, "get-pip.py")
        urllib.request.urlretrieve(url_get_pip, get_pip_path)
        
        # Ejecutar get-pip.py con timeout más largo
        resultado = subprocess.run(
            [python_exe, get_pip_path],
            capture_output=True,
            text=True,
            timeout=300,  # 5 minutos de timeout
            cwd=ruta_proyecto
        )
        
        # Limpiar archivo
        try:
            os.remove(get_pip_path)
        except:
            pass
        
        # Verificar si pip se instaló
        resultado_pip = subprocess.run(
            [python_exe, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if resultado_pip.returncode == 0:
            print("pip configurado correctamente")
            return True
        else:
            print(f"Advertencia: pip puede no estar configurado correctamente")
            print(f"Salida: {resultado.stdout}")
            return True  # Continuar de todos modos
            
    except subprocess.TimeoutExpired:
        print("Error: Timeout al configurar pip")
        return False
    except Exception as e:
        print(f"Error al configurar pip: {e}")
        return False


def instalar_dependencias():
    """Instala las dependencias del proyecto"""
    print("\nInstalando dependencias...")
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = os.path.join(ruta_proyecto, "Python", "python.exe")
    requirements_file = os.path.join(ruta_proyecto, "requirements.txt")
    
    try:
        # Verificar si pip está disponible
        resultado_pip = subprocess.run(
            [python_exe, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if resultado_pip.returncode != 0:
            print("Advertencia: pip no está disponible, saltando instalación de dependencias")
            print("Las tareas programadas intentarán usar el Python del sistema si falla")
            return True  # Continuar de todos modos
        
        # Actualizar pip con timeout
        print("Actualizando pip...")
        subprocess.run(
            [python_exe, "-m", "pip", "install", "--upgrade", "pip"],
            capture_output=True,
            timeout=300,
            check=False
        )
        
        # Instalar dependencias
        if os.path.exists(requirements_file):
            print("Instalando dependencias del proyecto...")
            resultado = subprocess.run(
                [python_exe, "-m", "pip", "install", "-r", requirements_file],
                capture_output=True,
                text=True,
                timeout=600,  # 10 minutos para instalar dependencias
                check=False
            )
            if resultado.returncode == 0:
                print("Dependencias instaladas correctamente")
            else:
                print(f"Advertencia: algunas dependencias pueden no haberse instalado")
                print(f"Detalles: {resultado.stdout[:200]}")
        else:
            print(f"Advertencia: archivo requirements.txt no encontrado")
        
        return True
            
    except subprocess.TimeoutExpired:
        print("Error: Timeout al instalar dependencias")
        return False
    except Exception as e:
        print(f"Error al instalar dependencias: {e}")
        return False


def obtener_python_ejecutable():
    """Obtiene la ruta al ejecutable de Python"""
    ruta_proyecto = obtener_ruta_proyecto()
    python_portable = os.path.join(ruta_proyecto, "Python", "python.exe")
    
    if os.path.exists(python_portable):
        return python_portable
    # Si no existe, intentar con python del sistema
    return "python"


def crear_tareas_programadas():
    """Crea las tareas programadas"""
    mensaje("CREANDO TAREAS PROGRAMADAS", "Configurando tareas en Windows...")
    
    ruta_proyecto = obtener_ruta_proyecto()
    python_exe = obtener_python_ejecutable()
    
    # Mostrar la ruta de Python que se usará
    print(f"\nPython ejecutable: {python_exe}")
    print(f"Ruta del proyecto: {ruta_proyecto}")
    
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
    print("\n--- TAREAS SEMANALES ---")
    for nombre_tarea, nombre_script, parametros, dia_semana, hora, minuto in tareas:
        print(f"\nCreando: {nombre_tarea}")
        
        if parametros:
            comando = f'cmd /c "cd /d "{ruta_proyecto}" && "{python_exe}" "{nombre_script}" {parametros}"'
        else:
            comando = f'cmd /c "cd /d "{ruta_proyecto}" && "{python_exe}" "{nombre_script}""'
        
        schtasks = [
            "schtasks", "/create", "/tn", nombre_tarea, "/tr", comando,
            "/sc", "weekly", "/d", dia_semana, "/st", f"{hora}:{str(minuto).zfill(2)}", "/f"
        ]
        
        resultado = subprocess.run(schtasks, capture_output=True, text=True)
        if resultado.returncode == 0:
            print(f"  ✓ Creada")
        else:
            print(f"  ✗ Error: {resultado.stderr[:100]}")
    
    # Tareas mensuales - Formato correcto para Windows
    # /d = día del mes (1-31), /m = mes (JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC)
    tareas_mensuales = [
        ("Vivero_ClasificacionABC_P1", "clasificacionABC.py", "--P1", "1", "JAN", 12, 0),
        ("Vivero_ClasificacionABC_P2", "clasificacionABC.py", "--P2", "1", "FEB", 9, 0),
        ("Vivero_ClasificacionABC_P3", "clasificacionABC.py", "--P3", "1", "MAY", 9, 0),
        ("Vivero_ClasificacionABC_P4", "clasificacionABC.py", "--P4", "1", "AUG", 9, 0),
        ("Vivero_Presentacion_Enero", "PRESENTACION.py", "", "1", "JAN", 12, 30),
        ("Vivero_Presentacion_Febrero", "PRESENTACION.py", "", "1", "FEB", 9, 30),
        ("Vivero_Presentacion_Mayo", "PRESENTACION.py", "", "1", "MAY", 9, 30),
        ("Vivero_Presentacion_Agosto", "PRESENTACION.py", "", "1", "AUG", 9, 30),
        ("Vivero_Informe_Enero", "INFORME.py", "", "1", "JAN", 12, 30),
        ("Vivero_Informe_Febrero", "INFORME.py", "", "1", "FEB", 9, 30),
        ("Vivero_Informe_Mayo", "INFORME.py", "", "1", "MAY", 9, 30),
        ("Vivero_Informe_Agosto", "INFORME.py", "", "1", "AUG", 9, 30),
    ]
    
    print("\n--- TAREAS MENSUALES ---")
    for nombre_tarea, nombre_script, parametros, dia, mes, hora, minuto in tareas_mensuales:
        print(f"\nCreando: {nombre_tarea}")
        
        if parametros:
            comando = f'cmd /c "cd /d "{ruta_proyecto}" && "{python_exe}" "{nombre_script}" {parametros}"'
        else:
            comando = f'cmd /c "cd /d "{ruta_proyecto}" && "{python_exe}" "{nombre_script}""'
        
        # Usar /d para el día y /m para el mes específico
        # Formatear hora con dos dígitos (HH:MM)
        hora_formateada = f"{str(hora).zfill(2)}:{str(minuto).zfill(2)}"
        schtasks = [
            "schtasks", "/create", "/tn", nombre_tarea, "/tr", comando,
            "/sc", "monthly", "/d", dia, "/m", mes,
            "/st", hora_formateada, "/f"
        ]
        
        resultado = subprocess.run(schtasks, capture_output=True, text=True)
        if resultado.returncode == 0:
            print(f"  ✓ Creada")
        else:
            print(f"  ✗ Error: {resultado.stderr[:150]}")
    
    print("\n¡Todas las tareas creadas!")


def main():
    """Función principal"""
    print("\n" + "="*60)
    print("CONFIGURADOR COMPLETO - VIVEVERDE")
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
