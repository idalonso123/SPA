"""
Script para crear tareas programadas en Windows
Ejecuta main.py, clasificacionABC.py, PRESENTACION.py e INFORME.py en fechas específicas
"""

import subprocess
import os
import sys
from datetime import datetime


def obtener_ruta_proyecto():
    """Obtiene la ruta del directorio del proyecto"""
    directorio_script = os.path.dirname(os.path.abspath(__file__))
    return directorio_script


def obtener_python_ejecutable():
    """
    Obtiene la ruta al ejecutable de Python
    Prioriza Python portable si está disponible
    """
    ruta_proyecto = obtener_ruta_proyecto()
    
    # Verificar si existe Python portable
    python_portable = os.path.join(ruta_proyecto, "python-portable", "python.exe")
    
    if os.path.exists(python_portable):
        print(f"  Detectado Python portable: {python_portable}")
        return python_portable
    
    # Si no existe, usar el Python del sistema
    print(f"  Usando Python del sistema: {sys.executable}")
    return sys.executable


def crear_tarea_programada(nombre_tarea, nombre_script, parametros, dia, mes, hora, minuto):
    """
    Crea una tarea programada mensual en Windows usando schtasks
    Usa rutas relativas para que sea portable entre diferentes PC
    
    Args:
        nombre_tarea: Nombre identificador de la tarea
        nombre_script: Nombre del script a ejecutar
        parametros: Parámetros adicionales para el script
        dia: Día del mes (1)
        mes: Mes (1, 2, 5, 8)
        hora: Hora de ejecución
        minuto: Minuto de ejecución
    """
    ruta_proyecto = obtener_ruta_proyecto()
    ruta_python = obtener_python_ejecutable()
    
    # Construir el comando usando cd para cambiar al directorio del proyecto
    # Esto hace que el script sea portable (no usa rutas absolutas en el comando)
    if parametros:
        comando_ejecutar = f'cmd /c "cd /d "{ruta_proyecto}" && "{ruta_python}" "{nombre_script}" {parametros}"'
    else:
        comando_ejecutar = f'cmd /c "cd /d "{ruta_proyecto}" && "{ruta_python}" "{nombre_script}""'
    
    # Crear el comando de schtasks para tarea mensual
    schtasks_comando = [
        "schtasks",
        "/create",
        "/tn", nombre_tarea,
        "/tr", comando_ejecutar,
        "/sc", "monthly",
        "/d", str(dia),
        "/m", str(mes).zfill(2),
        "/st", f"{hora}:{str(minuto).zfill(2)}",
        "/f"
    ]
    
    print(f"\n{'='*60}")
    print(f"Creando tarea: {nombre_tarea}")
    print(f"Comando: {comando_ejecutar}")
    print(f"Directorio de trabajo: {ruta_proyecto}")
    print(f"Programación: Día {dia} de cada mes a las {hora}:{str(minuto).zfill(2)}")
    print(f"{'='*60}")
    
    try:
        resultado = subprocess.run(
            schtasks_comando,
            capture_output=True,
            text=True,
            check=False
        )
        
        if resultado.returncode == 0:
            print(f"  Tarea '{nombre_tarea}' creada exitosamente")
        else:
            print(f"  Error al crear tarea '{nombre_tarea}': {resultado.stderr}")
            
    except Exception as e:
        print(f"  Excepción al crear tarea '{nombre_tarea}': {e}")


def crear_tarea_semanal(nombre_tarea, nombre_script, parametros, dia_semana, hora, minuto):
    """
    Crea una tarea programada semanal en Windows usando schtasks
    Usa rutas relativas para que sea portable entre diferentes PC
    
    Args:
        nombre_tarea: Nombre identificador de la tarea
        nombre_script: Nombre del script a ejecutar
        parametros: Parámetros adicionales para el script
        dia_semana: Día de la semana (MON, TUE, WED, THU, FRI, SAT, SUN)
        hora: Hora de ejecución
        minuto: Minuto de ejecución
    """
    ruta_proyecto = obtener_ruta_proyecto()
    ruta_python = obtener_python_ejecutable()
    
    # Construir el comando usando cd para cambiar al directorio del proyecto
    # Esto hace que el script sea portable (no usa rutas absolutas en el comando)
    if parametros:
        comando_ejecutar = f'cmd /c "cd /d "{ruta_proyecto}" && "{ruta_python}" "{nombre_script}" {parametros}"'
    else:
        comando_ejecutar = f'cmd /c "cd /d "{ruta_proyecto}" && "{ruta_python}" "{nombre_script}""'
    
    # Crear el comando de schtasks para tarea semanal
    schtasks_comando = [
        "schtasks",
        "/create",
        "/tn", nombre_tarea,
        "/tr", comando_ejecutar,
        "/sc", "weekly",
        "/d", dia_semana,
        "/st", f"{hora}:{str(minuto).zfill(2)}",
        "/f"
    ]
    
    print(f"\n{'='*60}")
    print(f"Creando tarea: {nombre_tarea}")
    print(f"Comando: {comando_ejecutar}")
    print(f"Directorio de trabajo: {ruta_proyecto}")
    print(f"Programación: Cada {dia_semana} a las {hora}:{str(minuto).zfill(2)}")
    print(f"{'='*60}")
    
    try:
        resultado = subprocess.run(
            schtasks_comando,
            capture_output=True,
            text=True,
            check=False
        )
        
        if resultado.returncode == 0:
            print(f"  Tarea '{nombre_tarea}' creada exitosamente")
        else:
            print(f"  Error al crear tarea '{nombre_tarea}': {resultado.stderr}")
            
    except Exception as e:
        print(f"  Excepción al crear tarea '{nombre_tarea}': {e}")


def mostrar_tareas_existentes():
    """Muestra las tareas programadas existentes relacionadas con el proyecto"""
    print("\n" + "="*60)
    print("Tareas programadas existentes relacionadas con el proyecto:")
    print("="*60)
    
    try:
        resultado = subprocess.run(
            ["schtasks", "/query", "/fo", "list"],
            capture_output=True,
            text=True,
            check=False
        )
        
        lineas = resultado.stdout.split('\n')
        tareas_encontradas = False
        
        for linea in lineas:
            if 'Vivero' in linea or 'clasificacion' in linea or 'Presentacion' in linea or 'Informe' in linea or 'Main' in linea:
                print(linea)
                tareas_encontradas = True
        
        if not tareas_encontradas:
            print("No se encontraron tareas relacionadas")
            
    except Exception as e:
        print(f"Error al consultar tareas: {e}")


def eliminar_tarea(nombre_tarea):
    """Elimina una tarea programada"""
    try:
        resultado = subprocess.run(
            ["schtasks", "/delete", "/tn", nombre_tarea, "/f"],
            capture_output=True,
            text=True,
            check=False
        )
        
        if resultado.returncode == 0:
            print(f"  Tarea '{nombre_tarea}' eliminada exitosamente")
        else:
            print(f"  Error al eliminar tarea '{nombre_tarea}': {resultado.stderr}")
            
    except Exception as e:
        print(f"  Excepción al eliminar tarea: {e}")


def crear_tareas_clasificacion_abc():
    """Crea las 4 tareas programadas para clasificacionABC.py"""
    print("\n" + "-"*60)
    print("Creando tareas para CLASIFICACIONABC.PY")
    print("-"*60)
    
    # Tarea 1: P1 - 1 de enero a las 12:00
    crear_tarea_programada(
        nombre_tarea="Vivero_ClasificacionABC_P1",
        nombre_script="clasificacionABC.py",
        parametros="--P1",
        dia=1,
        mes=1,
        hora=12,
        minuto=0
    )
    
    # Tarea 2: P2 - 1 de febrero a las 09:00
    crear_tarea_programada(
        nombre_tarea="Vivero_ClasificacionABC_P2",
        nombre_script="clasificacionABC.py",
        parametros="--P2",
        dia=1,
        mes=2,
        hora=9,
        minuto=0
    )
    
    # Tarea 3: P3 - 1 de mayo a las 09:00
    crear_tarea_programada(
        nombre_tarea="Vivero_ClasificacionABC_P3",
        nombre_script="clasificacionABC.py",
        parametros="--P3",
        dia=1,
        mes=5,
        hora=9,
        minuto=0
    )
    
    # Tarea 4: P4 - 1 de agosto a las 09:00
    crear_tarea_programada(
        nombre_tarea="Vivero_ClasificacionABC_P4",
        nombre_script="clasificacionABC.py",
        parametros="--P4",
        dia=1,
        mes=8,
        hora=9,
        minuto=0
    )


def crear_tareas_presentacion():
    """Crea las 4 tareas programadas para PRESENTACION.py"""
    print("\n" + "-"*60)
    print("Creando tareas para PRESENTACION.PY")
    print("-"*60)
    
    # Tarea 1: 1 de enero a las 12:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Presentacion_Enero",
        nombre_script="PRESENTACION.py",
        parametros="",
        dia=1,
        mes=1,
        hora=12,
        minuto=30
    )
    
    # Tarea 2: 1 de febrero a las 09:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Presentacion_Febrero",
        nombre_script="PRESENTACION.py",
        parametros="",
        dia=1,
        mes=2,
        hora=9,
        minuto=30
    )
    
    # Tarea 3: 1 de mayo a las 09:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Presentacion_Mayo",
        nombre_script="PRESENTACION.py",
        parametros="",
        dia=1,
        mes=5,
        hora=9,
        minuto=30
    )
    
    # Tarea 4: 1 de agosto a las 09:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Presentacion_Agosto",
        nombre_script="PRESENTACION.py",
        parametros="",
        dia=1,
        mes=8,
        hora=9,
        minuto=30
    )


def crear_tareas_informe():
    """Crea las 4 tareas programadas para INFORME.py"""
    print("\n" + "-"*60)
    print("Creando tareas para INFORME.PY")
    print("-"*60)
    
    # Tarea 1: 1 de enero a las 12:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Informe_Enero",
        nombre_script="INFORME.py",
        parametros="",
        dia=1,
        mes=1,
        hora=12,
        minuto=30
    )
    
    # Tarea 2: 1 de febrero a las 09:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Informe_Febrero",
        nombre_script="INFORME.py",
        parametros="",
        dia=1,
        mes=2,
        hora=9,
        minuto=30
    )
    
    # Tarea 3: 1 de mayo a las 09:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Informe_Mayo",
        nombre_script="INFORME.py",
        parametros="",
        dia=1,
        mes=5,
        hora=9,
        minuto=30
    )
    
    # Tarea 4: 1 de agosto a las 09:30
    crear_tarea_programada(
        nombre_tarea="Vivero_Informe_Agosto",
        nombre_script="INFORME.py",
        parametros="",
        dia=1,
        mes=8,
        hora=9,
        minuto=30
    )


def crear_tareas_main():
    """Crea la tarea programada semanal para main.py"""
    print("\n" + "-"*60)
    print("Creando tarea para MAIN.PY (Pedidos semanales)")
    print("-"*60)
    
    # Tarea semanal: Todos los jueves a las 21:00
    crear_tarea_semanal(
        nombre_tarea="Vivero_Main_Pedidos_Semanales",
        nombre_script="main.py",
        parametros="",
        dia_semana="THU",
        hora=21,
        minuto=0
    )


def crear_tareas_informes_adicionales():
    """Crea las 3 tareas programadas semanales para informes adicionales"""
    print("\n" + "-"*60)
    print("Creando tareas para INFORMES ADICIONALES (jueves 21:10)")
    print("-"*60)
    
    # Tarea 1: informe_compras_sin_autorizacion.py
    crear_tarea_semanal(
        nombre_tarea="Vivero_Informe_Compras_Sin_Autorizacion",
        nombre_script="informe_compras_sin_autorizacion.py",
        parametros="",
        dia_semana="THU",
        hora=21,
        minuto=10
    )
    
    # Tarea 2: Informe_artículos_no_comprados.py
    crear_tarea_semanal(
        nombre_tarea="Vivero_Informe_Articulos_No_Comprados",
        nombre_script="Informe_artículos_no_comprados.py",
        parametros="",
        dia_semana="THU",
        hora=21,
        minuto=10
    )
    
    # Tarea 3: analisis_categoria_cd.py
    crear_tarea_semanal(
        nombre_tarea="Vivero_Analisis_Categoria_CD",
        nombre_script="analisis_categoria_cd.py",
        parametros="",
        dia_semana="THU",
        hora=21,
        minuto=10
    )


def crear_todas_las_tareas():
    """Crea todas las tareas programadas"""
    print("\n" + "="*60)
    print("CREANDO TODAS LAS TAREAS PROGRAMADAS")
    print("="*60)
    print(f"Fecha de creación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Ruta del proyecto: {obtener_ruta_proyecto()}")
    print(f"Ruta de Python detectada: {obtener_python_ejecutable()}")
    
    # Crear tareas para cada script
    crear_tareas_main()
    crear_tareas_informes_adicionales()
    crear_tareas_clasificacion_abc()
    crear_tareas_presentacion()
    crear_tareas_informe()
    
    print("\n" + "="*60)
    print("RESUMEN DE TAREAS CREADAS")
    print("="*60)
    print("\nPedidos Semanales (main.py):")
    print("  - Todos los jueves a las 21:00 (con corrección automática)")
    print("\nInformes Adicionales (jueves 21:10):")
    print("  - informe_compras_sin_autorizacion.py")
    print("  - Informe_artículos_no_comprados.py")
    print("  - analisis_categoria_cd.py")
    print("\nClasificación ABC (clasificacionABC.py):")
    print("  - P1: 1 de enero a las 12:00")
    print("  - P2: 1 de febrero a las 09:00")
    print("  - P3: 1 de mayo a las 09:00")
    print("  - P4: 1 de agosto a las 09:00")
    print("\nPresentación (PRESENTACION.py):")
    print("  - Enero: 1 de enero a las 12:30")
    print("  - Febrero: 1 de febrero a las 09:30")
    print("  - Mayo: 1 de mayo a las 09:30")
    print("  - Agosto: 1 de agosto a las 09:30")
    print("\nInforme (INFORME.py):")
    print("  - Enero: 1 de enero a las 12:30")
    print("  - Febrero: 1 de febrero a las 09:30")
    print("  - Mayo: 1 de mayo a las 09:30")
    print("  - Agosto: 1 de agosto a las 09:30")
    print("="*60)


def eliminar_todas_las_tareas():
    """Elimina todas las tareas programadas"""
    print("\nEliminando todas las tareas...")
    
    # Tarea de Pedidos Semanales
    eliminar_tarea("Vivero_Main_Pedidos_Semanales")
    
    # Tareas de Informes Adicionales
    eliminar_tarea("Vivero_Informe_Compras_Sin_Autorizacion")
    eliminar_tarea("Vivero_Informe_Articulos_No_Comprados")
    eliminar_tarea("Vivero_Analisis_Categoria_CD")
    
    # Tareas de Clasificación ABC
    eliminar_tarea("Vivero_ClasificacionABC_P1")
    eliminar_tarea("Vivero_ClasificacionABC_P2")
    eliminar_tarea("Vivero_ClasificacionABC_P3")
    eliminar_tarea("Vivero_ClasificacionABC_P4")
    
    # Tareas de Presentación
    eliminar_tarea("Vivero_Presentacion_Enero")
    eliminar_tarea("Vivero_Presentacion_Febrero")
    eliminar_tarea("Vivero_Presentacion_Mayo")
    eliminar_tarea("Vivero_Presentacion_Agosto")
    
    # Tareas de Informe
    eliminar_tarea("Vivero_Informe_Enero")
    eliminar_tarea("Vivero_Informe_Febrero")
    eliminar_tarea("Vivero_Informe_Mayo")
    eliminar_tarea("Vivero_Informe_Agosto")
    
    print("Todas las tareas han sido eliminadas")


def mostrar_menu():
    """Muestra el menú de opciones"""
    print("\n" + "="*60)
    print("MENÚ DE TAREAS PROGRAMADAS - VIVEVERDE")
    print("="*60)
    print("1. Crear todas las tareas programadas")
    print("2. Crear tarea de Pedidos Semanales (main.py - jueves 21:00)")
    print("3. Crear tareas de Informes Adicionales (jueves 21:10)")
    print("4. Crear solo tareas de Clasificación ABC")
    print("5. Crear solo tareas de PRESENTACION")
    print("6. Crear solo tareas de INFORME")
    print("7. Mostrar tareas existentes")
    print("8. Eliminar tarea de Pedidos Semanales")
    print("9. Eliminar tarea de Informes Adicionales")
    print("10. Eliminar tarea de Clasificación ABC (especificar)")
    print("11. Eliminar tarea de PRESENTACION")
    print("12. Eliminar tarea de INFORME")
    print("13. Eliminar todas las tareas")
    print("14. Salir")
    print("="*60)


def main():
    """Función principal"""
    while True:
        mostrar_menu()
        opcion = input("\nSelecciona una opción (1-12): ").strip()
        
        if opcion == "1":
            crear_todas_las_tareas()
        elif opcion == "2":
            crear_tareas_main()
        elif opcion == "3":
            crear_tareas_informes_adicionales()
        elif opcion == "4":
            crear_tareas_clasificacion_abc()
        elif opcion == "5":
            crear_tareas_presentacion()
        elif opcion == "6":
            crear_tareas_informe()
        elif opcion == "7":
            mostrar_tareas_existentes()
        elif opcion == "8":
            eliminar_tarea("Vivero_Main_Pedidos_Semanales")
        elif opcion == "9":
            print("\nSelecciona la tarea a eliminar:")
            print("1. informe_compras_sin_autorizacion.py")
            print("2. Informe_artículos_no_comprados.py")
            print("3. analisis_categoria_cd.py")
            print("4. Todas las anteriores")
            subopcion = input("Opción (1-4): ").strip()
            if subopcion == "1":
                eliminar_tarea("Vivero_Informe_Compras_Sin_Autorizacion")
            elif subopcion == "2":
                eliminar_tarea("Vivero_Informe_Articulos_No_Comprados")
            elif subopcion == "3":
                eliminar_tarea("Vivero_Analisis_Categoria_CD")
            elif subopcion == "4":
                eliminar_tarea("Vivero_Informe_Compras_Sin_Autorizacion")
                eliminar_tarea("Vivero_Informe_Articulos_No_Comprados")
                eliminar_tarea("Vivero_Analisis_Categoria_CD")
        elif opcion == "10":
            print("\nSelecciona la tarea a eliminar:")
            print("1. P1 (Enero - 12:00)")
            print("2. P2 (Febrero - 09:00)")
            print("3. P3 (Mayo - 09:00)")
            print("4. P4 (Agosto - 09:00)")
            subopcion = input("Opción (1-4): ").strip()
            if subopcion == "1":
                eliminar_tarea("Vivero_ClasificacionABC_P1")
            elif subopcion == "2":
                eliminar_tarea("Vivero_ClasificacionABC_P2")
            elif subopcion == "3":
                eliminar_tarea("Vivero_ClasificacionABC_P3")
            elif subopcion == "4":
                eliminar_tarea("Vivero_ClasificacionABC_P4")
        elif opcion == "9":
            print("\nSelecciona la tarea a eliminar:")
            print("1. Enero (12:30)")
            print("2. Febrero (09:30)")
            print("3. Mayo (09:30)")
            print("4. Agosto (09:30)")
            subopcion = input("Opción (1-4): ").strip()
            if subopcion == "1":
                eliminar_tarea("Vivero_Presentacion_Enero")
            elif subopcion == "2":
                eliminar_tarea("Vivero_Presentacion_Febrero")
            elif subopcion == "3":
                eliminar_tarea("Vivero_Presentacion_Mayo")
            elif subopcion == "4":
                eliminar_tarea("Vivero_Presentacion_Agosto")
        elif opcion == "10":
            print("\nSelecciona la tarea a eliminar:")
            print("1. Enero (12:30)")
            print("2. Febrero (09:30)")
            print("3. Mayo (09:30)")
            print("4. Agosto (09:30)")
            subopcion = input("Opción (1-4): ").strip()
            if subopcion == "1":
                eliminar_tarea("Vivero_Informe_Enero")
            elif subopcion == "2":
                eliminar_tarea("Vivero_Informe_Febrero")
            elif subopcion == "3":
                eliminar_tarea("Vivero_Informe_Mayo")
            elif subopcion == "4":
                eliminar_tarea("Vivero_Informe_Agosto")
        elif opcion == "11":
            confirmar = input("¿Estás seguro de eliminar todas las tareas? (s/n): ").strip().lower()
            if confirmar == "s":
                eliminar_todas_las_tareas()
            else:
                print("Operación cancelada")
        elif opcion == "14":
            print("\n¡Hasta luego!")
            break
        else:
            print("\nOpción no válida. Por favor, selecciona 1-14.")
        
        input("\nPresiona Enter para continuar...")


if __name__ == "__main__":
    # Si se pasa el argumento --auto, crea las tareas sin preguntar
    if len(sys.argv) > 1 and sys.argv[1] == "--auto":
        crear_todas_las_tareas()
    else:
        main()
