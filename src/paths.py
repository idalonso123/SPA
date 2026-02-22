"""
Módulo centralizado de rutas del proyecto SPA.

Este módulo contiene todas las rutas del proyecto en un único lugar,
facilitando el mantenimiento y permitiendo cambios en un solo punto.

Uso:
    from src.paths import INPUT_DIR, OUTPUT_DIR, ARCHIVO_VENTAS
"""

import os
from pathlib import Path

# ==============================================================================
# CONFIGURACIÓN DE DIRECTORIOS
# ==============================================================================

# Directorio base del proyecto (se calcula automáticamente desde la ubicación de este archivo)
# src/paths.py está en: BASE_DIR/src/paths.py
BASE_DIR = Path(__file__).resolve().parent.parent

# Directorios principales
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
CONFIG_DIR = BASE_DIR / "config"
LOGS_DIR = BASE_DIR / "logs"
DOCS_DIR = BASE_DIR / "docs"

# ==============================================================================
# SUBDIRECTORIOS DE SALIDA
# ==============================================================================

# Subdirectorios para organizar archivos de salida por tipo
PEDIDOS_SEMANALES_DIR = OUTPUT_DIR / "Pedidos_semanales"
INFORMES_DIR = OUTPUT_DIR / "Informes"
PRESENTACIONES_DIR = OUTPUT_DIR / "Presentaciones"
ANALISIS_DIR = OUTPUT_DIR / "Analisis"
ANALISIS_CATEGORIA_CD_DIR = OUTPUT_DIR / "Analisis_categoria_C_y_D"  # Directorio específico para análisis de categorías C y D
COMPARACION_CATEGORIA_CD_DIR = OUTPUT_DIR / "Comparacion_categoria_C_y_D"  # Directorio específico para comparaciones de categorías C y D
RESUMENES_DIR = OUTPUT_DIR / "Resumenes"
COMPRAS_SIN_AUTORIZACION_DIR = OUTPUT_DIR / "Compras_sin_autorizacion"
ARTICULOS_NO_COMPRADOS_DIR = OUTPUT_DIR / "Articulos_no_comprados"

# ==============================================================================
# ARCHIVOS DE DATOS COMUNES
# ==============================================================================

# Archivos de ventas
ARCHIVO_VENTAS = INPUT_DIR / "SPA_ventas.xlsx"
ARCHIVO_VENTAS_SEMANA = INPUT_DIR / "SPA_ventas_semana.xlsx"

# Archivos de stock
ARCHIVO_STOCK_ACTUAL = INPUT_DIR / "SPA_stock_actual.xlsx"
ARCHIVO_STOCK_P1 = INPUT_DIR / "SPA_stock_P1.xlsx"
ARCHIVO_STOCK_P2 = INPUT_DIR / "SPA_stock_P2.xlsx"
ARCHIVO_STOCK_P3 = INPUT_DIR / "SPA_stock_P3.xlsx"
ARCHIVO_STOCK_P4 = INPUT_DIR / "SPA_stock_P4.xlsx"

# Lista de archivos de stock por período
ARCHIVOS_STOCK = {
    "P1": ARCHIVO_STOCK_P1,
    "P2": ARCHIVO_STOCK_P2,
    "P3": ARCHIVO_STOCK_P3,
    "P4": ARCHIVO_STOCK_P4,
}

# Archivos de costes
ARCHIVO_COSTES = INPUT_DIR / "SPA_coste.xlsx"

# Archivos históricos
HISTORICO_COMPRAS_SIN_PEDIDO = DATA_DIR / "compras_sin_pedido_historico.json"

# Archivos de compras
ARCHIVO_COMPRAS = INPUT_DIR / "SPA_compras.xlsx"

# ==============================================================================
# CONFIGURACIÓN DE PERÍODOS - DESDE config_comun.json
# ==============================================================================

def cargar_periodos_desde_config():
    """
    Carga la definición de períodos desde config/comun.json
    
    Returns:
        dict: Diccionario de períodos con su configuración de fechas
    """
    import json
    periodos = {
        "P1": {"mes_inicio": 1, "dia_inicio": 1, "mes_fin": 2, "dia_fin": 28},
        "P2": {"mes_inicio": 3, "dia_inicio": 1, "mes_fin": 5, "dia_fin": 31},
        "P3": {"mes_inicio": 6, "dia_inicio": 1, "mes_fin": 8, "dia_fin": 31},
        "P4": {"mes_inicio": 9, "dia_inicio": 1, "mes_fin": 12, "dia_fin": 31},
    }
    
    try:
        ruta_config = CONFIG_DIR / "config_comun.json"
        if ruta_config.exists():
            with open(ruta_config, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Intentar cargar desde configuracion_periodo_clasificacion
                if 'configuracion_periodo_clasificacion' in config:
                    config_periodos = config['configuracion_periodo_clasificacion'].get('periodos', {})
                    if config_periodos:
                        # Normalizar el formato
                        for key, valor in config_periodos.items():
                            if isinstance(valor, dict):
                                periodos[key] = {
                                    'mes_inicio': valor.get('mes_inicio'),
                                    'dia_inicio': valor.get('dia_inicio'),
                                    'mes_fin': valor.get('mes_fin'),
                                    'dia_fin': valor.get('dia_fin')
                                }
                        print(f"INFO: Períodos cargados desde config/comun.json")
    except Exception as e:
        print(f"ADVERTENCIA: No se pudieron cargar períodos desde config: {e}. Usando valores por defecto.")
    
    return periodos

# Cargar períodos desde configuración
PERIODOS = cargar_periodos_desde_config()


def obtener_periodo_actual(fecha=None):
    """
    Determina el período actual (P1, P2, P3, P4) según la fecha proporcionada.
    
    Args:
        fecha: Objeto datetime. Si es None, usa la fecha actual.
    
    Returns:
        str: Período actual (P1, P2, P3 o P4)
    """
    from datetime import datetime
    if fecha is None:
        fecha = datetime.now()
    
    mes = fecha.month
    
    if mes <= 2:
        return "P1"
    elif mes <= 5:
        return "P2"
    elif mes <= 8:
        return "P3"
    else:
        return "P4"


def obtener_periodo_siguiente(periodo_actual):
    """
    Obtiene el período siguiente al actual.
    
    Args:
        periodo_actual: Período actual (P1, P2, P3, P4)
    
    Returns:
        str: Período siguiente (P2->P3, P3->P4, P4->P1, P1->P2)
    """
    orden_periodos = ["P1", "P2", "P3", "P4"]
    idx = orden_periodos.index(periodo_actual)
    return orden_periodos[(idx + 1) % 4]


def obtener_periodo_anterior(periodo_actual):
    """
    Obtiene el período anterior al actual.
    
    Args:
        periodo_actual: Período actual (P1, P2, P3, P4)
    
    Returns:
        str: Período anterior (P2->P1, P3->P2, P4->P3, P1->P4)
    """
    orden_periodos = ["P1", "P2", "P3", "P4"]
    idx = orden_periodos.index(periodo_actual)
    return orden_periodos[(idx - 1) % 4]


def generar_nombre_archivo_clasificacion(categoria, periodo, año):
    """
    Genera el nombre del archivo de clasificación ABC según la categoría, período y año.
    
    Args:
        categoria: Nombre de la categoría (ej: 'maf', 'interior', 'vivero')
        periodo: Período (ej: 'P1', 'P2', 'P3', 'P4')
        año: Año (ej: 2025)
    
    Returns:
        str: Nombre del archivo (ej: 'CLASIFICACION_ABC+D_MAF_P2_2025.xlsx')
    """
    return f"CLASIFICACION_ABC+D_{categoria.upper()}_{periodo}_{año}.xlsx"


def buscar_archivo_clasificacion(categoria, periodo, año):
    """
    Busca un archivo de clasificación ABC existente.
    
    Args:
        categoria: Nombre de la categoría
        periodo: Período
        año: Año
    
    Returns:
        Path: Ruta al archivo encontrado o None
    """
    nombre_archivo = generar_nombre_archivo_clasificacion(categoria, periodo, año)
    ruta = INPUT_DIR / nombre_archivo
    return ruta if ruta.exists() else None


# Variables globales que se configurarán dinámicamente
PERIODO_ACTUAL = None  # Se configurará al importar
AÑO_ACTUAL = None      # Se configurará al importar


def inicializar_periodos_dinamicos():
    """
    Inicializa las variables globales de período y año basándose en la fecha actual.
    Debe llamarse al inicio del programa principal.
    """
    global PERIODO_ACTUAL, AÑO_ACTUAL
    from datetime import datetime
    
    ahora = datetime.now()
    PERIODO_ACTUAL = obtener_periodo_actual(ahora)
    AÑO_ACTUAL = ahora.year


# Inicializar automáticamente al importar
inicializar_periodos_dinamicos()


# ==============================================================================
# PATRONES DE ARCHIVOS CLASIFICACIÓN ABC
# ==============================================================================

# Patrón glob para buscar archivos de clasificación ABC
PATRON_CLASIFICACION_ABC = str(INPUT_DIR / "CLASIFICACION_ABC+D_*.xlsx")

# ==============================================================================
# PATRONES PARA BÚSQUEDA DINÁMICA
# ==============================================================================

PATRON_STOCK_PERIODO = str(INPUT_DIR / "SPA_stock_P*.xlsx")

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================


def get_ruta_salida(nombre_archivo: str) -> Path:
    """
    Genera la ruta completa para un archivo de salida.
    
    Args:
        nombre_archivo: Nombre del archivo de salida
        
    Returns:
        Path completo al archivo en el directorio de salida
    """
    return OUTPUT_DIR / nombre_archivo


def get_ruta_entrada(nombre_archivo: str) -> Path:
    """
    Genera la ruta completa para un archivo de entrada.
    
    Args:
        nombre_archivo: Nombre del archivo de entrada
        
    Returns:
        Path completo al archivo en el directorio de entrada
    """
    return INPUT_DIR / nombre_archivo


def get_ruta_clasificacion_abc(categoria: str, periodo: str = None, año: int = None) -> Path | None:
    """
    Obtiene la ruta del archivo de clasificación ABC para una categoría específica.
    
    Args:
        categoria: Nombre de la categoría (ej: 'maf', 'interior', 'vivero')
        periodo: Período (ej: 'P1', 'P2'). Si es None, usa el período actual.
        año: Año (ej: 2025). Si es None, usa el año actual.
    
    Returns:
        Path al archivo o None si no existe
    """
    global PERIODO_ACTUAL, AÑO_ACTUAL
    
    # Usar valores por defecto si no se especifican
    if periodo is None:
        periodo = PERIODO_ACTUAL
    if año is None:
        año = AÑO_ACTUAL
    
    # Generar el nombre del archivo dinámicamente
    nombre_archivo = generar_nombre_archivo_clasificacion(categoria, periodo, año)
    ruta = INPUT_DIR / nombre_archivo
    
    # Verificar si existe el archivo
    if ruta.exists():
        return ruta
    
    # Si no existe, buscar cualquier archivo que coincida con la categoría
    import glob
    patron_busqueda = str(INPUT_DIR / f"CLASIFICACION_ABC+D_{categoria.upper()}_*.xlsx")
    archivos_encontrados = glob.glob(patron_busqueda)
    
    if archivos_encontrados:
        return Path(archivos_encontrados[0])
    
    return None


def verificar_directorios() -> bool:
    """
    Verifica que los directorios principales y subdirectorios de salida existan.
    
    Returns:
        True si todos los directorios existen, False en caso contrario
    """
    directorios = [
        INPUT_DIR, OUTPUT_DIR, CONFIG_DIR, LOGS_DIR,
        PEDIDOS_SEMANALES_DIR, INFORMES_DIR, PRESENTACIONES_DIR,
        ANALISIS_DIR, ANALISIS_CATEGORIA_CD_DIR, COMPARACION_CATEGORIA_CD_DIR, RESUMENES_DIR, COMPRAS_SIN_AUTORIZACION_DIR,
        ARTICULOS_NO_COMPRADOS_DIR
    ]
    return all(d.exists() for d in directorios)


def crear_directorios_si_no_existen():
    """Crea los directorios principales y subdirectorios de salida si no existen."""
    # Directorios principales
    directorios = [INPUT_DIR, OUTPUT_DIR, CONFIG_DIR, LOGS_DIR]
    
    # Subdirectorios de salida
    subdirectorios_salida = [
        PEDIDOS_SEMANALES_DIR,
        INFORMES_DIR,
        PRESENTACIONES_DIR,
        ANALISIS_DIR,
        ANALISIS_CATEGORIA_CD_DIR,
        COMPARACION_CATEGORIA_CD_DIR,
        RESUMENES_DIR,
        COMPRAS_SIN_AUTORIZACION_DIR,
        ARTICULOS_NO_COMPRADOS_DIR,
    ]
    
    for directorio in directorios + subdirectorios_salida:
        directorio.mkdir(parents=True, exist_ok=True)


# ==============================================================================
# EXPORTACIÓN DE RUTAS COMO STRINGS (para compatibilidad hacia atrás)
# ==============================================================================

# Variables de clase string para compatibilidad (usar Path cuando sea posible)
INPUT_DIR_STR = str(INPUT_DIR)
OUTPUT_DIR_STR = str(OUTPUT_DIR)
CONFIG_DIR_STR = str(CONFIG_DIR)
LOGS_DIR_STR = str(LOGS_DIR)

# Strings de archivos comunes
ARCHIVO_VENTAS_STR = str(ARCHIVO_VENTAS)
ARCHIVO_COSTES_STR = str(ARCHIVO_COSTES)
ARCHIVO_STOCK_ACTUAL_STR = str(ARCHIVO_STOCK_ACTUAL)

# Strings de directorios de salida
PEDIDOS_SEMANALES_DIR_STR = str(PEDIDOS_SEMANALES_DIR)
INFORMES_DIR_STR = str(INFORMES_DIR)
PRESENTACIONES_DIR_STR = str(PRESENTACIONES_DIR)
ANALISIS_DIR_STR = str(ANALISIS_DIR)
ANALISIS_CATEGORIA_CD_DIR_STR = str(ANALISIS_CATEGORIA_CD_DIR)
COMPARACION_CATEGORIA_CD_DIR_STR = str(COMPARACION_CATEGORIA_CD_DIR)
RESUMENES_DIR_STR = str(RESUMENES_DIR)
COMPRAS_SIN_AUTORIZACION_DIR_STR = str(COMPRAS_SIN_AUTORIZACION_DIR)
ARTICULOS_NO_COMPRADOS_DIR_STR = str(ARTICULOS_NO_COMPRADOS_DIR)
