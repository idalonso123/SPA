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
HISTORICO_COMPRAS_SIN_PEDIDO = OUTPUT_DIR / "compras_sin_pedido_historico.json"

# Archivos de compras
ARCHIVO_COMPRAS = INPUT_DIR / "SPA_compras.xlsx"

# ==============================================================================
# PATRONES DE ARCHIVOS CLASIFICACIÓN ABC
# ==============================================================================

# Patrón glob para buscar archivos de clasificación ABC
PATRON_CLASIFICACION_ABC = str(INPUT_DIR / "CLASIFICACION_ABC+D_*.xlsx")

# Diccionario de archivos ABC por categoría (se genera dinámicamente)
# Esta estructura permite acceder directamente por nombre de categoría
ARCHIVOS_CLASIFICACION_ABC = {
    "interior": INPUT_DIR / "CLASIFICACION_ABC+D_INTERIOR_P1_2025.xlsx",
    "maf": INPUT_DIR / "CLASIFICACION_ABC+D_MAF_P1_2025.xlsx",
    "deco_interior": INPUT_DIR / "CLASIFICACION_ABC+D_DECO_INTERIOR_P1_2025.xlsx",
    "deco_exterior": INPUT_DIR / "CLASIFICACION_ABC+D_DECO_EXTERIOR_P1_2025.xlsx",
    "semillas": INPUT_DIR / "CLASIFICACION_ABC+D_SEMILLAS_P1_2025.xlsx",
    "mascotas_vivo": INPUT_DIR / "CLASIFICACION_ABC+D_MASCOTAS_VIVO_P1_2025.xlsx",
    "mascotas_manufacturado": INPUT_DIR / "CLASIFICACION_ABC+D_MASCOTAS_MANUFACTURADO_P1_2025.xlsx",
    "fitos": INPUT_DIR / "CLASIFICACION_ABC+D_FITOS_P1_2025.xlsx",
    "vivero": INPUT_DIR / "CLASIFICACION_ABC+D_VIVERO_P1_2025.xlsx",
    "utiles_jardin": INPUT_DIR / "CLASIFICACION_ABC+D_UTILES_JARDIN_P1_2025.xlsx",
    "tierras_aridos": INPUT_DIR / "CLASIFICACION_ABC+D_TIERRA_ARIDOS_P1_2025.xlsx",
}

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


def get_ruta_clasificacion_abc(categoria: str) -> Path | None:
    """
    Obtiene la ruta del archivo de clasificación ABC para una categoría específica.
    
    Args:
        categoria: Nombre de la categoría (debe coincidir con las claves en ARCHIVOS_CLASIFICACION_ABC)
        
    Returns:
        Path al archivo o None si la categoría no existe
    """
    return ARCHIVOS_CLASIFICACION_ABC.get(categoria.lower())


def verificar_directorios() -> bool:
    """
    Verifica que los directorios principales existan.
    
    Returns:
        True si todos los directorios existen, False en caso contrario
    """
    directorios = [INPUT_DIR, OUTPUT_DIR, CONFIG_DIR, LOGS_DIR]
    return all(d.exists() for d in directorios)


def crear_directorios_si_no_existen():
    """Crea los directorios principales si no existen."""
    for directorio in [INPUT_DIR, OUTPUT_DIR, CONFIG_DIR, LOGS_DIR]:
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
