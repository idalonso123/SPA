#!/usr/bin/env python3
"""
Módulo CorrectionDataLoader - Carga de datos de corrección para FASE 2

Este módulo extiende el DataLoader existente para leer el archivo de stock
actual que alimenta el sistema de corrección de pedidos:
- SPA_stock_actual.xlsx: Inventario disponible al momento del cálculo

La integración de estos datos permite ajustar las proyecciones teóricas de la
FASE 1 contra la realidad operativa del almacén, corrigiendo el pedido generado
en función de la diferencia entre el stock real y el stock mínimo configurado.

Autor: Sistema de Pedidos Vivero V2 - FASE 2
Fecha: 2026-02-03
"""

import pandas as pd
import numpy as np
import os
import glob
import logging
import unicodedata
from typing import Optional, Dict, List, Tuple, Any
from datetime import datetime
from src.data_loader import DataLoader

# Configuración del logger
logger = logging.getLogger(__name__)

# ============================================================================
# FUNCIONES DE NORMALIZACIÓN PARA BÚSQUEDAS INTELIGENTES
# ============================================================================

def normalizar_texto(texto):
    """
    Normaliza un texto para comparación:
    - Convierte a minúsculas
    - Elimina acentos
    - Elimina puntuación (puntos, guiones, espacios, paréntesis, etc.)
    
    Esta función es la base para todas las búsquedas normalizadas en el script.
    
    Ejemplos:
    - 'Cóste' -> 'coste'
    - 'Últ. Comp' -> 'ultcomp'
    - 'ÚLTIMA COMPRA' -> 'ultimacompra'
    - 'Coste Unitario' -> 'costeunitario'
    
    Args:
        texto: Texto a normalizar
    
    Returns:
        str: Texto normalizado (minúsculas, sin acentos, sin puntuación) o cadena vacía si es None/NaN
    """
    if pd.isna(texto):
        return ''
    texto = str(texto)
    # Convertir a minúsculas
    texto = texto.lower()
    # Normalizar unicode: á → a, é → e, ñ → n, etc.
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    # Eliminar puntuación: puntos, guiones, espacios, paréntesis, etc.
    texto = ''.join(c for c in texto if c.isalnum())
    return texto

def normalizar_con_espacios(texto):
    """
    Normaliza un texto conservando los espacios pero eliminando:
    - Mayúsculas/minúsculas
    - Acentos
    - Puntuación (excepto espacios)
    
    Útil cuando quieres mantener pero ignorar detalles.
 la estructura de palabras    
    Ejemplos:
    - 'Cóste' -> 'coste'
    - 'Últ. Comp' -> 'ult comp'
    - 'ÚLTIMA COMPRA' -> 'ultima compra'
    
    Args:
        texto: Texto a normalizar
    
    Returns:
        str: Texto normalizado (minúsculas, sin acentos,保留 espacios, sin otra puntuación)
    """
    if pd.isna(texto):
        return ''
    texto = str(texto)
    # Convertir a minúsculas
    texto = texto.lower()
    # Normalizar unicode: á → a, é → e, ñ → n, etc.
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    # Reemplazar puntuación (excepto espacios) por nada
    texto = ''.join(c if c.isalnum() or c == ' ' else '' for c in texto)
    # Normalizar espacios múltiples
    texto = ' '.join(texto.split())
    return texto

def encontrar_columna(columnas, nombre_buscado):
    """
    Busca una columna por nombre, ignorando mayúsculas, acentos y puntuación.
    
    Esta función permite encontrar columnas aunque sus nombres tengan variaciones:
    - 'Coste', 'COSTE', 'cósté', 'Cóste' → todas dan positivo para 'coste'
    - 'Últ. compra', 'ULTIMA COMPRA', 'ultima compra' → todas dan positivo para 'ultimacompra'
    - 'Últ. Comp', 'Ult. Comp', 'últ comp' → todas dan positivo para 'ultcomp'
    
    Args:
        columnas: Lista de nombres de columnas del DataFrame
        nombre_buscado: Nombre base a buscar (puede tener acentos, mayúsculas, etc.)
    
    Returns:
        str: El nombre real de la columna encontrada, o None si no existe
    """
    nombre_normalizado = normalizar_texto(nombre_buscado)
    
    for columna in columnas:
        if normalizar_texto(columna) == nombre_normalizado:
            return columna
    
    return None

def obtener_columna_segura(df, nombre_buscado):
    """
    Obtiene una columna del DataFrame buscando por nombre normalizado.
    
    Args:
        df: DataFrame con los datos
        nombre_buscado: Nombre base a buscar (ej: 'coste', 'fecha', 'últ. comp')
    
    Returns:
        Series: La columna encontrada, o una serie vacía si no existe
    """
    nombre_real = encontrar_columna(list(df.columns), nombre_buscado)
    if nombre_real:
        return df[nombre_real]
    else:
        logger.warning(f"No se encontró columna '{nombre_buscado}' en el DataFrame")
        return pd.Series([], dtype='object')

def filtrar_por_valor_normalizado(df, nombre_columna, valor_buscado):
    """
    Filtra un DataFrame buscando un valor en una columna específica,
    ignorando mayúsculas, acentos y puntuación.
    
    Esta función es ideal para búsquedas en columnas de texto donde los valores
    pueden tener variaciones en su escritura.
    
    Ejemplos de uso:
    - Buscar 'coste' en columna 'Nombre': encuentra 'Coste', 'COSTE', 'cósté', etc.
    - Buscar 'últ. comp' en columna 'Concepto': encuentra 'Últ. Comp', 'Ult Comp', etc.
    
    Args:
        df: DataFrame con los datos
        nombre_columna: Nombre de la columna donde buscar (se normaliza la búsqueda)
        valor_buscado: Valor a buscar (puede tener variaciones de caso, acentos, etc.)
    
    Returns:
        DataFrame: Filas que coinciden con el valor buscado
    """
    # Normalizar el valor buscado
    valor_normalizado = normalizar_texto(valor_buscado)
    
    # Encontrar la columna (si no existe, retornar DataFrame vacío)
    columna_real = encontrar_columna(list(df.columns), nombre_columna)
    if columna_real is None:
        logger.warning(f"No se encontró la columna '{nombre_columna}' para filtrar")
        return df.iloc[0:0]
    
    # Filtrar comparando valores normalizados
    mask = df[columna_real].apply(lambda x: normalizar_texto(x) == valor_normalizado)
    return df[mask]

def filtrar_por_valor_parcial(df, nombre_columna, fragmento_buscado):
    """
    Filtra un DataFrame buscando un fragmento de texto en una columna,
    ignorando mayúsculas, acentos y puntuación.
    
    Útil cuando quieres encontrar valores que CONTENGAN el texto buscado,
    no solo valores que sean EXACTAMENTE iguales.
    
    Ejemplos de uso:
    - Buscar 'coste' en columna 'Nombre': encuentra cualquier nombre que contenga 'coste'
    - Buscar 'últ' en columna 'Concepto': encuentra 'Últ. Comp', 'Última', 'Último', etc.
    
    Args:
        df: DataFrame con los datos
        nombre_columna: Nombre de la columna donde buscar
        fragmento_buscado: Fragmento de texto a buscar
    
    Returns:
        DataFrame: Filas que contienen el fragmento buscado
    """
    # Normalizar el fragmento buscado
    fragmento_normalizado = normalizar_texto(fragmento_buscado)
    
    # Encontrar la columna
    columna_real = encontrar_columna(list(df.columns), nombre_columna)
    if columna_real is None:
        logger.warning(f"No se encontró la columna '{nombre_columna}' para filtrar")
        return df.iloc[0:0]
    
    # Filtrar buscando el fragmento en valores normalizados
    mask = df[columna_real].apply(lambda x: fragmento_normalizado in normalizar_texto(x))
    return df[mask]

def buscar_valores_unicos_normalizados(df, nombre_columna):
    """
    Obtiene los valores únicos de una columna, junto con su versión normalizada.
    
    Útil para analizar qué valores únicos existen en una columna y poder
    hacer búsquedas normalizadas posteriormente.
    
    Args:
        df: DataFrame con los datos
        nombre_columna: Nombre de la columna a analizar
    
    Returns:
        dict: Diccionario {valor_normalizado: [lista de valores originales]}
    """
    columna_real = encontrar_columna(list(df.columns), nombre_columna)
    if columna_real is None:
        logger.warning(f"No se encontró la columna '{nombre_columna}'")
        return {}
    
    resultado = {}
    for valor in df[columna_real].unique():
        clave = normalizar_texto(valor)
        if clave not in resultado:
            resultado[clave] = []
        if valor not in resultado[clave]:
            resultado[clave].append(valor)
    
    return resultado

# ============================================================================


class CorrectionDataLoader:
    """
    Clase para la carga y normalización de datos de corrección.
    
    Extiende el DataLoader base para manejar el archivo específico
    de stock actual de la FASE 2.
    
    Attributes:
        config (dict): Configuración del sistema
        rutas (dict): Rutas de archivos y directorios
        correction_files (dict): Rutas de archivos de corrección
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el CorrectionDataLoader con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.rutas = config.get('rutas', {})
        self.correction_files = config.get('archivos_correccion', {})
        
        # Usar el DataLoader base para funciones compartidas
        self.base_loader = DataLoader(config)
        
        logger.info("CorrectionDataLoader inicializado correctamente")
    
    def normalizar_texto(self, texto: Any) -> str:
        """
        Normaliza un texto para comparaciones insensibles a mayúsculas y acentos.
        
        Args:
            texto (Any): Texto a normalizar
        
        Returns:
            str: Texto normalizado
        """
        return self.base_loader.normalizar_texto(texto)
    
    def texto_igual(self, texto1: Any, texto2: Any) -> bool:
        """
        Compara si dos textos son iguales ignorando mayúsculas y acentos.
        
        Args:
            texto1 (Any): Primer texto
            texto2 (Any): Segundo texto
        
        Returns:
            bool: True si son iguales
        """
        return self.base_loader.normalizar_texto(texto1) == self.base_loader.normalizar_texto(texto2)
    
    def obtener_directorio_entrada(self) -> str:
        """
        Obtiene el directorio de entrada configurado.
        
        Returns:
            str: Ruta del directorio de entrada
        """
        base = self.rutas.get('directorio_base', '.')
        entrada = self.rutas.get('directorio_entrada', './data/input')
        
        if not os.path.isabs(entrada):
            entrada = os.path.join(base, entrada)
        
        return entrada
    
    def leer_excel(self, ruta_archivo: str, hoja: Optional[str] = None) -> Optional[pd.DataFrame]:
        """
        Lee un archivo Excel y devuelve un DataFrame.
        
        Args:
            ruta_archivo (str): Ruta del archivo Excel
            hoja (Optional[str]): Nombre de la hoja (None para todas)
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con los datos o None si hay error
        """
        try:
            if not os.path.exists(ruta_archivo):
                logger.warning(f"Archivo de corrección no encontrado: {ruta_archivo}")
                return None
            
            logger.info(f"Leyendo archivo de corrección: {ruta_archivo}")
            
            if hoja:
                df = pd.read_excel(ruta_archivo, sheet_name=hoja)
            else:
                df = pd.read_excel(ruta_archivo, sheet_name=None)
            
            logger.info(f"Archivo leído exitosamente: {len(df) if isinstance(df, pd.DataFrame) else len(df)} hojas")
            return df
            
        except Exception as e:
            logger.error(f"Error al leer archivo {ruta_archivo}: {str(e)}")
            return None
    
    def buscar_archivo_correccion(self, nombre_archivo: str) -> Optional[str]:
        """
        Busca un archivo de corrección en el directorio de entrada.
        
        Args:
            nombre_archivo (str): Nombre del archivo a buscar
        
        Returns:
            Optional[str]: Ruta completa del archivo o None si no existe
        """
        dir_entrada = self.obtener_directorio_entrada()
        ruta_archivo = os.path.join(dir_entrada, nombre_archivo)
        
        if os.path.exists(ruta_archivo):
            logger.info(f"Archivo de corrección encontrado: {ruta_archivo}")
            return ruta_archivo
        
        # Intentar búsqueda con wildcards
        patron_buscar = os.path.join(dir_entrada, f"*{nombre_archivo}*")
        archivos_encontrados = glob.glob(patron_buscar)
        
        if archivos_encontrados:
            logger.info(f"Archivo encontrado (búsqueda amplia): {archivos_encontrados[0]}")
            return archivos_encontrados[0]
        
        logger.warning(f"Archivo de corrección no encontrado: {nombre_archivo}")
        return None
    
    def leer_stock_actual(self, semana: Optional[int] = None) -> Optional[pd.DataFrame]:
        """
        Lee el archivo de stock actual (SPA_stock_actual.xlsx).
        
        Este archivo contiene el inventario disponible al momento del cálculo,
        incluyendo código de artículo, nombre, talla, color, unidades en stock,
        fecha del último movimiento y antigüedad del stock.
        
        Args:
            semana (Optional[int]): Número de semana para buscar archivo específico
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con el stock actual o None si hay error
        """
        nombre_base = self.correction_files.get('stock_actual', 'SPA_stock_actual.xlsx')
        
        # Si se especifica semana, buscar con patrón de semana
        if semana:
            # Intentar buscar archivo con semana en el nombre
            nombre_con_semana = nombre_base.replace('.xlsx', f'_Semana_{semana}.xlsx')
            ruta = self.buscar_archivo_correccion(nombre_con_semana)
            
            if ruta is None:
                # Buscar con otro patrón común
                nombre_con_semana = f"Stock_semana_{semana}.xlsx"
                ruta = self.buscar_archivo_correccion(nombre_con_semana)
        
        # Si no se encontró archivo con semana, usar el base
        if semana is None or not self.buscar_archivo_correccion(nombre_base):
            # Verificar si existe el archivo base
            dir_entrada = self.obtener_directorio_entrada()
            ruta_base = os.path.join(dir_entrada, nombre_base)
            if os.path.exists(ruta_base):
                ruta = ruta_base
            else:
                logger.warning(f"No se encontró archivo de stock actual")
                return None
        
        if semana:
            ruta = self.buscar_archivo_correccion(nombre_base)
        
        df = self.leer_excel(ruta)
        
        if df is None:
            return None
        
        # Si devuelve diccionario (múltiples hojas), tomar la primera
        if isinstance(df, dict):
            primera_hoja = list(df.keys())[0]
            df = df[primera_hoja]
            logger.debug(f"Usando hoja: {primera_hoja}")
        
        df = df.copy()
        
        # Normalizar columna de código de artículo
        self._normalizar_columnas_stock(df)
        
        # Validar que tenemos columna de stock
        if 'Stock_Fisico' not in df.columns and 'Stock' not in df.columns:
            # Buscar columna que contenga 'stock'
            for col in df.columns:
                if 'stock' in self.normalizar_texto(col):
                    df.rename(columns={col: 'Stock_Fisico'}, inplace=True)
                    break
        
        logger.info(f"Stock actual cargado: {len(df)} registros")
        return df
    
    def _normalizar_columnas_stock(self, df: pd.DataFrame) -> None:
        """
        Normaliza las columnas del DataFrame de stock.
        
        Args:
            df (pd.DataFrame): DataFrame a normalizar
        """
        # Renombrar columnas comunes a nombres estándar
        mapeo_columnas = {}
        
        for col in df.columns:
            col_norm = self.normalizar_texto(col)
            
            if 'articulo' in col_norm and 'codigo' in col_norm:
                mapeo_columnas[col] = 'Codigo_Articulo'
            elif 'codigo' in col_norm:
                mapeo_columnas[col] = 'Codigo_Articulo'
            elif 'nombre' in col_norm and 'articulo' in col_norm:
                mapeo_columnas[col] = 'Nombre_Articulo'
            elif col_norm == 'nombre':
                mapeo_columnas[col] = 'Nombre_Articulo'
            elif 'stock' in col_norm and ('fisico' in col_norm or 'actual' in col_norm or col_norm == 'stock'):
                mapeo_columnas[col] = 'Stock_Fisico'
            elif 'unidades' in col_norm and 'stock' in col_norm:
                mapeo_columnas[col] = 'Stock_Fisico'
            elif 'talla' in col_norm:
                mapeo_columnas[col] = 'Talla'
            elif 'color' in col_norm:
                mapeo_columnas[col] = 'Color'
            elif 'fecha' in col_norm and 'ultimo' in col_norm:
                mapeo_columnas[col] = 'Fecha_Ultimo_Movimiento'
            elif 'antiguedad' in col_norm:
                mapeo_columnas[col] = 'Antiguedad_Stock'
        
        if mapeo_columnas:
            df.rename(columns=mapeo_columnas, inplace=True)
            logger.debug(f"Columnas renombradas en stock: {list(mapeo_columnas.values())}")
    
    def cargar_datos_correccion(self, semana: Optional[int] = None) -> Dict[str, Optional[pd.DataFrame]]:
        """
        Carga los datos de corrección para una semana específica.
        
        Args:
            semana (Optional[int]): Número de semana para la que cargar datos
        
        Returns:
            Dict[str, Optional[pd.DataFrame]]: Diccionario con:
                - 'stock': DataFrame de stock actual
        """
        logger.info("=" * 60)
        logger.info(f"CARGANDO DATOS DE CORRECCIÓN PARA SEMANA {semana if semana else 'actual'}")
        logger.info("=" * 60)
        
        datos = {
            'stock': None
        }
        
        # Cargar stock actual
        datos['stock'] = self.leer_stock_actual(semana)
        if datos['stock'] is not None:
            logger.info(f"  Stock: {len(datos['stock'])} registros")
        else:
            logger.warning("No se encontraron datos de stock para la corrección")
        
        return datos
    
    def merge_con_pedido_teorico(
        self, 
        pedido_teorico: pd.DataFrame, 
        datos_correccion: Dict[str, Optional[pd.DataFrame]],
        clave_cols: List[str] = ['Codigo_Articulo', 'Talla', 'Color']
    ) -> pd.DataFrame:
        """
        Fusiona los datos de corrección con el pedido teórico de la FASE 1.
        
        Realiza un LEFT JOIN para mantener todos los artículos del pedido
        teórico, añadiendo las columnas de datos de stock real.
        
        Args:
            pedido_teorico (pd.DataFrame): DataFrame del pedido generado en FASE 1
            datos_correccion (Dict): Diccionario con datos de corrección
            clave_cols (List[str]): Columnas usadas como clave de unión
        
        Returns:
            pd.DataFrame: DataFrame fusionado con todos los datos
        """
        logger.info("Fusionando datos de corrección con pedido teórico...")
        
        df = pedido_teorico.copy()
        
        # Preparar claves de unión normalizadas
        df['_clave'] = (
            df.get('Codigo_Articulo', df.get('Código artículo', df.get('Codigo', ''))).astype(str) + '|' +
            df.get('Talla', '').astype(str) + '|' +
            df.get('Color', '').astype(str)
        )
        
        # Fusionar stock actual
        if datos_correccion['stock'] is not None:
            stock_df = datos_correccion['stock'].copy()
            stock_df['_clave'] = (
                stock_df.get('Codigo_Articulo', '').astype(str) + '|' +
                stock_df.get('Talla', '').astype(str) + '|' +
                stock_df.get('Color', '').astype(str)
            )
            
            # Seleccionar columnas relevantes
            cols_stock = ['_clave', 'Stock_Fisico']
            cols_disponibles = [c for c in cols_stock if c in stock_df.columns]
            stock_df = stock_df[cols_disponibles]
            
            # Agrupar por clave (si hay duplicados)
            stock_df = stock_df.groupby('_clave').agg({
                'Stock_Fisico': 'sum' if 'Stock_Fisico' in stock_df.columns else 'first'
            }).reset_index()
            
            df = df.merge(stock_df, on='_clave', how='left')
            
            # Rellenar NaN con 0
            if 'Stock_Fisico' in df.columns:
                df['Stock_Fisico'] = df['Stock_Fisico'].fillna(0)
        
        # Limpiar columna de clave temporal
        df.drop(columns=['_clave'], inplace=True, errors='ignore')
        
        # Añadir columna de stock faltante con valor por defecto
        if 'Stock_Fisico' not in df.columns:
            df['Stock_Fisico'] = 0
        
        logger.info(f"Fusión completada: {len(df)} registros")
        
        return df


# Función de utilidad para uso directo
def crear_correction_data_loader(config: dict) -> CorrectionDataLoader:
    """
    Crea una instancia del CorrectionDataLoader.
    
    Args:
        config (dict): Configuración del sistema
    
    Returns:
        CorrectionDataLoader: Instancia del loader de corrección
    """
    return CorrectionDataLoader(config)


if __name__ == "__main__":
    # Ejemplo de uso
    print("CorrectionDataLoader - Módulo de carga de datos de corrección FASE 2")
    print("=" * 60)
    
    # Configurar logging
    import logging
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración mínima
    config_ejemplo = {
        'rutas': {
            'directorio_base': '.',
            'directorio_entrada': './data/input'
        },
        'archivos_correccion': {
            'stock_actual': 'SPA_stock_actual.xlsx'
        }
    }
    
    loader = CorrectionDataLoader(config_ejemplo)
    print("CorrectionDataLoader inicializado y listo para usar.")


# ============================================================================
# FUNCIONES PARA CÁLCULO DE TENDENCIA DE VENTAS
# ============================================================================

def encontrar_archivo_semana_anterior(directorio_base: str, semana_actual: int, seccion: Optional[str] = None) -> Optional[str]:
    """
    Busca el archivo de pedido de la semana anterior basándose en la semana actual.
    
    El sistema genera archivos con el patrón: Pedido_Semana_NN_DDMMAAAA.xlsx
    Esta función busca el archivo correspondiente a la semana anterior (semana_actual - 1).
    
    Args:
        directorio_base (str): Directorio donde buscar los archivos de pedido
        semana_actual (int): Número de la semana actual (1-52)
        seccion (Optional[str]): Nombre de la sección para filtrar (ej: 'interior', 'vivero')
                                 Si es None, busca cualquier archivo de la semana anterior
    
    Returns:
        Optional[str]: Ruta completa del archivo encontrado, o None si no existe
    
    Ejemplo:
        Si semana_actual = 9 y seccion = 'interior', buscará: Pedido_Semana_8_*_interior.xlsx
    """
    # Calcular semana anterior
    semana_anterior = semana_actual - 1
    
    # Manejo especial para semana 1 (buscar última semana del año anterior)
    if semana_anterior < 1:
        semana_anterior = 52  # Asumimos que el año anterior tenía 52 semanas
    
    # Patrón de búsqueda para archivos de la semana anterior (formato de 2 dígitos: 07, 08, etc.)
    if seccion:
        # Filtrar por sección: Pedido_Semana_08_*_interior.xlsx
        patron_busqueda = f"Pedido_Semana_{semana_anterior:02d}_*_{seccion}.xlsx"
    else:
        # Buscar cualquier archivo de la semana anterior
        patron_busqueda = f"Pedido_Semana_{semana_anterior:02d}_*.xlsx"
    
    ruta_patron = os.path.join(directorio_base, patron_busqueda)
    
    # Buscar archivos que coincidan con el patrón
    archivos_encontrados = glob.glob(ruta_patron)
    
    if archivos_encontrados:
        # Si hay múltiples archivos, seleccionar el más reciente por fecha de modificación
        archivo_mas_reciente = max(archivos_encontrados, key=os.path.getmtime)
        logger.info(f"Archivo de semana anterior encontrado: {os.path.basename(archivo_mas_reciente)}")
        return archivo_mas_reciente
    
    if seccion:
        logger.warning(f"No se encontró archivo de pedido para la semana {semana_anterior} sección '{seccion}'")
    else:
        logger.warning(f"No se encontró archivo de pedido para la semana {semana_anterior}")
    
    logger.info(f"  Patrón buscado: {patron_busqueda}")
    logger.info(f"  Directorio: {directorio_base}")
    return None


def leer_archivo_ventas_reales(directorio_entrada: str) -> Tuple[Optional[pd.DataFrame], bool]:
    """
    Lee el archivo de ventas reales del ERP (SPA_ventas_reales.xlsx).
    
    Este archivo tiene un nombre fijo porque el ERP no permite nombres dinámicos
    con la semana actual. Si el archivo no existe, genera una advertencia pero
    no interrumpe el proceso.
    
    Args:
        directorio_entrada (str): Directorio donde se encuentra el archivo
    
    Returns:
        Tuple[Optional[pd.DataFrame], bool]: 
            - DataFrame con las ventas reales o None si hay error
            - Boolean indicando si el archivo existía
    
    Nota:
        El archivo no encontrado genera una ADVERTENCIA (no un error), ya que
        el sistema debe continuar generando los pedidos incluso sin este dato.
        Esta advertencia será utilizada posteriormente para el sistema de notificaciones por email.
    """
    nombre_archivo = "SPA_ventas_reales.xlsx"
    ruta_archivo = os.path.join(directorio_entrada, nombre_archivo)
    
    if not os.path.exists(ruta_archivo):
        logger.warning(f"ADVERTENCIA: No se encontró el archivo de ventas reales")
        logger.warning(f"  Archivo: {nombre_archivo}")
        logger.warning(f"  Directorio: {directorio_entrada}")
        logger.warning("  El sistema continuará generando pedidos sin el dato de ventas reales.")
        logger.warning("  NOTA: Esta advertencia será enviada por email en una futura actualización.")
        return None, False
    
    try:
        df = pd.read_excel(ruta_archivo)
        logger.info(f"Archivo de ventas reales cargado: {len(df)} registros")
        return df, True
    except Exception as e:
        logger.error(f"Error al leer el archivo de ventas reales: {str(e)}")
        logger.warning(f"ADVERTENCIA: Error al leer {nombre_archivo}. El sistema continuará sin datos de ventas reales.")
        return None, False


def leer_archivo_stock_actual(directorio_entrada: str) -> Optional[pd.DataFrame]:
    """
    Lee el archivo de stock actual del ERP (SPA_stock_actual.xlsx).
    
    Args:
        directorio_entrada (str): Directorio donde se encuentra el archivo
    
    Returns:
        Optional[pd.DataFrame]: DataFrame con el stock actual o None si hay error
    """
    nombre_archivo = "SPA_stock_actual.xlsx"
    ruta_archivo = os.path.join(directorio_entrada, nombre_archivo)
    
    if not os.path.exists(ruta_archivo):
        logger.warning(f"No se encontró el archivo de stock actual: {nombre_archivo}")
        return None
    
    try:
        df = pd.read_excel(ruta_archivo)
        logger.info(f"Archivo de stock actual cargado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Error al leer el archivo de stock actual: {str(e)}")
        return None


def normalizar_datos_historicos(df_pedido_anterior: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza el DataFrame del pedido de la semana anterior para obtener
    solo las columnas necesarias para el cálculo de tendencia.
    
    Ahora usa 'Unidades_Finales' (Unidades Calculadas) en lugar de 'Ventas_Objetivo'
    para mostrar las unidades objetivo de la semana anterior.
    
    Args:
        df_pedido_anterior (pd.DataFrame): DataFrame del archivo Pedido_Semana_*.xlsx
    
    Returns:
        pd.DataFrame: DataFrame normalizado con clave de artículo y Unidades_Calculadas_Semana_Pasada
    """
    if df_pedido_anterior is None or len(df_pedido_anterior) == 0:
        return pd.DataFrame()
    
    # Usar encontrar_columna para buscar la columna de código (puede ser 'Codigo_Articulo' o 'Código artículo')
    col_codigo = encontrar_columna(list(df_pedido_anterior.columns), 'codigoarticulo')
    
    if col_codigo:
        df = df_pedido_anterior.copy()
        
        # Buscar columnas de talla y color
        col_talla = encontrar_columna(list(df.columns), 'talla')
        col_color = encontrar_columna(list(df.columns), 'color')
        
        # Crear clave compuesta si las columnas existen
        # Convertir a string y eliminar decimales si viene de formato numérico
        df[col_codigo] = df[col_codigo].astype(str).str.replace(r'\.0$', '', regex=True)
        if col_talla and col_color:
            df['Clave_Articulo'] = (df[col_codigo] + '|' + 
                                   df[col_talla].astype(str) + '|' + 
                                   df[col_color].astype(str))
        else:
            df['Clave_Articulo'] = df[col_codigo]
        
        # NUEVO: Buscar primero 'Unidades_Finales' (Unidades Calculadas) - prioridad absoluta
        # Esto es lo que el usuario quiere: las unidades calculadas de la semana anterior
        if 'Unidades_Finales' in df.columns:
            df_resultado = df[['Clave_Articulo', 'Unidades_Finales']].copy()
            df_resultado.rename(columns={'Unidades_Finales': 'Unidades_Calculadas_Semana_Pasada'}, inplace=True)
            logger.info("Usando 'Unidades_Finales' para unidades calculadas de semana anterior")
        else:
            # Si no existe Unidades_Finales, buscar variaciones de nombre
            col_unidades_calc = encontrar_columna(list(df.columns), 'unidadescalculadas')
            if col_unidades_calc:
                df_resultado = df[['Clave_Articulo', col_unidades_calc]].copy()
                df_resultado.rename(columns={col_unidades_calc: 'Unidades_Calculadas_Semana_Pasada'}, inplace=True)
                logger.info(f"Usando '{col_unidades_calc}' para unidades calculadas de semana anterior")
            else:
                logger.warning("No se encontró columna 'Unidades_Finales' ni 'Unidades Calculadas' en el archivo de la semana anterior")
                return pd.DataFrame()
        
        # Eliminar duplicados (quedarse con el último registro de cada artículo)
        df_resultado.drop_duplicates(subset=['Clave_Articulo'], keep='last', inplace=True)
        
        return df_resultado
    
    logger.warning("No se encontró columna de código de artículo en el archivo de la semana anterior")
    return pd.DataFrame()


def fusionar_datos_tendencia(
    pedidos_df: pd.DataFrame,
    df_ventas_reales: Optional[pd.DataFrame],
    df_stock_actual: Optional[pd.DataFrame],
    df_ventas_objetivo_anterior: Optional[pd.DataFrame]
) -> pd.DataFrame:
    """
    Fusiona los datos históricos (unidades calculadas semana anterior, ventas reales,
    stock actual) con el DataFrame de pedidos actual para calcular la tendencia.
    
    Nota: Ahora usa 'Unidades_Calculadas_Semana_Pasada' en lugar de 'Ventas_Objetivo_Semana_Pasada'
    para mostrar las unidades calculadas de la semana anterior.
    
    Args:
        pedidos_df (pd.DataFrame): DataFrame con los pedidos calculados
        df_ventas_reales (Optional[pd.DataFrame]): Ventas reales del ERP
        df_stock_actual (Optional[pd.DataFrame]): Stock actual del ERP
        df_ventas_objetivo_anterior (Optional[pd.DataFrame]): Unidades calculadas de la semana anterior
    
    Returns:
        pd.DataFrame: DataFrame con las nuevas columnas fusionadas
    """
    if pedidos_df is None or len(pedidos_df) == 0:
        return pedidos_df
    
    df_resultado = pedidos_df.copy()
    
    # Inicializar nuevas columnas con valores por defecto
    df_resultado['Unidades_Calculadas_Semana_Pasada'] = 0
    df_resultado['Ventas_Reales'] = 0
    df_resultado['Stock_Real'] = 0
    
    # Crear clave de artículo en el DataFrame de pedidos si no existe
    if 'Clave_Articulo' not in df_resultado.columns:
        if 'Codigo_Articulo' in df_resultado.columns:
            df_resultado['Clave_Articulo'] = (
                df_resultado['Codigo_Articulo'].astype(str) + '|' +
                df_resultado['Talla'].astype(str) + '|' +
                df_resultado['Color'].astype(str)
            )
    
    # Fusionar unidades calculadas de la semana anterior
    if df_ventas_objetivo_anterior is not None and len(df_ventas_objetivo_anterior) > 0:
        # Seleccionar solo las columnas necesarias y eliminar la columna original para evitar conflictos
        df_anterior = df_ventas_objetivo_anterior[['Clave_Articulo', 'Unidades_Calculadas_Semana_Pasada']].copy()
        df_anterior = df_anterior.rename(columns={'Unidades_Calculadas_Semana_Pasada': 'Unidades_Pasada_Merge'})
        
        # Eliminar la columna original de df_resultado antes del merge para evitar conflictos
        if 'Unidades_Calculadas_Semana_Pasada' in df_resultado.columns:
            df_resultado = df_resultado.drop(columns=['Unidades_Calculadas_Semana_Pasada'])
        
        df_resultado = df_resultado.merge(
            df_anterior,
            on='Clave_Articulo',
            how='left'
        )
        # Renombrar la columna de vuelta y llenar NaN con 0, convertir a entero
        if 'Unidades_Pasada_Merge' in df_resultado.columns:
            df_resultado = df_resultado.rename(columns={'Unidades_Pasada_Merge': 'Unidades_Calculadas_Semana_Pasada'})
            df_resultado['Unidades_Calculadas_Semana_Pasada'] = pd.to_numeric(
                df_resultado['Unidades_Calculadas_Semana_Pasada'], errors='coerce'
            ).fillna(0).astype(int)
        else:
            df_resultado['Unidades_Calculadas_Semana_Pasada'] = 0
    
    # Fusionar ventas reales
    if df_ventas_reales is not None and len(df_ventas_reales) > 0:
        # Normalizar el DataFrame de ventas reales
        df_ventas = df_ventas_reales.copy()
        
        # Crear clave de artículo
        if 'Codigo_Articulo' in df_ventas.columns:
            if 'Talla' in df_ventas.columns and 'Color' in df_ventas.columns:
                df_ventas['Clave_Articulo'] = (
                    df_ventas['Codigo_Articulo'].astype(str) + '|' +
                    df_ventas['Talla'].astype(str) + '|' +
                    df_ventas['Color'].astype(str)
                )
            else:
                df_ventas['Clave_Articulo'] = df_ventas['Codigo_Articulo'].astype(str)
            
            # Buscar columna de unidades vendidas
            col_unidades = encontrar_columna(list(df_ventas.columns), 'unidadesvendidas')
            if col_unidades is None:
                col_unidades = encontrar_columna(list(df_ventas.columns), 'ventas')
            if col_unidades is None:
                col_unidades = encontrar_columna(list(df_ventas.columns), 'unidades')
            
            if col_unidades:
                df_ventas = df_ventas[['Clave_Articulo', col_unidades]].copy()
                df_ventas.rename(columns={col_unidades: 'Ventas_Reales_Tmp'}, inplace=True)
                df_ventas['Ventas_Reales_Tmp'] = pd.to_numeric(df_ventas['Ventas_Reales_Tmp'], errors='coerce').fillna(0)
                
                df_resultado = df_resultado.merge(
                    df_ventas[['Clave_Articulo', 'Ventas_Reales_Tmp']],
                    on='Clave_Articulo',
                    how='left'
                )
                df_resultado['Ventas_Reales'] = df_resultado['Ventas_Reales_Tmp'].fillna(0).astype(int)
                df_resultado.drop(columns=['Ventas_Reales_Tmp'], inplace=True)
    
    # Fusionar stock actual
    if df_stock_actual is not None and len(df_stock_actual) > 0:
        df_stock = df_stock_actual.copy()
        
        # Crear clave de artículo
        if 'Codigo_Articulo' in df_stock.columns:
            if 'Talla' in df_stock.columns and 'Color' in df_stock.columns:
                df_stock['Clave_Articulo'] = (
                    df_stock['Codigo_Articulo'].astype(str) + '|' +
                    df_stock['Talla'].astype(str) + '|' +
                    df_stock['Color'].astype(str)
                )
            else:
                df_stock['Clave_Articulo'] = df_stock['Codigo_Articulo'].astype(str)
            
            # Buscar columna de stock
            col_stock = encontrar_columna(list(df_stock.columns), 'stock')
            if col_stock is None:
                col_stock = encontrar_columna(list(df_stock.columns), 'stockfisico')
            
            if col_stock:
                df_stock = df_stock[['Clave_Articulo', col_stock]].copy()
                df_stock.rename(columns={col_stock: 'Stock_Real_Tmp'}, inplace=True)
                df_stock['Stock_Real_Tmp'] = pd.to_numeric(df_stock['Stock_Real_Tmp'], errors='coerce').fillna(0).astype(int)
                
                df_resultado = df_resultado.merge(
                    df_stock[['Clave_Articulo', 'Stock_Real_Tmp']],
                    on='Clave_Articulo',
                    how='left'
                )
                df_resultado['Stock_Real'] = df_resultado['Stock_Real_Tmp'].fillna(0).astype(int)
                df_resultado.drop(columns=['Stock_Real_Tmp'], inplace=True)
    
    logger.info(f"Datos de tendencia fusionados: {len(df_resultado)} registros")
    logger.info(f"  - Unidades_Calculadas_Semana_Pasada: {df_resultado['Unidades_Calculadas_Semana_Pasada'].sum()}")
    logger.info(f"  - Ventas_Reales: {df_resultado['Ventas_Reales'].sum()}")
    logger.info(f"  - Stock_Real: {df_resultado['Stock_Real'].sum()}")

    return df_resultado
