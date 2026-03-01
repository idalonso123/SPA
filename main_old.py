#!/usr/bin/env python3
"""
Sistema de Pedidos de Compra - Viveverde V2

Autor: Sistema de Pedidos Viveverde V2
Fecha: 2026-02-05 (Actualizado con correcciones de bugs de email)
"""

import sys
import os
import json
import logging
import argparse
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, Tuple, List

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.paths import INPUT_DIR, OUTPUT_DIR, PEDIDOS_SEMANALES_DIR, RESUMENES_DIR
from src.data_loader import DataLoader
from src.state_manager import StateManager
from src.forecast_engine import ForecastEngine
from src.order_generator import OrderGenerator
from src.scheduler_service import SchedulerService, EstadoEjecucion

from src.correction_data_loader import CorrectionDataLoader
from src.correction_engine import CorrectionEngine, crear_correction_engine
from src.correction_data_loader import (
    encontrar_archivo_semana_anterior,
    leer_archivo_ventas_reales,
    leer_archivo_ventas_semana,
    leer_archivo_stock_actual,
    normalizar_datos_historicos,
    fusionar_datos_tendencia
)

from src.email_service import EmailService, crear_email_service
from src.alert_service import (
    iniciar_sistema_alertas,
    crear_alert_service,
    AlertLoggingHandler,
    configurar_excepthook
)

import pandas as pd

# Variable global para el logger
logger = None

def configurar_logging(nivel: int = logging.INFO, log_file: Optional[str] = None) -> logging.Logger:
    global logger
    formato = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    logger = logging.getLogger()
    logger.setLevel(nivel)
    logger.handlers = []
    
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formato)
    logger.addHandler(console_handler)
    
    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
        file_handler.setFormatter(formato)
        logger.addHandler(file_handler)
    
    return logger

def cargar_configuracion(ruta: str = 'config/config.json') -> Optional[Dict[str, Any]]:
    try:
        dir_base = os.path.dirname(os.path.abspath(__file__))
        ruta_completa = os.path.join(dir_base, ruta)
        
        if not os.path.exists(ruta_completa):
            print(f"ERROR: No se encontró el archivo de configuración: {ruta_completa}")
            return None
        
        with open(ruta_completa, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print(f"Configuración cargada desde: {ruta}")
        return config
        
    except Exception as e:
        print(f"ERROR al cargar configuración: {str(e)}")
        return None

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
    
    Útil cuando quieres mantener la estructura de palabras pero ignorar detalles.
    
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

def verificar_archivos_correccion(config: Dict[str, Any], semana: int) -> Dict[str, bool]:
    # Usar ruta centralizada por defecto, permitir override desde config
    dir_entrada_config = config.get('rutas', {}).get('directorio_entrada')
    dir_entrada = str(INPUT_DIR) if dir_entrada_config is None else dir_entrada_config
    
    archivos_correccion = config.get('archivos_correccion', {})
    
    disponibilidad = {'stock': False, 'ventas': False, 'compras': False}
    
    patrones = {
        'stock': ['SPA_stock_actual.xlsx', f'SPA_stock_semana_{semana}.xlsx'],
        'ventas': [f'SPA_ventas_semana_{semana}.xlsx', f'SPA_ventas_Semana_{semana}.xlsx', 'SPA_ventas_semana.xlsx']
    }
    
    for tipo, patrones_archivo in patrones.items():
        for patron in patrones_archivo:
            ruta = os.path.join(dir_entrada, patron)
            if os.path.exists(ruta):
                disponibilidad[tipo] = True
                logger.info(f"Archivo de {tipo} encontrado: {patron}")
                break
    
    return disponibilidad

def aplicar_correccion_pedido(
    pedido_teorico: pd.DataFrame,
    semana: int,
    config: Dict[str, Any],
    seccion: str,
    parametros_abc: Optional[Dict[str, Any]] = None
) -> Tuple[Optional[pd.DataFrame], Dict[str, Any]]:
    logger.info("\n" + "=" * 60)
    logger.info("FASE 2: APLICANDO CORRECCIÓN AL PEDIDO")
    logger.info("=" * 60)
    
    params_correccion = config.get('parametros_correccion', {})
    if not params_correccion.get('habilitar_correccion', True):
        logger.info("Corrección deshabilitada en configuración. Usando pedido teórico.")
        return pedido_teorico.copy(), {'correccion_aplicada': False}
    
    disponibilidad = verificar_archivos_correccion(config, semana)
    
    if not any(disponibilidad.values()):
        logger.warning("No se encontraron archivos de corrección. Usando pedido teórico.")
        return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'sin_archivos'}
    
    logger.info(f"Archivos de corrección disponibles: {disponibilidad}")
    
    try:
        correction_loader = CorrectionDataLoader(config)
        datos_correccion = correction_loader.cargar_datos_correccion(semana)
        
        datos_cargados = sum(1 for v in datos_correccion.values() if v is not None)
        if datos_cargados == 0:
            logger.warning("No se pudieron cargar datos de corrección. Usando pedido teórico.")
            return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'sin_datos'}
        
        pedido_fusionado = correction_loader.merge_con_pedido_teorico(
            pedido_teorico, datos_correccion
        )
        
        config_abc = {
            'pesos_categoria': config.get('parametros', {}).get('pesos_categoria', {})
        }
        
        engine = crear_correction_engine(
            config_abc=config_abc
        )
        
        pedido_corregido = engine.aplicar_correccion_dataframe(
            pedido_fusionado,
            columna_pedido='Pedido_Corregido_Stock',
            columna_stock_minimo='Stock_Minimo_Objetivo',
            columna_stock_real='Stock_Fisico',
            columna_categoria='Categoria',
            columna_ventas_reales='Ventas_Reales',
            columna_ventas_objetivo='Ventas_Objetivo',
            columna_compras_reales='Unidades_Recibidas',
            columna_compras_sugeridas='Pedido_Corregido_Stock',
            seccion=seccion
        )
        
        metricas = engine.calcular_metricas_correccion(
            pedido_corregido,
            columna_pedido_original='Pedido_Corregido_Stock',
            columna_pedido_corregido='Pedido_Corregido',
            columna_ventas_reales='Ventas_Reales',
            columna_ventas_objetivo='Ventas_Objetivo'
        )
        metricas['correccion_aplicada'] = True
        metricas['datos_cargados'] = datos_cargados
        
        # Métricas adicionales de tendencia
        if 'articulos_tendencia' in metricas:
            metricas['articulos_tendencia'] = metricas.get('articulos_tendencia', 0)
            metricas['incremento_tendencia_total'] = metricas.get('incremento_tendencia_total', 0)
        
        alertas = engine.generar_alertas(pedido_corregido)
        if alertas:
            metricas['alertas'] = alertas
            # Las alertas se generan internamente pero no se envían por email
            # No necesitamos logging adicional aquí
        
        logger.info("\nRESUMEN DE CORRECCIÓN:")
        logger.info(f"  Artículos corregidos: {metricas['articulos_corregidos']}/{metricas['total_articulos']}")
        logger.info(f"  Porcentaje corregido: {metricas['porcentaje_corregidos']:.1f}%")
        logger.info(f"  Diferencia unidades: {int(metricas['diferencia_unidades']):+d}")
        logger.info(f"  Porcentaje cambio: {metricas['porcentaje_cambio']:+.1f}%")
        
        # Mostrar métricas de tendencia de ventas (NUEVA VARIABLE)
        if 'articulos_tendencia' in metricas and metricas['articulos_tendencia'] > 0:
            logger.info(f"\n  CORRECCIÓN POR TENDENCIA DE VENTAS:")
            logger.info(f"    Artículos con tendencia: {metricas['articulos_tendencia']}")
            logger.info(f"    Incremento total aplicado: {metricas['incremento_tendencia_total']} unidades")
            logger.info(f"    Incremento promedio: {metricas.get('incremento_tendencia_promedio', 0):.1f} unidades")
        
        return pedido_corregido, metricas
        
    except Exception as e:
        logger.error(f"Error al aplicar corrección: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'error', 'error': str(e)}

def generar_archivo_pedido_corregido(
    pedido_corregido: pd.DataFrame,
    semana: int,
    seccion: str,
    parametros_seccion: Dict[str, Any],
    config: Dict[str, Any],
    order_generator: OrderGenerator
) -> Optional[str]:
    try:
        from datetime import datetime, timedelta
        fecha_base = datetime.now()
        
        dia_semana = fecha_base.weekday()
        dias_hasta_lunes = (7 - dia_semana) % 7
        fecha_lunes = fecha_base + timedelta(days=dias_hasta_lunes + (7 * ((semana - fecha_base.isocalendar()[1]) % 52)))
        
        if fecha_lunes < datetime.now():
            fecha_lunes = datetime.now()
        
        fecha_lunes_str = fecha_lunes.strftime('%Y-%m-%d')
        
        # Usar ruta centralizada por defecto, permitir override desde config
        dir_salida_config = config.get('rutas', {}).get('directorio_salida')
        if dir_salida_config:
            dir_salida = dir_salida_config
        else:
            dir_salida = str(PEDIDOS_SEMANALES_DIR)
        
        nombre_archivo = f"Pedido_Semana_{semana}_{fecha_lunes_str}_{seccion}_CORREGIDO.xlsx"
        ruta_archivo = os.path.join(dir_salida, nombre_archivo)
        
        df_exportar = pedido_corregido.copy()
        
        renombrar = {
            'Pedido_Corregido_Stock': 'Pedido_Teorico',
            'Pedido_Corregido': 'Pedido_Final',
            'Stock_Minimo_Objetivo': 'Stock_Minimo',
            'Stock_Fisico': 'Stock_Real',
            'Unidades_Vendidas': 'Ventas_Reales',
            'Unidades_Recibidas': 'Compras_Recibidas',
            'Diferencia_Stock': 'Ajuste_Stock',
            'Razon_Correccion': 'Correccion_Aplicada'
        }
        
        for col_vieja, col_nueva in renombrar.items():
            if col_vieja in df_exportar.columns:
                df_exportar.rename(columns={col_vieja: col_nueva}, inplace=True)
        
        # Columnas de tendencia (NUEVA VARIABLE) - mantener nombres originales
        columnas_tendencia = ['Porcentaje_Consumido_Stock', 'Incremento_Tendencia', 'Tendencia_Aplicada']
        columnas_finales = [
            'Código artículo', 'Nombre artículo', 'Talla', 'Color', 'Categoria',
            'Pedido_Teorico', 'Stock_Minimo', 'Stock_Real', 'Ajuste_Stock',
            'Pedido_Final', 'Correccion_Aplicada', 'Escenario',
            'Ventas_Reales', 'Ventas_Objetivo', 'Compras_Recibidas',
            'PVP', 'Coste', 'Proveedor', 'Unidades_ABC', 'Ventas_Objetivo',
            # Nuevas columnas de tendencia
            'Porcentaje_Consumido_Stock', 'Incremento_Tendencia', 'Tendencia_Aplicada'
        ]
        
        columnas_finales = [col for col in columnas_finales if col in df_exportar.columns]
        df_exportar = df_exportar[columnas_finales]
        
        df_exportar.to_excel(ruta_archivo, index=False, sheet_name=seccion.capitalize())
        
        logger.info(f"Archivo de pedido corregido generado: {nombre_archivo}")
        return ruta_archivo
        
    except Exception as e:
        logger.error(f"Error al generar archivo corregido: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None

def enviar_emails_pedidos(
    semana: int,
    config: Dict[str, Any],
    archivos_por_seccion: Dict[str, List[str]]
) -> Tuple[Dict[str, Any], Any]:
    """
    Envía los emails de pedidos a los responsables de cada sección.

    Args:
        semana (int): Número de semana
        config (Dict[str, Any]): Configuración del sistema
        archivos_por_seccion (Dict[str, List[str]]): Archivos generados por sección

    Returns:
        Tuple[Dict[str, Any], Any]: Tupla con (resultado del envío, email_service)
    """
    logger.info("\n" + "=" * 60)
    logger.info("ENVÍO DE EMAILS A RESPONSABLES")
    logger.info("=" * 60)

    email_config = config.get('email', {})
    if not email_config.get('habilitar_envio', True):
        logger.info("Envío de emails deshabilitado en configuración.")
        return {'exito': False, 'razon': 'deshabilitado'}, None

    try:
        email_service = crear_email_service(config)

        verificacion = email_service.verificar_configuracion()
        if not verificacion['valido']:
            logger.warning("Problemas en la configuración de email:")
            for problema in verificacion['problemas']:
                logger.warning(f"  - {problema}")

            if any('EMAIL_PASSWORD' in p for p in verificacion['problemas']):
                logger.error("No se puede enviar emails sin configurar la variable EMAIL_PASSWORD")
                return {'exito': False, 'razon': 'sin_password'}, None

        resultados = {}
        emails_enviados = 0
        emails_fallidos = 0

        for seccion, archivos in archivos_por_seccion.items():
            if not archivos:
                logger.info(f"Sin archivos para la sección {seccion}. Saltando.")
                continue

            logger.info(f"\nEnviando email para sección: {seccion}")
            logger.info(f"Archivos: {archivos}")

            resultado = email_service.enviar_pedido_por_seccion(
                semana=semana,
                seccion=seccion,
                archivos=archivos
            )

            resultados[seccion] = resultado

            if resultado.get('enviado', False):
                emails_enviados += 1
                logger.info(f"✓ Email enviado exitosamente a {seccion}")
            else:
                emails_fallidos += 1
                logger.error(f"✗ Error al enviar email a {seccion}: {resultado.get('error', 'Error desconocido')}")

        logger.info("\n" + "=" * 60)
        logger.info("RESUMEN DE ENVÍO DE EMAILS")
        logger.info("=" * 60)
        logger.info(f"Emails enviados exitosamente: {emails_enviados}")
        logger.info(f"Emails fallidos: {emails_fallidos}")
        logger.info(f"Secciones procesadas: {len(resultados)}")

        resultado_final = {
            'exito': emails_enviados > 0,
            'emails_enviados': emails_enviados,
            'emails_fallidos': emails_fallidos,
            'resultados': resultados
        }

        return resultado_final, email_service

    except Exception as e:
        logger.error(f"Error al enviar emails: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return {'exito': False, 'error': str(e)}, None

def agrupar_archivos_por_seccion(
    archivos_generados: List[str],
    config: Dict[str, Any]
) -> Dict[str, List[str]]:
    """
    CORRECCIÓN: Esta función ahora maneja correctamente secciones con guiones bajos
    y ambos formatos de fecha (YYYY-MM-DD y DDMMYYYY).
    
    El problema original era que para archivos como:
    - Pedido_Semana_14_2026-02-03_mascotas_vivo.xlsx
    - Pedido_Semana_14_03022026_mascotas_vivo.xlsx
    - Pedido_Semana_14_2026-02-03_deco_interior.xlsx
    
    El código dividía por '_' y tomaba solo la última parte, convirtiendo:
    - 'mascotas_vivo' en 'vivo'
    - 'deco_interior' en 'interior'
    
    La solución es reconstruir la sección uniendo todas las partes después de la fecha.
    
    Formatos de fecha soportados:
    - YYYY-MM-DD (con guiones): 2026-02-03
    - DDMMYYYY (sin guiones): 03022026
    """
    archivos_por_seccion = {}
    
    # Usar ruta centralizada por defecto, permitir override desde config
    dir_salida_config = config.get('rutas', {}).get('directorio_salida')
    if dir_salida_config:
        dir_salida = dir_salida_config
    else:
        dir_salida = str(PEDIDOS_SEMANALES_DIR)

    for archivo in archivos_generados:
        if not archivo:
            continue

        nombre_archivo = os.path.basename(archivo)
        if 'RESUMEN' in nombre_archivo.upper():
            continue

        nombre_sin_extension = nombre_archivo.replace('.xlsx', '')
        
        partes = nombre_sin_extension.split('_')
        
        if len(partes) >= 4:
            # Detectar formato de fecha en partes[3]
            parte_fecha = partes[3]
            
            # Formato 1: YYYY-MM-DD (contiene '-')
            if '-' in parte_fecha:
                # La sección es todo lo que viene después de la fecha (índice 4 en adelante)
                seccion = '_'.join(partes[4:])
            # Formato 2: DDMMYYYY (8 dígitos numéricos)
            elif parte_fecha.isdigit() and len(parte_fecha) == 8:
                # La sección es todo lo que viene después de la fecha (índice 4 en adelante)
                seccion = '_'.join(partes[4:])
            else:
                # Formato inesperado: usar la última parte como sección
                logger.warning(f"Formato de fecha no reconocido en archivo: {nombre_archivo}")
                logger.warning(f"Parte de fecha esperada: '{parte_fecha}'")
                seccion = partes[-1]

            if seccion not in archivos_por_seccion:
                archivos_por_seccion[seccion] = []

            archivos_por_seccion[seccion].append(archivo)

    logger.debug(f"[DEBUG] Archivos agrupados por sección: {archivos_por_seccion}")

    return archivos_por_seccion

def procesar_pedido_semana(
    semana: int, 
    config: Dict[str, Any], 
    state_manager: StateManager,
    forzar: bool = False,
    aplicar_correccion: bool = True,
    enviar_email: bool = True,
    alert_service=None
) -> Tuple[bool, Optional[str], int, float, Dict[str, Any], Dict[str, Any]]:
    logger.info("=" * 70)
    logger.info(f"PROCESANDO PEDIDO PARA SEMANA {semana}")
    logger.info("=" * 70)
    
    if aplicar_correccion:
        logger.info("MODO: FASE 1 (Forecast) + FASE 2 (Corrección)")
    else:
        logger.info("MODO: Solo FASE 1 (Forecast) - Corrección deshabilitada")
    
    data_loader = DataLoader(config)
    forecast_engine = ForecastEngine(config)
    order_generator = OrderGenerator(config)
    scheduler = SchedulerService(config)
    
    fecha_lunes, fecha_domingo, fecha_archivo = scheduler.calcular_fechas_semana_pedido(semana)
    logger.info(f"Período de la semana: {fecha_lunes} al {fecha_domingo}")
    
    stock_acumulado = state_manager.obtener_stock_acumulado()
    logger.info(f"Stock acumulado cargado: {len(stock_acumulado)} artículos")
    
    secciones = config.get('secciones_activas', [])
    pedidos_totales = {}
    pedidos_corregidos = {}
    datos_semanales = {}
    articulos_totales = 0
    importe_total = 0.0
    metricas_correccion_total = {}
    
    archivos_generados = []
    
    # ============================================================================
    # CARGA DE DATOS PARA CÁLCULO DE TENDENCIA DE VENTAS
    # ============================================================================
    # Estos datos se usan para calcular la tendencia comparando ventas reales con objetivo
    
    # Determinar directorios
    dir_base = os.path.dirname(os.path.abspath(__file__))
    dir_entrada_config = config.get('rutas', {}).get('directorio_entrada')
    # Usar ruta centralizada por defecto si no hay override en config
    if dir_entrada_config is None:
        dir_entrada = str(INPUT_DIR)
    else:
        dir_entrada = os.path.join(dir_base, dir_entrada_config)
    
    # Usar ruta centralizada por defecto, permitir override desde config
    dir_salida_config = config.get('rutas', {}).get('directorio_salida')
    if dir_salida_config is None:
        dir_salida = str(PEDIDOS_SEMANALES_DIR)
    else:
        dir_salida = os.path.join(dir_base, dir_salida_config)
    
    # Cargar archivo de ventas de semana (SPA_ventas_semana.xlsx) - Contiene las ventas reales de la semana anterior
    df_ventas_reales, ventas_reales_existe = leer_archivo_ventas_semana(dir_entrada)
    
    # Cargar archivo de stock actual (SPA_stock_actual.xlsx)
    df_stock_actual = leer_archivo_stock_actual(dir_entrada)
    
    for seccion in secciones:
        logger.info(f"\n{'=' * 50}")
        logger.info(f"SECCION: {seccion.upper()}")
        logger.info(f"{'=' * 50}")
        
        # Actualizar contexto de alertas con la sección actual
        try:
            if 'alert_service' in dir():
                fecha_str = datetime.now().strftime('%Y-%m-%d')
                alert_service.establecer_contexto(seccion=seccion, fecha=fecha_str, area='Pedidos')
                logger.debug(f"Contexto de alertas actualizado: Sección={seccion}")
        except Exception as e:
            logger.debug(f"No se pudo actualizar contexto de alertas: {e}")
        
        try:
            abc_df, ventas_df, costes_df = data_loader.leer_datos_seccion(seccion, semana)
            
            logger.debug(f"[DEBUG] abc_df: {len(abc_df) if abc_df is not None else 0} registros")
            logger.debug(f"[DEBUG] ventas_df: {len(ventas_df) if ventas_df is not None else 0} registros")
            logger.debug(f"[DEBUG] costes_df: {len(costes_df) if costes_df is not None else 0} registros")
            
            if abc_df is None or ventas_df is None or costes_df is None:
                logger.error(f"No se pudieron leer los datos para la seccion '{seccion}'")
                continue
            
            if 'Semana' not in ventas_df.columns:
                if 'Fecha' in ventas_df.columns:
                    ventas_df['Fecha'] = pd.to_datetime(ventas_df['Fecha'], errors='coerce')
                    ventas_df['Semana'] = ventas_df['Fecha'].apply(
                        lambda x: x.isocalendar()[1] if pd.notna(x) else None
                    )
                else:
                    logger.warning(f"No hay columna 'Fecha' ni 'Semana' en ventas de '{seccion}'")
                    continue
            
            datos_semana = ventas_df[ventas_df['Semana'] == semana]
            
            if len(datos_semana) == 0:
                logger.warning(f"No hay datos de ventas para la semana {semana} en '{seccion}'")
                continue
            
            logger.info(f"Datos de ventas: {len(datos_semana)} registros")
            
            parametros_seccion = {
                'objetivos_semanales': config.get('secciones', {}).get(seccion, {}).get('objetivos_semanales', {}),
                'objetivo_crecimiento': config.get('parametros', {}).get('objetivo_crecimiento', 0.05),
                'stock_minimo_porcentaje': config.get('parametros', {}).get('stock_minimo_porcentaje', 0.30),
                'festivos': config.get('festivos', {})
            }
            
            pedidos = forecast_engine.calcular_pedido_semana(
                semana, datos_semana, abc_df, costes_df, seccion
            )
            
            if len(pedidos) == 0:
                logger.warning(f"No se generaron pedidos para '{seccion}'")
                continue
            
            # ============================================================================
            # FUSIÓN DE DATOS PARA CÁLCULO DE TENDENCIA (ANTES DE APLICAR STOCK MÍNIMO)
            # ============================================================================
            # Añadir columnas: Unidades_Calculadas_Semana_Pasada, Ventas_Reales, Stock_Real
            # Buscar archivo de pedido de la semana anterior para esta sección
            archivo_semana_anterior = encontrar_archivo_semana_anterior(dir_salida, semana, seccion)
            df_ventas_objetivo_anterior = None
            if archivo_semana_anterior:
                try:
                    # IMPORTANTE: El archivo Excel tiene los encabezados en la fila 2 (índice 1)
                    # La primera fila es un índice. Sin header=1, pandas lee índices (1,2,3...) como columnas
                    df_pedido_anterior = pd.read_excel(archivo_semana_anterior, header=1)
                    df_ventas_objetivo_anterior = normalizar_datos_historicos(df_pedido_anterior)
                    logger.info(f"Cargados datos de la semana anterior ({seccion}): {len(df_ventas_objetivo_anterior)} registros")
                except Exception as e:
                    logger.warning(f"No se pudo leer el archivo de la semana anterior para '{seccion}': {str(e)}")
            
            pedidos = fusionar_datos_tendencia(
                pedidos,
                df_ventas_reales,
                df_stock_actual,
                df_ventas_objetivo_anterior
            )
            
            # ============================================================================
            # EXTRAER DICCIONARIOS PARA APLICAR STOCK MÍNIMO
            # ============================================================================
            # Crear diccionarios de stock real y ventas a partir de los datos fusionados
            stock_real_dict = {}
            ventas_reales_dict = {}
            ventas_objetivo_dict = {}

            for idx, row in pedidos.iterrows():
                # Normalizar Codigo_Articulo igual que en fusionar_datos_tendencia
                codigo_raw = row.get('Codigo_Articulo', '')
                codigo = str(codigo_raw).replace('.0', '', 1).strip() if pd.notna(codigo_raw) else ''
                clave = f"{codigo}|{row.get('Talla', '')}|{row.get('Color', '')}"
                stock_real_dict[clave] = row.get('Stock_Real', 0)
                ventas_reales_dict[clave] = row.get('Ventas_Reales', 0)
                ventas_objetivo_dict[clave] = row.get('Unidades_Calculadas_Semana_Pasada', 0)
            
            # ============================================================================
            # APLICAR STOCK MÍNIMO Y CALCULAR PEDIDO FINAL
            # ============================================================================
            pedidos, nuevo_stock, ajustes = forecast_engine.aplicar_stock_minimo(
                pedidos, semana, stock_acumulado, stock_real_dict, ventas_reales_dict, ventas_objetivo_dict
            )
            
            stock_acumulado.update(nuevo_stock)
            
            if aplicar_correccion:
                pedidos_corregido, metricas = aplicar_correccion_pedido(
                    pedidos.copy(), semana, config, seccion,
                    parametros_abc=config.get('parametros', {})
                )

                if metricas.get('correccion_aplicada', False):
                    metricas_correccion_total[seccion] = metricas

                    # Ya no generamos archivo separado con "_CORREGIDO"
                    # El archivo corregido se genera en la línea 835 con el formato correcto
                    # usando order_generator.generar_archivo_pedido()

                    pedidos_final = pedidos_corregido
                    pedidos_corregidos[seccion] = pedidos_corregido
                else:
                    pedidos_final = pedidos
                    logger.info("Usando pedido teórico (sin corrección)")
            else:
                pedidos_final = pedidos

            archivo = order_generator.generar_archivo_pedido(pedidos_final, semana, seccion, parametros_seccion)
            
            if archivo:
                archivos_generados.append(archivo)
                
                if 'Pedido_Final' in pedidos_final.columns:
                    pedidos_validos = pedidos_final[pedidos_final['Pedido_Final'] > 0]
                    articulos = len(pedidos_validos)
                    importe = pedidos_validos['Ventas_Objetivo'].sum()
                else:
                    pedidos_validos = pedidos_final[pedidos_final['Pedido_Corregido_Stock'] > 0]
                    articulos = len(pedidos_validos)
                    importe = pedidos_validos['Ventas_Objetivo'].sum()
                
                articulos_totales += articulos
                importe_total += importe
                
                logger.info(f"Archivo generado: {archivo}")
                logger.info(f"  Articulos: {articulos}")
                logger.info(f"  Importe: {importe:.2f}€")
            else:
                logger.warning(f"No se generó archivo para '{seccion}'")
            
            pedidos_totales[seccion] = pedidos_final
            datos_semanales[seccion] = datos_semana
            
        except Exception as e:
            logger.error(f"Error procesando seccion '{seccion}': {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            continue
    
    if stock_acumulado:
        state_manager.actualizar_stock_acumulado(stock_acumulado)
    
    # CORRECCIÓN: Generar archivo de resumen para CADA SECCIÓN y uno consolidado
    if pedidos_totales:
        resumen_data = []
        for seccion, pedidos in pedidos_totales.items():
            if len(pedidos) > 0:
                # CORRECCIÓN: Pasar la sección correcta a generar_resumen_pedido
                resumen_seccion = forecast_engine.generar_resumen_pedido(
                    pedidos, semana, datos_semanales.get(seccion, pd.DataFrame()), seccion
                )
                if resumen_seccion:
                    resumen_data.append(resumen_seccion)
        
        if resumen_data:
            resumen_df = pd.DataFrame(resumen_data)
            # CORRECCIÓN: Generar un resumen consolidado con TODAS las secciones
            archivo_resumen = order_generator.generar_resumen_excel(resumen_df, 'CONSOLIDADO')
            if archivo_resumen:
                archivos_generados.append(archivo_resumen)
    
    archivo_principal = archivos_generados[0] if archivos_generados else None
    
    notas_correccion = ""
    if metricas_correccion_total:
        articulos_corregidos_total = sum(
            m.get('articulos_corregidos', 0) for m in metricas_correccion_total.values()
        )
        notas_correccion = f" - {articulos_corregidos_total} artículos corregidos en FASE 2"
    
    state_manager.registrar_ejecucion(
        semana=semana,
        archivo_generado=archivo_principal or "Sin archivo",
        articulos=articulos_totales,
        importe=importe_total,
        exitosa=len(archivos_generados) > 0,
        notas=f"Procesadas {len(secciones)} secciones{notas_correccion}"
    )
    
    resultado_email = {'exito': False, 'razon': 'no_enviado'}
    resultado_resumen_gestion = {'enviado': False, 'razon': 'no_enviado'}
    
    if enviar_email and archivos_generados:
        logger.info("\n" + "=" * 60)
        logger.info("PREPARANDO ENVÍO DE EMAILS")
        logger.info("=" * 60)
        
        # CORRECCIÓN: Usar la función grouping corregida
        archivos_por_seccion = agrupar_archivos_por_seccion(archivos_generados, config)

        resultado_email, email_service = enviar_emails_pedidos(semana, config, archivos_por_seccion)

        # Enviar resumen a los responsables de gestión (Sandra, Ivan, Pedro)
        if email_service:
            logger.info("\n" + "=" * 60)
            logger.info("PREPARANDO ENVÍO DE RESUMEN A RESPONSABLES DE GESTIÓN")
            logger.info("=" * 60)
            
            # Buscar el archivo de resumen consolidado (excluir archivos antiguos)
            archivo_resumen = None
            for archivo in archivos_generados:
                if archivo and 'RESUMEN' in Path(archivo).name.upper():
                    # Excluir archivos antiguos (que empiezan con número o contienen 'old')
                    nombre = Path(archivo).name
                    if not nombre[0].isdigit() and 'old' not in nombre.lower():
                        archivo_resumen = archivo
                        break
            
            if archivo_resumen:
                logger.info(f"Archivo de resumen encontrado: {Path(archivo_resumen).name}")
                resultado_resumen_gestion = email_service.enviar_resumen_gestion(semana, archivo_resumen)
            else:
                logger.warning("No se encontró archivo de resumen consolidado")
                logger.info("Omitiendo envío de resumen a responsables de gestión")
    
    logger.info("\n" + "=" * 70)
    logger.info("RESUMEN DE EJECUCION")
    logger.info("=" * 70)
    logger.info(f"Semana procesada: {semana}")
    logger.info(f"Archivos generados: {len(archivos_generados)}")
    logger.info(f"Total articulos: {articulos_totales}")
    logger.info(f"Total importe: {importe_total:.2f}€")
    
    if metricas_correccion_total:
        logger.info("\nMÉTRICAS DE CORRECCIÓN (FASE 2):")
        logger.info("-" * 40)
        articulos_corregidos_total = 0
        diferencia_total = 0
        articulos_tendencia_total = 0
        incremento_tendencia_total = 0
        
        for seccion, metricas in metricas_correccion_total.items():
            articulos_corregidos_total += metricas.get('articulos_corregidos', 0)
            diferencia_total += metricas.get('diferencia_unidades', 0)
            articulos_tendencia_total += metricas.get('articulos_tendencia', 0)
            incremento_tendencia_total += metricas.get('incremento_tendencia_total', 0)
            logger.info(f"  {seccion}: {metricas.get('articulos_corregidos', 0)} artículos corregidos")
        
        logger.info(f"  TOTAL: {articulos_corregidos_total} artículos corregidos")
        logger.info(f"  Diferencia neta: {int(diferencia_total):+d} unidades")
        
        # Mostrar métricas de tendencia (NUEVA VARIABLE)
        if articulos_tendencia_total > 0:
            logger.info(f"\nCORRECCIÓN POR TENDENCIA DE VENTAS:")
            logger.info(f"  Artículos con tendencia detectada: {articulos_tendencia_total}")
            logger.info(f"  Incremento total aplicado: {incremento_tendencia_total} unidades")
    
    # Enviar alerta si el envío de emails a encargados falla
    if not resultado_email.get('exito') and resultado_email.get('emails_fallidos', 0) > 0:
        logger.warning(f"\n✗ EMAILS FALLIDOS: {resultado_email.get('emails_fallidos', 0)}")
        # Intentar enviar alerta de error de email
        try:
            from src.alert_service import crear_alert_service
            alert_service = crear_alert_service(config)
            # Agrupar los errores por sección
            errores_secciones = []
            for seccion, res in resultado_email.get('resultados', {}).items():
                if not res.get('exito', False):
                    errores_secciones.append(f"{seccion}: {res.get('error', 'Error desconocido')}")

            if errores_secciones:
                alert_service.alerta_error_envio(
                    destinatario='encargados_secciones',
                    asunto=f"Error en envío de pedidos - Semana {semana}",
                    tipo_error='Envío fallido a encargado',
                    detalles='; '.join(errores_secciones)
                )
                logger.info("Alerta de error de email enviada")
        except Exception as e:
            logger.warning(f"No se pudo enviar alerta de error de email: {e}")

    if resultado_email.get('exito'):
        logger.info(f"\n✓ EMAILS ENVIADOS A ENCARGADOS: {resultado_email.get('emails_enviados', 0)}")
    elif resultado_email.get('razon') == 'deshabilitado':
        logger.info("\nEmails deshabilitados")
    elif resultado_email.get('razon') == 'sin_password':
        logger.warning("\n✗ No se enviaron emails (falta configurar EMAIL_PASSWORD)")
    
    # Mostrar resultado del resumen de gestión
    if resultado_resumen_gestion.get('enviado'):
        logger.info(f"\n✓ RESUMEN ENVIADO A RESPONSABLES DE GESTIÓN: {resultado_resumen_gestion.get('emails_enviados', 0)}/3")
    elif resultado_resumen_gestion.get('razon') == 'archivo_no_encontrado':
        logger.info("\nResumen de gestión no enviado (archivo no encontrado)")
    elif resultado_resumen_gestion.get('emails_fallidos', 0) > 0:
        logger.warning(f"\n✗ RESUMEN FALLIDO A RESPONSABLES: {resultado_resumen_gestion.get('emails_fallidos', 0)}/3")
        # Enviar alerta si el resumen de gestión falla
        try:
            from src.alert_service import crear_alert_service
            alert_service = crear_alert_service(config)
            alert_service.alerta_error_envio(
                destinatario='responsables_gestion',
                asunto=f"Error en envío de resumen - Semana {semana}",
                tipo_error='Resumen de gestión no enviado',
                detalles=f"Fallidos: {resultado_resumen_gestion.get('emails_fallidos', 0)}/3"
            )
            logger.info("Alerta de error de resumen enviada")
        except Exception as e:
            logger.warning(f"No se pudo enviar alerta de error de resumen: {e}")
    
    logger.info("=" * 70)
    
    return len(archivos_generados) > 0, archivo_principal, articulos_totales, importe_total, metricas_correccion_total, resultado_email, resultado_resumen_gestion

def main():
    parser = argparse.ArgumentParser(
        description='Sistema de Generación de Pedidos de Compra - Viveverde V2 (FASE 1 + FASE 2 + Email)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python main.py                      # Ejecución normal (jueves 21:50)
  python main.py --semana 15          # Forzar semana específica
  python main.py --continuo           # Modo continuo (se ejecuta y espera la siguiente semana)
  python main.py --status             # Mostrar estado del sistema
  python main.py --reset              # Resetear estado del sistema
  python main.py --semana 15 --sin-correccion    # Solo FASE 1
  python main.py --semana 15 --con-correccion     # FASE 1 + FASE 2 (forzado)
  python main.py --semana 15 --sin-email          # Sin enviar emails
  python main.py --verificar-email                # Verificar configuración de email
        """
    )
    
    parser.add_argument('--semana', '-s', type=int, help='Número de semana a procesar (para pruebas)')
    parser.add_argument('--continuo', '-c', action='store_true', help='Ejecutar en modo continuo')
    parser.add_argument('--status', action='store_true', help='Mostrar estado del sistema y salir')
    parser.add_argument('--reset', action='store_true', help='Resetear el estado del sistema')
    parser.add_argument('--verbose', '-v', action='store_true', help='Activar logging detallado (DEBUG)')
    parser.add_argument('--log', type=str, default='logs/sistema.log', help='Archivo de log (default: logs/sistema.log)')
    parser.add_argument('--sin-correccion', action='store_true', help='Ejecutar solo FASE 1 (sin corrección)')
    parser.add_argument('--con-correccion', action='store_true', help='Forzar ejecución con corrección FASE 2')
    parser.add_argument('--sin-email', action='store_true', help='No enviar emails después de generar los pedidos')
    parser.add_argument('--verificar-email', action='store_true', help='Verificar la configuración de email y salir')
    
    args = parser.parse_args()
    
    nivel_log = logging.DEBUG if args.verbose else logging.INFO
    
    global logger
    logger = configurar_logging(nivel=nivel_log, log_file=args.log)
    
    logger.info("=" * 70)
    logger.info("SISTEMA DE PEDIDOS DE COMPRA - VIVEVERDE V2")
    logger.info(f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)
    
    config = cargar_configuracion()
    if config is None:
        logger.error("No se pudo cargar la configuración. Saliendo.")
        sys.exit(1)
    
    # ========================================================================
    # INTEGRACIÓN DEL SISTEMA DE ALERTAS
    # ========================================================================
    # IMPORTAR Y CONFIGURAR EL SERVICIO DE ALERTAS
    try:
        from src.alert_service import (
            iniciar_sistema_alertas,
            crear_alert_service,
            AlertLoggingHandler,
            configurar_excepthook
        )
        
        # Obtener destinatario de alertas desde configuración
        env_config = config.get('env_email', {})
        destinatario_alertas = env_config.get('destinatario_alertas', 'ivan.delgado@viveverde.es')
        
        # Crear alert service con destinatario configurable
        alert_service = crear_alert_service(config, destinatario=destinatario_alertas)
        
        # Iniciar sistema completo de alertas (logging handler + excepthook)
        iniciar_sistema_alertas(config)
        logger.info(f"Sistema de alertas automáticamente configurado (destinatario: {destinatario_alertas})")
        
    except Exception as e:
        # Si falla la configuración de alertas, continuar sin ellas pero avisar
        logger.warning(f"No se pudo inicializar el sistema de alertas: {e}")
        logger.warning("El sistema continuará sin notificaciones por email")
    
    if args.verificar_email:
        logger.info("\nVERIFICANDO CONFIGURACIÓN DE EMAIL:")
        logger.info("-" * 40)
        
        try:
            email_service = crear_email_service(config)
            verificacion = email_service.verificar_configuracion()
            
            if verificacion['valido']:
                logger.info("✓ Configuración de email válida")
            else:
                logger.warning("✗ Problemas en la configuración:")
                for problema in verificacion['problemas']:
                    logger.warning(f"  - {problema}")
            
            import os
            if os.environ.get('EMAIL_PASSWORD'):
                logger.info("✓ Variable EMAIL_PASSWORD configurada")
            else:
                logger.warning("✗ Variable EMAIL_PASSWORD NO configurada")
                logger.info("  Para configurarla, ejecuta:")
                logger.info("  export EMAIL_PASSWORD='tu_contraseña'")
            
        except Exception as e:
            logger.error(f"Error al verificar email: {str(e)}")
        
        sys.exit(0)
    
    params_correccion = config.get('parametros_correccion', {})
    correccion_habilitada = params_correccion.get('habilitar_correccion', True)
    aplicar_correccion = correccion_habilitada and not args.sin_correccion
    
    if args.con_correccion:
        aplicar_correccion = True
    
    email_config = config.get('email', {})
    email_habilitado = email_config.get('habilitar_envio', True)
    enviar_email = email_habilitado and not args.sin_email
    
    logger.info(f"Modo de ejecución: {'FASE 1 + FASE 2' if aplicar_correccion else 'Solo FASE 1'}")
    logger.info(f"Envío de emails: {'Sí' if enviar_email else 'No'}")
    
    state_manager = StateManager(config)
    state_manager.cargar_estado()
    
    if args.reset:
        logger.info("Reseteando estado del sistema...")
        state_manager.resetear_estado()
        logger.info("Estado reseteado correctamente.")
        sys.exit(0)
    
    if args.status:
        logger.info("\nESTADO DEL SISTEMA:")
        logger.info(state_manager.obtener_resumen_estado())
        
        scheduler = SchedulerService(config)
        logger.info("\n" + scheduler.simular_proxima_ejecucion())
        
        ultima = state_manager.obtener_ultima_semana_procesada()
        logger.info(f"\nÚltima semana procesada: {ultima if ultima else 'Ninguna'}")
        
        metricas = state_manager.obtener_metricas()
        logger.info(f"Métricas: {metricas}")
        
        sys.exit(0)
    
    scheduler = SchedulerService(config)
    ultima_procesada = state_manager.obtener_ultima_semana_procesada()
    
    # Ejecución directa - sin verificación de horario
    # El horario lo controla la tarea programada del sistema operativo
    if args.semana:
        semana = args.semana
        logger.info(f"Semana forzada por argumento: {semana}")
    else:
        semana, msg_semana = scheduler.calcular_semana_a_procesar(ultima_procesada)
        
        if semana is None:
            logger.info(msg_semana)
            sys.exit(0)
        
        logger.info(msg_semana)
    
    if state_manager.verificar_semana_procesada(semana) and not args.semana:
        logger.warning(f"La semana {semana} ya fue procesada anteriormente.")
        logger.info("Use --semana para forzar el reprocesamiento.")
        sys.exit(0)
    
    exito, archivo, articulos, importe, metricas_correccion, resultado_email, resultado_resumen_gestion = procesar_pedido_semana(
        semana, config, state_manager, 
        forzar=args.semana is not None,
        aplicar_correccion=aplicar_correccion,
        enviar_email=enviar_email,
        alert_service=alert_service if 'alert_service' in dir() else None
    )
    
    if exito:
        logger.info(f"\n¡PEDIDO GENERADO EXITOSAMENTE!")
        logger.info(f"Archivo principal: {archivo}")
        logger.info(f"Artículos: {articulos}")
        logger.info(f"Importe: {importe:.2f}€")
        
        if metricas_correccion:
            logger.info(f"\nFASE 2 completada: {len(metricas_correccion)} secciones corregidas")
            
            # Mostrar métricas de tendencia si existen
            if any('articulos_tendencia' in m for m in metricas_correccion.values()):
                total_tendencia = sum(m.get('articulos_tendencia', 0) for m in metricas_correccion.values())
                total_incremento = sum(m.get('incremento_tendencia_total', 0) for m in metricas_correccion.values())
                if total_tendencia > 0:
                    logger.info(f"  - Tendencia de ventas: {total_tendencia} artículos con incremento adicional")
                    logger.info(f"  - Incremento total aplicado: {total_incremento} unidades")
        
        if resultado_email.get('exito'):
            logger.info(f"\nEmails enviados a encargados: {resultado_email.get('emails_enviados', 0)}")
        
        if resultado_resumen_gestion.get('enviado'):
            logger.info(f"Resumen enviado a responsables de gestión: {resultado_resumen_gestion.get('emails_enviados', 0)}/3")
        elif resultado_resumen_gestion.get('emails_fallidos', 0) > 0:
            logger.warning(f"Resumen fallido a responsables de gestión: {resultado_resumen_gestion.get('emails_fallidos', 0)}/3")
        
        sys.exit(0)
    else:
        logger.error("\nERROR: No se pudo generar el pedido.")
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nEjecución cancelada por el usuario")
        sys.exit(0)
    except Exception as e:
        # Esta alerta es redundante ya que el excepthook debería capturarlo
        # pero por seguridad la mantenemos
        try:
            from src.alert_service import crear_alert_service
            from src.config_loader import cargar_configuracion
            config = cargar_configuracion()
            if config:
                alert_service = crear_alert_service(config)
                alert_service.alerta_excepcion(e, seccion="main")
        except:
            pass  # Si falla la alerta, al menos mostrar el error
        raise