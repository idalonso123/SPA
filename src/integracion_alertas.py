#!/usr/bin/env python3
"""
Módulo de Integración de Alertas para Scripts del Sistema SPA

Este módulo proporciona funciones de conveniencia para integrar rápidamente
el sistema de alertas (AlertService) en cualquier script del proyecto.

Autor: Sistema de Pedidos VIVEVERDE
Fecha: 2026-02-25
"""

import os
import sys
import logging
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any, Callable

# Agregar el directorio raíz al path para poder importar módulos del proyecto
DIRECTORIO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, DIRECTORIO_BASE)

# Importar módulos necesarios
try:
    from src.alert_service import crear_alert_service, iniciar_sistema_alertas, AlertService
    from src.config_loader import cargar_configuracion
    ALERT_SERVICE_DISPONIBLE = True
except ImportError as e:
    ALERT_SERVICE_DISPONIBLE = False
    print(f"ADVERTENCIA: No se pudo importar alert_service: {e}")

# Configurar logger
logger = logging.getLogger(__name__)


class IntegradorAlertas:
    """
    Clase facilitadora para integrar el sistema de alertas en cualquier script.
    
    Proporciona métodos convenientes para:
    - Inicializar el sistema de alertas
    - Capturar y reportar errores
    - Reportar advertencias
    - Establecer contexto de procesamiento
    
    Usage:
        from src.integracion_alertas import IntegradorAlertas
        
        # Inicializar al inicio del script
        integrador = IntegradorAlertas("mi_script")
        integrador.inicializar()
        
        # Usar en bloques try/except
        try:
            # código que puede fallar
            resultado = proceso_peligroso()
        except Exception as e:
            integrador.reportar_error(e, seccion="proceso_peligroso")
            raise
        
        # Para advertencias
        integrador.reportar_advertencia("Datos insuficientes", seccion="validacion")
        
        # Al final del script
        integrador.finalizar()
    """
    
    def __init__(self, nombre_script: str, destinatario: str = "ivan.delgado@viveverde.es"):
        """
        Inicializa el integrador de alertas.
        
        Args:
            nombre_script (str): Nombre del script que está usando el integrador
            destinatario (str): Email destinatario para las alertas
        """
        self.nombre_script = nombre_script
        self.destinatario = destinatario
        self.alert_service: Optional[AlertService] = None
        self.config: Optional[Dict] = None
        self.inicializado = False
        self.seccion_actual = "general"
        self.fecha_proceso = datetime.now().strftime('%Y-%m-%d')
        
    def inicializar(self, seccion: str = None, fecha: str = None) -> bool:
        """
        Inicializa el sistema de alertas.
        
        Args:
            seccion (str): Sección inicial a reportar
            fecha (str): Fecha del proceso
            
        Returns:
            bool: True si se inicializó correctamente
        """
        if not ALERT_SERVICE_DISPONIBLE:
            logger.warning("AlertService no disponible. Las alertas no se enviarán.")
            return False
            
        try:
            # Cargar configuración
            self.config = cargar_configuracion()
            
            if self.config is None:
                logger.warning("No se pudo cargar configuración. Usando configuración por defecto.")
                self.config = {}
            
            # Crear servicio de alertas
            self.alert_service = crear_alert_service(self.config, self.destinatario)
            
            # Establecer contexto
            if seccion:
                self.seccion_actual = seccion
            if fecha:
                self.fecha_proceso = fecha
                
            self.alert_service.establecer_contexto(
                seccion=self.seccion_actual,
                fecha=self.fecha_proceso,
                area=self.nombre_script
            )
            
            self.inicializado = True
            logger.info(f"IntegradorAlertas inicializado para {self.nombre_script}")
            return True
            
        except Exception as e:
            logger.error(f"Error al inicializar IntegradorAlertas: {e}")
            return False
    
    def establecer_contexto(self, seccion: str = None, fecha: str = None, area: str = None):
        """
        Establece el contexto para las alertas.
        
        Args:
            seccion (str): Nombre de la sección actual
            fecha (str): Fecha del proceso
            area (str): Área o nombre del proceso
        """
        if seccion:
            self.seccion_actual = seccion
        if fecha:
            self.fecha_proceso = fecha
            
        if self.alert_service and self.inicializado:
            self.alert_service.establecer_contexto(
                seccion=self.seccion_actual,
                fecha=self.fecha_proceso,
                area=area or self.nombre_script
            )
    
    def reportar_error(self, error: Exception, seccion: str = None, 
                       fase: str = None, contexto: Dict = None) -> bool:
        """
        Reporta un error mediante el sistema de alertas.
        
        Args:
            error (Exception): Excepción ocurrida
            seccion (str): Sección donde ocurrió el error
            fase (str): Fase del proceso donde ocurrió
            contexto (Dict): Información adicional de contexto
            
        Returns:
            bool: True si se reportó correctamente
        """
        seccion_final = seccion or self.seccion_actual
        
        if not self.alert_service or not self.inicializado:
            logger.error(f"Error en {seccion_final}: {error}")
            return False
            
        try:
            contexto_completo = {
                'seccion': seccion_final,
                'fase': fase or 'general',
                **(contexto or {})
            }
            
            # Determinar el tipo de alerta según la excepción
            if isinstance(error, FileNotFoundError):
                return self.alert_service.alerta_archivo_no_encontrado(
                    archivo=str(error.filename),
                    seccion=seccion_final
                )
            elif isinstance(error, PermissionError):
                return self.alert_service.enviar_alerta(
                    "PERMISO_DENEGADO",
                    {'archivo': str(error), 'seccion': seccion_final}
                )
            elif isinstance(error, MemoryError):
                return self.alert_service.enviar_alerta(
                    "MEMORY_ERROR",
                    {'seccion': seccion_final}
                )
            else:
                return self.alert_service.alerta_error_procesamiento(
                    seccion=seccion_final,
                    fase=fase or 'general',
                    error=error
                )
                
        except Exception as e:
            logger.error(f"Error al reportar alerta: {e}")
            return False
    
    def reportar_advertencia(self, titulo: str, descripcion: str = None,
                            seccion: str = None, contexto: Dict = None) -> bool:
        """
        Reporta una advertencia mediante el sistema de alertas.
        
        Args:
            titulo (str): Título de la advertencia
            descripcion (str): Descripción detallada
            seccion (str): Sección donde ocurrió
            contexto (Dict): Información adicional
            
        Returns:
            bool: True si se reportó correctamente
        """
        seccion_final = seccion or self.seccion_actual
        
        if not self.alert_service or not self.inicializado:
            logger.warning(f"Advertencia en {seccion_final}: {titulo}")
            return False
            
        try:
            return self.alert_service.alerta_warning(
                titulo=titulo,
                descripcion=descripcion or titulo,
                seccion=seccion_final
            )
        except Exception as e:
            logger.error(f"Error al reportar advertencia: {e}")
            return False
    
    def reportar_datos_vacios(self, seccion: str, tipo_datos: str,
                             registros: int, minimo: int = 1) -> bool:
        """
        Reporta cuando se encuentran datos vacíos o insuficientes.
        
        Args:
            seccion (str): Sección afectada
            tipo_datos (str): Tipo de datos que está vacío
            registros (int): Número de registros encontrados
            minimo (int): Número mínimo esperado
            
        Returns:
            bool: True si se reportó correctamente
        """
        if not self.alert_service or not self.inicializado:
            logger.warning(f"Datos vacíos en {seccion}: {tipo_datos} - {registros} registros")
            return False
            
        return self.alert_service.alerta_datos_vacios(
            seccion=seccion,
            tipo_datos=tipo_datos,
            registros=registros,
            minimo=minimo
        )
    
    def reportar_columna_faltante(self, columna: str, archivo: str,
                                  columnas_disponibles: list, seccion: str = None) -> bool:
        """
        Reporta cuando falta una columna en un DataFrame.
        
        Args:
            columna (str): Nombre de la columna faltante
            archivo (str): Archivo donde se buscaba
            columnas_disponibles (list): Columnas disponibles
            seccion (str): Sección afectada
            
        Returns:
            bool: True si se reportó correctamente
        """
        seccion_final = seccion or self.seccion_actual
        
        if not self.alert_service or not self.inicializado:
            logger.warning(f"Columna '{columna}' no encontrada en {archivo}")
            return False
            
        return self.alert_service.alerta_columna_faltante(
            columna=columna,
            archivo=archivo,
            columnas_disponibles=columnas_disponibles,
            seccion=seccion_final
        )
    
    def reportar_excel_error(self, archivo: str, tipo_error: str,
                            seccion: str = None) -> bool:
        """
        Reporta errores al leer o escribir archivos Excel.
        
        Args:
            archivo (str): Ruta del archivo
            tipo_error (str): Descripción del error
            seccion (str): Sección afectada
            
        Returns:
            bool: True si se reportó correctamente
        """
        seccion_final = seccion or self.seccion_actual
        
        if not self.alert_service or not self.inicializado:
            logger.error(f"Error Excel en {archivo}: {tipo_error}")
            return False
            
        return self.alert_service.alerta_excel_error(
            archivo=archivo,
            tipo_error=tipo_error,
            seccion=seccion_final
        )
    
    def reportar_categoria_faltante(self, seccion: str, articulos: list) -> bool:
        """
        Reporta cuando hay artículos sin categoría ABC asignada.
        
        Args:
            seccion (str): Sección afectada
            articulos (list): Lista de artículos sin categoría
            
        Returns:
            bool: True si se reportó correctamente
        """
        if not self.alert_service or not self.inicializado:
            logger.warning(f"Categoría faltante en {seccion}: {len(articulos)} artículos")
            return False
            
        return self.alert_service.alerta_categoria_faltante(
            seccion=seccion,
            articulos=articulos
        )
    
    def obtener_resumen(self) -> Dict[str, Any]:
        """
        Obtiene un resumen de las alertas enviadas.
        
        Returns:
            Dict con información de alertas enviadas
        """
        if not self.alert_service or not self.inicializado:
            return {'total': 0, 'alertas': []}
            
        return self.alert_service.obtener_resumen_alertas()
    
    def finalizar(self) -> Dict[str, Any]:
        """
        Finaliza el integrador y obtiene resumen de alertas.
        
        Returns:
            Dict con resumen de alertas enviadas
        """
        resumen = self.obtener_resumen()
        
        if self.inicializado:
            logger.info(f"IntegradorAlertas finalizar. Total alertas: {resumen.get('total_enviadas', 0)}")
        
        return resumen


def crear_integrador(nombre_script: str, destinatario: str = "ivan.delgado@viveverde.es") -> IntegradorAlertas:
    """
    Crea una instancia del IntegradorAlertas.
    
    Args:
        nombre_script (str): Nombre del script
        destinatario (str): Email destinatario
        
    Returns:
        IntegradorAlertas: Instancia configurada
    """
    return IntegradorAlertas(nombre_script, destinatario)


def ejecutar_con_alertas(integrador: IntegradorAlertas, seccion: str, 
                        func: Callable, *args, **kwargs) -> Any:
    """
    Ejecuta una función con manejo automático de alertas.
    
    Args:
        integrador (IntegradorAlertas): Instancia del integrador
        seccion (str): Nombre de la sección/proceso
        func (Callable): Función a ejecutar
        *args, **kwargs: Argumentos para la función
        
    Returns:
        Resultado de la función
        
    Raises:
        Excepción original si ocurre algún error
    """
    integrador.establecer_contexto(seccion=seccion)
    
    try:
        return func(*args, **kwargs)
    except Exception as e:
        integrador.reportar_error(e, seccion=seccion)
        raise


# ============================================================================
# FUNCIONES DE INTEGRACIÓN RÁPIDA PARA SCRIPTS EXISTENTES
# ============================================================================

def inicializar_alertas_script(nombre_script: str, seccion: str = None) -> IntegradorAlertas:
    """
    Inicializa rápidamente el sistema de alertas para un script.
    
    Args:
        nombre_script (str): Nombre del script
        seccion (str): Sección inicial
        
    Returns:
        IntegradorAlertas: Instancia configurada
    """
    integrador = crear_integrador(nombre_script)
    integrador.inicializar(seccion=seccion)
    return integrador


def envolver_con_try_except(integrador: IntegradorAlertas, seccion: str, 
                           func: Callable, *args, **kwargs) -> tuple:
    """
    Envuelve una función en un bloque try/except con reporte automático de errores.
    
    Args:
        integrador (IntegradorAlertas): Instancia del integrador
        seccion (str): Nombre de la sección
        func (Callable): Función a ejecutar
        *args, **kwargs: Argumentos para la función
        
    Returns:
        tuple: (resultado, error) donde resultado es None si hay error
    """
    integrador.establecer_contexto(seccion=seccion)
    
    try:
        resultado = func(*args, **kwargs)
        return resultado, None
    except Exception as e:
        integrador.reportar_error(e, seccion=seccion)
        return None, e


# ============================================================================
# EJEMPLO DE USO
# ============================================================================

if __name__ == "__main__":
    print("=" * 70)
    print("INTEGRADOR DE ALERTAS - MÓDULO DE INTEGRACIÓN")
    print("=" * 70)
    print("\nEste módulo proporciona funciones para integrar alertas en scripts.")
    print("\nUso básico:")
    print("""
    # 1. Importar el módulo
    from src.integracion_alertas import crear_integrador
    
    # 2. Inicializar al inicio del script
    integrador = crear_integrador("mi_script")
    integrador.inicializar(seccion="seccion_principal")
    
    # 3. Usar en bloques try/except
    try:
        # Tu código aquí
        resultado = proceso_datos()
    except Exception as e:
        integrador.reportar_error(e, seccion="proceso_datos")
        raise
    
    # 4. Para advertencias
    if datos_insuficientes:
        integrador.reportar_advertencia(
            titulo="Datos insuficientes",
            descripcion="Solo se encontraron X registros",
            seccion="validacion"
        )
    
    # 5. Al final del script
    integrador.finalizar()
    """)
    print("\n" + "=" * 70)
