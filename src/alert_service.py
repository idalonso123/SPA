#!/usr/bin/env python3
"""
Módulo AlertService - Sistema de notificaciones de alertas y errores

Este módulo gestiona el envío de correos electrónicos ONLY cuando hay problemas reales.
NO envía correos en ejecuciones exitosas - solo cuando hay errores, advertencias o problemas.

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-12
"""

import smtplib
import ssl
import os
import json
import logging
import traceback
import sys
import platform

# psutil es opcional - si no está disponible, se usa una versión simple
try:
    import psutil
    PSUTIL_DISPONIBLE = True
except ImportError:
    PSUTIL_DISPONIBLE = False

from datetime import datetime
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from typing import Optional, Dict, Any, List
from pathlib import Path
from enum import Enum

# Configuración del logger
logger = logging.getLogger(__name__)


class NivelAlerta(Enum):
    """Niveles de severidad de las alertas"""
    CRITICAL = "CRITICAL"    # Error crítico que detiene la ejecución
    ERROR = "ERROR"          # Error que afecta funcionalidad
    WARNING = "WARNING"      # Advertencia, proceso continúa
    INFO = "INFO"            # Información (no se usa para emails)


# ============================================================================
# TABLA COMPLETA DE ESCENARIOS DE ERROR Y PLANTILLAS DE MENSAJE
# ============================================================================

ESCENARIOS_ALERTAS = {
    # -------------------------------------------------------------------------
    # ERRORES DE ARCHIVOS Y SISTEMA DE ARCHIVOS
    # -------------------------------------------------------------------------
    "ARCHIVO_NO_ENCONTRADO": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Archivo no encontrado - {archivo}",
        "cuerpo": """Se ha producido un error al intentar acceder a un archivo necesario.

DETALLE DEL ERROR:
- Archivo: {archivo}
- Tipo de operación: {operacion}
- Ubicación esperada: {ruta}

INFORMACIÓN ADICIONAL:
- Directorio de trabajo: {cwd}
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar que el archivo existe en la ruta especificada y que los permisos son correctos.

Este error puede afectar el procesamiento de la sección: {seccion}""",
        "icono": "📁"
    },
    
    "PERMISO_DENEGADO": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Permiso denegado - {archivo}",
        "cuerpo": """Se ha producido un error de permisos al intentar acceder a un archivo.

DETALLE DEL ERROR:
- Archivo: {archivo}
- Tipo de operación: {operacion}
- Usuario del sistema: {usuario}

INFORMACIÓN ADICIONAL:
- Sistema operativo: {so}
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar los permisos del archivo y directorio. Puede ser necesario ejecutar el script con privilegios adecuados.""",
        "icono": "🔒"
    },
    
    "DIRECTORIO_NO_ENCONTRADO": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Directorio no encontrado - {directorio}",
        "cuerpo": """No se puede acceder al directorio necesario.

DETALLE DEL ERROR:
- Directorio: {directorio}
- Tipo de operación: {operacion}

INFORMACIÓN ADICIONAL:
- Directorio de trabajo: {cwd}
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Crear el directorio o verificar la configuración de rutas en config.json.""",
        "icono": "📂"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE DATOS Y EXCEL
    # -------------------------------------------------------------------------
    "EXCEL_ERROR_LECTURA": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al leer archivo Excel - {archivo}",
        "cuerpo": """Se ha producido un error al intentar leer un archivo Excel.

DETALLE DEL ERROR:
- Archivo: {archivo}
- Tipo de error: {tipo_error}
- Hoja(s) afectada(s): {hojas}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}
- Tamaño del archivo: {tamano}

ACCIÓN RECOMENDADA:
Verificar que el archivo Excel no está corrupto y que el formato es correcto (xlsx o xls).

Traza del error:
{traza}""",
        "icono": "📊"
    },
    
    "EXCEL_ERROR_ESCRITURA": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al escribir archivo Excel - {archivo}",
        "cuerpo": """Se ha producido un error al intentar guardar un archivo Excel.

DETALLE DEL ERROR:
- Archivo: {archivo}
- Tipo de error: {tipo_error}
- Espacio en disco disponible: {espacio_disco}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar espacio en disco, permisos de escritura y que el archivo no esté abierto en otra aplicación.

Traza del error:
{traza}""",
        "icono": "💾"
    },
    
    "JSON_ERROR_PARSEO": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al parsear JSON - {archivo}",
        "cuerpo": """Se ha producido un error al intentar leer un archivo JSON.

DETALLE DEL ERROR:
- Archivo: {archivo}
- Línea/posición del error: {linea}

CONTENIDO PROBLEMÁTICO (primeros 500 caracteres):
{contenido_problematico}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar la sintaxis del archivo JSON. Puede haber caracteres inválidos o comas faltantes.""",
        "icono": "📄"
    },
    
    "CSV_ERROR_LECTURA": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al leer archivo CSV - {archivo}",
        "cuerpo": """Se ha producido un error al intentar leer un archivo CSV.

DETALLE DEL ERROR:
- Archivo: {archivo}
- Tipo de error: {tipo_error}
- Encoding detectado: {encoding}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar el formato del archivo CSV y el encoding utilizado (utf-8, latin-1, etc.).""",
        "icono": "📋"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE DATOS Y VALIDACIÓN
    # -------------------------------------------------------------------------
    "COLUMNA_NO_ENCONTRADA": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Columna no encontrada - {columna}",
        "cuerpo": """No se ha encontrado una columna requerida en los datos.

DETALLE DEL ERROR:
- Columna buscada: {columna}
- Archivo de origen: {archivo}
- Columnas disponibles: {columnas_disponibles}

INFORMACIÓN ADICIONAL:
- Sección afectada: {seccion}
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar que el archivo de datos tiene la columna correcta. Posiblemente el nombre de la columna ha cambiado.""",
        "icono": "🔍"
    },
    
    "DATOS_VACIOS": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Datos vacíos o insuficientes - {seccion}",
        "cuerpo": """El procesamiento ha encontrado datos vacíos o insuficientes.

DETALLE:
- Sección: {seccion}
- Tipo de datos: {tipo_datos}
- Registros encontrados: {registros}
- Registros esperados (mínimo): {minimo_registros}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}
- Archivo fuente: {archivo}

ACCIÓN RECOMENDADA:
Verificar que el archivo de datos contiene información válida y está actualizado.""",
        "icono": "⚠️"
    },
    
    "VALIDACION_FALLIDA": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Validación de datos fallida - {seccion}",
        "cuerpo": """La validación de datos ha encontrado problemas.

DETALLE:
- Sección: {seccion}
- Campo(s) afectado(s): {campos}
- Valor(es) problemático(s): {valores}
- Regla de validación: {regla}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Revisar los datos de entrada para corregir los valores problemáticos.""",
        "icono": "✅"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE CONFIGURACIÓN
    # -------------------------------------------------------------------------
    "CONFIG_ERROR": {
        "nivel": NivelAlerta.CRITICAL,
        "asunto": "[CRITICAL] Error de configuración - {detalle}",
        "cuerpo": """Se ha encontrado un error en la configuración del sistema.

DETALLE DEL ERROR:
- Configuración afectada: {configuracion}
- Valor actual: {valor_actual}
- Valor esperado: {valor_esperado}
- Archivo de configuración: {archivo_config}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Revisar el archivo de configuración y corregir los valores indicados.""",
        "icono": "⚙️"
    },
    
    "VARIABLE_ENTORNO_FALTANTE": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Variable de entorno faltante - {variable}",
        "cuerpo": """Una variable de entorno requerida no está configurada.

DETALLE DEL ERROR:
- Variable: {variable}
- Descripción: {descripcion}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Configurar la variable de entorno en el sistema o en un archivo .env""",
        "icono": "🔧"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE PROCESAMIENTO
    # -------------------------------------------------------------------------
    "PROCESAMIENTO_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error en procesamiento - {seccion}",
        "cuerpo": """Se ha producido un error durante el procesamiento de datos.

DETALLE:
- Sección: {seccion}
- Fase del proceso: {fase}
- Paso específico: {paso}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

Traza del error:
{traza}

ACCIÓN RECOMENDADA:
Revisar los logs para más detalles sobre el error específico.""",
        "icono": "🔄"
    },
    
    "EXCEPCION_NO_ESPERADA": {
        "nivel": NivelAlerta.CRITICAL,
        "asunto": "[CRITICAL] Excepción no manejada - {tipo_excepcion}",
        "cuerpo": """Se ha producido una excepción no manejada que ha detenido la ejecución.

DETALLE:
- Tipo de excepción: {tipo_excepcion}
- Mensaje: {mensaje}
- Sección afectada: {seccion}

INFORMACIÓN DEL SISTEMA:
- Python: {python_version}
- Sistema operativo: {so}
- Memoria RAM usada: {memoria_usada} MB
- Fecha y hora: {timestamp}

Traza completa del error:
{traza}

ACCIÓN RECOMENDADA:
Este es un error crítico. Revisar el código y la traza para identificar la causa raíz.""",
        "icono": "💥"
    },
    
    "MEMORY_ERROR": {
        "nivel": NivelAlerta.CRITICAL,
        "asunto": "[CRITICAL] Error de memoria - Out of Memory",
        "cuerpo": """El sistema se ha quedado sin memoria durante la ejecución.

DETALLE:
- Memoria total: {memoria_total} MB
- Memoria usada: {memoria_usada} MB
- Memoria disponible: {memoria_disponible} MB
- Porcentaje usado: {porcentaje_usado}%

INFORMACIÓN ADICIONAL:
- Sección procesando: {seccion}
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
- Reducir el tamaño de los datos procesados
- Reiniciar el script
- Aumentar la memoria disponible""",
        "icono": "🧠"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE EMAIL
    # -------------------------------------------------------------------------
    "EMAIL_ERROR_ENVIO": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al enviar email - {destinatario}",
        "cuerpo": """Se ha producido un error al intentar enviar un correo electrónico.

DETALLE:
- Destinatario: {destinatario}
- Asunto del email: {asunto}
- Tipo de error SMTP: {tipo_error}
- Detalles adicionales: {detalles}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
- Verificar la configuración SMTP
- Verificar la conexión a internet
- Comprobar que el servidor de correo está disponible""",
        "icono": "📧"
    },
    
    "EMAIL_CONFIG_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Configuración de email incorrecta",
        "cuerpo": """La configuración del servicio de email es incorrecta.

DETALLE:
- Servidor SMTP: {smtp_servidor}
- Puerto SMTP: {smtp_puerto}
- Remitente: {remitente}
- Error específico: {error_especifico}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Revisar la configuración del email en config.json y verificar las credenciales.""",
        "icono": "📬"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE RED (si aplica)
    # -------------------------------------------------------------------------
    "CONEXION_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error de conexión - {servicio}",
        "cuerpo": """Se ha producido un error de conexión de red.

DETALLE:
- Servicio/Dominio: {servicio}
- Tipo de conexión: {tipo_conexion}
- Código de error: {codigo_error}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
- Verificar la conexión a internet
- Verificar que el servicio está disponible
- Comprobar el firewall""",
        "icono": "🌐"
    },
    
    "TIMEOUT_ERROR": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Timeout - operación tardó demasiado - {operacion}",
        "cuerpo": """Una operación ha excedido el tiempo máximo de espera.

DETALLE:
- Operación: {operacion}
- Tiempo límite: {tiempo_limite} segundos
- Tiempo transcurrido: {tiempo_transcurrido} segundos

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
- Verificar la conexión de red
- El servicio puede estar sobrecargado
- Considerar aumentar el timeout""",
        "icono": "⏱️"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE CLASIFICACIÓN ABC
    # -------------------------------------------------------------------------
    "CLASIFICACION_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error en clasificación ABC - {seccion}",
        "cuerpo": """Se ha producido un error al procesar la clasificación ABC.

DETALLE:
- Sección: {seccion}
- Archivo de clasificación: {archivo}
- Tipo de error: {tipo_error}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar el archivo de clasificación ABC para la sección afectada.""",
        "icono": "🔢"
    },
    
    "CLASIFICACION_CATEGORIA_FALTANTE": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Categoría faltante en clasificación - {articulos_count} artículos",
        "cuerpo": """Se han encontrado artículos sin categoría ABC asignada.

DETALLE:
- Cantidad de artículos sin categoría: {articulos_count}
- Sección: {seccion}
- Archivo de origen: {archivo}

ARTÍCULOS AFECTADOS (primeros 20):
{articulos_afectados}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Revisar el archivo de clasificación ABC y asignar categoría a los artículos listados.""",
        "icono": "❓"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE ESTADO (STATE)
    # -------------------------------------------------------------------------
    "STATE_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error en gestión de estado - {operacion}",
        "cuerpo": """Se ha producido un error al gestionar el archivo de estado.

DETALLE:
- Operación: {operacion}
- Archivo de estado: {archivo_state}
- Tipo de error: {tipo_error}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar el archivo state.json y sus permisos. Puede ser necesario reiniciar el estado.""",
        "icono": "💼"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE PEDIDOS
    # -------------------------------------------------------------------------
    "PEDIDO_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al generar pedido - {seccion}",
        "cuerpo": """Se ha producido un error al generar el pedido de compra.

DETALLE:
- Sección: {seccion}
- Semana: {semana}
- Fase del error: {fase}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

Traza del error:
{traza}

ACCIÓN RECOMENDADA:
Revisar los datos de entrada y los logs para identificar el problema.""",
        "icono": "📦"
    },
    
    "PEDIDO_SIN_DATOS": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Pedido sin datos生成 - {seccion}",
        "cuerpo": """No se han podido generar pedidos para una sección.

DETALLE:
- Sección: {seccion}
- Semana: {semana}
- Razón: {razon}
- Registros procesados: {registros}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Verificar que existen datos de ventas y stock para la sección indicada.""",
        "icono": "📭"
    },
    
    # -------------------------------------------------------------------------
    # ERRORES DE RESUMEN
    # -------------------------------------------------------------------------
    "RESUMEN_ERROR": {
        "nivel": NivelAlerta.ERROR,
        "asunto": "[ERROR] Error al generar resumen",
        "cuerpo": """Se ha producido un error al generar el resumen consolidado.

DETALLE:
- Tipo de error: {tipo_error}
- Secciones procesadas: {secciones_procesadas}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

Traza del error:
{traza}

ACCIÓN RECOMENDADA:
Revisar los logs para más detalles.""",
        "icono": "📝"
    },
    
    # -------------------------------------------------------------------------
    # WARNING GENERALES
    # -------------------------------------------------------------------------
    "WARNING_GENERICO": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Advertencia - {titulo}",
        "cuerpo": """Se ha producido una advertencia durante la ejecución.

DETALLE:
- Título: {titulo}
- Descripción: {descripcion}
- Sección afectada: {seccion}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Revisar los logs para más detalles sobre esta advertencia.""",
        "icono": "⚡"
    },
    
    "DUPLICADOS_DETECTADOS": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Duplicados detectados - {seccion}",
        "cuerpo": """Se han detectado registros duplicados en los datos.

DETALLE:
- Sección: {seccion}
- Cantidad de duplicados: {cantidad}
- Campo(s) duplicado(s): {campos}
- Registros afectados (primeros 10): {registros}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Revisar los datos de origen para eliminar duplicados.""",
        "icono": "🔄"
    },
    
    "DATOS_INCONSISTENTES": {
        "nivel": NivelAlerta.WARNING,
        "asunto": "[WARNING] Datos inconsistentes detectados - {seccion}",
        "cuerpo": """Se han detectado inconsistencias en los datos.

DETALLE:
- Sección: {seccion}
- Tipo de inconsistencia: {tipo_inconsistencia}
- Descripción: {descripcion}
- Registros afectados: {registros_afectados}

INFORMACIÓN ADICIONAL:
- Fecha y hora: {timestamp}

ACCIÓN RECOMENDADA:
Investigar la causa de las inconsistencias y corregir los datos de origen.""",
        "icono": "🔀"
    },
}


class AlertService:
    """
    Servicio de alertas y notificaciones de errores.
    
    Esta clase encapsula toda la lógica de envío de alertas por email.
    SOLO envía correos cuando hay errores o problemas - NO envía en ejecuciones exitosas.
    
    Attributes:
        config (dict): Configuración del sistema
        smtp_config (dict): Configuración del servidor SMTP
        remitente (dict): Información del remitente
        destinatario_principal (str): Email principal para alertas (ivan.delgado@viveverde.es)
        alertas_enviadas (list): Historial de alertas enviadas en esta ejecución
    """
    
    def __init__(self, config: dict, destinatario: str = "ivan.delgado@viveverde.es"):
        """
        Inicializa el AlertService con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
            destinatario (str): Email destinatario para las alertas
        """
        self.config = config
        self.destinatario_principal = destinatario
        self.smtp_config = {}
        self.remitente = {}
        self.alertas_enviadas = []
        self.alertas_contador = {}  # Para evitar spam de alertas similares
        self._alertas_deshabilitadas = False  # Para deshabilitar alertas si hay error de autenticación
        
        # Cargar configuración
        self._cargar_configuracion()
        
        logger.info("AlertService inicializado correctamente")
        logger.info(f"Destinatario de alertas: {self.destinatario_principal}")
    
    def _cargar_configuracion(self):
        """Carga la configuración del email desde el diccionario config."""
        email_config = self.config.get('email', {})
        
        # Configuración SMTP
        self.smtp_config = {
            'servidor': email_config.get('servidor', 'smtp.serviciodecorreo.es'),
            'puerto': email_config.get('puerto', 465),
            'usar_ssl': email_config.get('usar_ssl', True),
            'usar_tls': email_config.get('usar_tls', False)
        }
        
        # Remitente
        self.remitente = {
            'email': email_config.get('remitente', {}).get('email', 'ivan.delgado@viveverde.es'),
            'nombre': email_config.get('remitente', {}).get('nombre', 'Sistema de Alertas VIVEVERDE')
        }
        
        # Destinatario de alertas - leer desde config o usar variable
        env_config = self.config.get('env_email', {})
        self.destinatario_principal = env_config.get('destinatario_alertas', self.destinatario_principal)
    
    def _obtener_password(self) -> str:
        """Obtiene la contraseña del remitente desde variable de entorno."""
        email_config = self.config.get('email', {})
        password_var = email_config.get('password_var', 'EMAIL_PASSWORD')
        
        password = os.environ.get(password_var)
        
        if not password:
            logger.warning(f"Variable de entorno '{password_var}' no configurada para alertas")
            return None
        
        return password
    
    def _obtener_info_sistema(self) -> Dict[str, str]:
        """Obtiene información del sistema para incluir en las alertas."""
        try:
            if PSUTIL_DISPONIBLE:
                memoria = psutil.virtual_memory()
                return {
                    'python_version': f"{sys.version.split()[0]}",
                    'so': f"{platform.system()} {platform.release()}",
                    'memoria_usada': int(memoria.used / (1024 * 1024)),
                    'memoria_total': int(memoria.total / (1024 * 1024)),
                    'memoria_disponible': int(memoria.available / (1024 * 1024)),
                    'porcentaje_usado': memoria.percent,
                    'usuario': os.environ.get('USERNAME', os.environ.get('USER', 'desconocido')),
                }
            else:
                raise Exception("psutil no disponible")
        except Exception:
            return {
                'python_version': sys.version.split()[0],
                'so': platform.system(),
                'memoria_usada': 'N/A',
                'memoria_total': 'N/A',
                'memoria_disponible': 'N/A',
                'porcentaje_usado': 'N/A',
                'usuario': os.environ.get('USERNAME', os.environ.get('USER', 'desconocido')),
            }
    
    def _obtener_contexto_base(self) -> Dict[str, str]:
        """Obtiene información base del contexto para todas las alertas."""
        info_sistema = self._obtener_info_sistema()
        
        return {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'cwd': os.getcwd(),
            **info_sistema
        }
    
    def _formatear_mensaje(self, tipo_alerta: str, contexto: Dict[str, str]) -> Dict[str, str]:
        """
        Formatea el asunto y cuerpo del mensaje según el tipo de alerta.
        
        Args:
            tipo_alerta (str): Tipo de alerta de ESCENARIOS_ALERTAS
            contexto (dict): Variables específicas para la alerta
            
        Returns:
            dict: Asunto y cuerpo formateados
        """
        if tipo_alerta not in ESCENARIOS_ALERTAS:
            tipo_alerta = "WARNING_GENERICO"
        
        plantilla = ESCENARIOS_ALERTAS[tipo_alerta]
        
        # Combinar contexto base con contexto específico
        contexto_completo = {**self._obtener_contexto_base(), **contexto}
        
        # Formatear asunto
        try:
            asunto = plantilla['asunto'].format(**contexto_completo)
        except KeyError:
            asunto = f"[ALERTA] {tipo_alerta} - {contexto.get('seccion', 'Sistema')}"
        
        # Formatear cuerpo
        try:
            cuerpo = plantilla['cuerpo'].format(**contexto_completo)
        except KeyError as e:
            cuerpo = f"Error al formatear mensaje: clave faltante {e}\n\nContexto: {contexto_completo}"
        
        return {
            'asunto': asunto,
            'cuerpo': cuerpo,
            'nivel': plantilla['nivel'],
            'icono': plantilla['icono']
        }
    
    def _crear_mensaje_alerta(self, asunto: str, cuerpo: str, tipo_alerta: str) -> MIMEMultipart:
        """Crea el mensaje MIME para la alerta."""
        msg = MIMEMultipart()
        msg['From'] = f"{self.remitente['nombre']} <{self.remitente['email']}>"
        msg['To'] = self.destinatario_principal
        msg['Subject'] = asunto
        
        # Crear cuerpo HTML con estilo
        html_cuerpo = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background-color: #ff6b6b; color: white; padding: 15px; border-radius: 5px; }}
                .content {{ background-color: #f9f9f9; padding: 15px; border-radius: 5px; margin-top: 10px; }}
                .footer {{ color: #666; font-size: 12px; margin-top: 20px; }}
                pre {{ background-color: #f4f4f4; padding: 10px; border-radius: 5px; overflow-x: auto; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h2>🔔 Alerta del Sistema de Pedidos VIVEVERDE</h2>
            </div>
            <div class="content">
                <pre>{cuerpo}</pre>
            </div>
            <div class="footer">
                <p>Este email ha sido enviado automáticamente por el Sistema de Pedidos VIVEVERDE.</p>
                <p>Fecha de generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </div>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        msg.attach(MIMEText(html_cuerpo, 'html', 'utf-8'))
        
        return msg
    
    def _enviar_email(self, msg: MIMEMultipart) -> bool:
        """Envía el email a través del servidor SMTP."""
        try:
            password = self._obtener_password()
            
            if not password:
                logger.error("No se puede enviar alerta: contraseña no configurada")
                return False
            
            # Crear contexto SSL
            context = ssl.create_default_context()
            
            # Conectar al servidor
            if self.smtp_config.get('usar_ssl', True):
                with smtplib.SMTP_SSL(
                    self.smtp_config['servidor'],
                    self.smtp_config['puerto'],
                    context=context
                ) as server:
                    server.login(self.remitente['email'], password)
                    server.sendmail(
                        self.remitente['email'],
                        self.destinatario_principal,
                        msg.as_string()
                    )
            else:
                with smtplib.SMTP(
                    self.smtp_config['servidor'],
                    self.smtp_config['puerto']
                ) as server:
                    if self.smtp_config.get('usar_tls', False):
                        server.starttls(context=context)
                    server.login(self.remitente['email'], password)
                    server.sendmail(
                        self.remitente['email'],
                        self.destinatario_principal,
                        msg.as_string()
                    )
            
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            logger.error(f"Error de autenticación SMTP al enviar alerta: {e}")
            # No intentar enviar más alertas si la contraseña es inválida
            logger.warning("Alertas deshabilitadas temporalmente debido a error de autenticación")
            self._alertas_deshabilitadas = True
            return False
        except Exception as e:
            logger.error(f"Error al enviar alerta por email: {e}")
            return False
    
    def _evitar_spam(self, tipo_alerta: str, clave_unica: str = None) -> bool:
        """
        Evita enviar múltiples alertas similares en poco tiempo.
        
        Args:
            tipo_alerta: Tipo de alerta
            clave_unica: Clave única para identificar la alerta específica
            
        Returns:
            bool: True si se puede enviar, False si ya se envió recientemente
        """
        import time
        
        if clave_unica is None:
            clave_unica = tipo_alerta
        
        # Crear clave de identificación
        key = f"{tipo_alerta}_{clave_unica}"
        
        # Verificar si ya se envió recientemente (últimos 60 minutos)
        tiempo_actual = time.time()
        
        if key in self.alertas_contador:
            tiempo_anterior = self.alertas_contador[key]
            if tiempo_actual - tiempo_anterior < 3600:  # 1 hora
                logger.debug(f"Alerta {tipo_alerta} suprimida (enviada recientemente)")
                return False
        
        # Registrar esta alerta
        self.alertas_contador[key] = tiempo_actual
        return True
    
    def enviar_alerta(self, tipo_alerta: str, contexto: Dict[str, str], 
                     clave_unica: str = None) -> bool:
        """
        Envía una alerta por email.
        
        Args:
            tipo_alerta (str): Tipo de alerta de ESCENARIOS_ALERTAS
            contexto (dict): Variables específicas para la alerta
            clave_unica (str): Clave única para evitar spam de alertas similares
            
        Returns:
            bool: True si se envió correctamente, False si no
        """
        # Verificar si las alertas están deshabilitadas debido a error de autenticación
        if self._alertas_deshabilitadas:
            logger.debug(f"Alerta {tipo_alerta} omitida (alertas deshabilitadas por error de autenticación)")
            return False
        
        # Verificar si ya se envió recientemente
        if not self._evitar_spam(tipo_alerta, clave_unica):
            logger.info(f"Alerta {tipo_alerta} suprimida para evitar spam")
            return False
        
        # Formatear mensaje
        mensaje = self._formatear_mensaje(tipo_alerta, contexto)
        
        # Crear mensaje MIME
        msg = self._crear_mensaje_alerta(
            mensaje['asunto'], 
            mensaje['cuerpo'], 
            tipo_alerta
        )
        
        # Enviar email
        logger.info(f"Enviando alerta: {tipo_alerta} - {mensaje['asunto']}")
        enviado = self._enviar_email(msg)
        
        if enviado:
            self.alertas_enviadas.append({
                'tipo': tipo_alerta,
                'asunto': mensaje['asunto'],
                'timestamp': datetime.now().isoformat(),
                'contexto': contexto
            })
            logger.info(f"Alerta enviada correctamente: {tipo_alerta}")
        else:
            logger.error(f"Error al enviar alerta: {tipo_alerta}")
        
        return enviado
    
    # -------------------------------------------------------------------------
    # MÉTODOS DE convenience PARA ERRORES COMUNES
    # -------------------------------------------------------------------------
    
    def alerta_archivo_no_encontrado(self, archivo: str, seccion: str = "N/A", 
                                     operacion: str = "lectura") -> bool:
        """Envía alerta de archivo no encontrado."""
        return self.enviar_alerta("ARCHIVO_NO_ENCONTRADO", {
            'archivo': archivo,
            'seccion': seccion,
            'operacion': operacion
        }, clave_unica=archivo)
    
    def alerta_excel_error(self, archivo: str, tipo_error: str, 
                          seccion: str = "N/A") -> bool:
        """Envía alerta de error al leer/escribir Excel."""
        return self.enviar_alerta("EXCEL_ERROR_LECTURA", {
            'archivo': archivo,
            'tipo_error': tipo_error,
            'seccion': seccion,
            'hojas': 'N/A',
            'tamano': 'N/A'
        }, clave_unica=archivo)
    
    def alerta_columna_faltante(self, columna: str, archivo: str, 
                               columnas_disponibles: List[str], 
                               seccion: str = "N/A") -> bool:
        """Envía alerta de columna no encontrada."""
        return self.enviar_alerta("COLUMNA_NO_ENCONTRADA", {
            'columna': columna,
            'archivo': archivo,
            'columnas_disponibles': ', '.join(columnas_disponibles[:10]),
            'seccion': seccion
        }, clave_unica=f"{archivo}_{columna}")
    
    def alerta_excepcion(self, excepcion: Exception, seccion: str = "N/A") -> bool:
        """Envía alerta de excepción no manejada."""
        return self.enviar_alerta("EXCEPCION_NO_ESPERADA", {
            'tipo_excepcion': type(excepcion).__name__,
            'mensaje': str(excepcion),
            'seccion': seccion,
            'traza': traceback.format_exc()
        })
    
    def alerta_datos_vacios(self, seccion: str, tipo_datos: str,
                           registros: int, minimo: int = 1) -> bool:
        """Envía alerta de datos vacíos."""
        return self.enviar_alerta("DATOS_VACIOS", {
            'seccion': seccion,
            'tipo_datos': tipo_datos,
            'registros': registros,
            'minimo_registros': minimo,
            'archivo': 'N/A'
        }, clave_unica=f"{seccion}_{tipo_datos}")
    
    def alerta_categoria_faltante(self, seccion: str, articulos: List[str]) -> bool:
        """Envía alerta de categoría ABC faltante."""
        return self.enviar_alerta("CLASIFICACION_CATEGORIA_FALTANTE", {
            'seccion': seccion,
            'articulos_count': len(articulos),
            'archivo': 'CLASIFICACION_ABC',
            'articulos_afectados': '\n'.join(articulos[:20])
        }, clave_unica=seccion)
    
    def alerta_error_procesamiento(self, seccion: str, fase: str, 
                                  error: Exception) -> bool:
        """Envía alerta de error en procesamiento."""
        return self.enviar_alerta("PROCESAMIENTO_ERROR", {
            'seccion': seccion,
            'fase': fase,
            'paso': 'N/A',
            'traza': traceback.format_exc()
        })
    
    def alerta_pedido_error(self, seccion: str, semana: int, 
                           error: Exception) -> bool:
        """Envía alerta de error al generar pedido."""
        return self.enviar_alerta("PEDIDO_ERROR", {
            'seccion': seccion,
            'semana': semana,
            'fase': 'generación',
            'traza': traceback.format_exc()
        }, clave_unica=f"{seccion}_{semana}")
    
    def alerta_warning(self, titulo: str, descripcion: str, 
                      seccion: str = "N/A") -> bool:
        """Envía una advertencia genérica."""
        return self.enviar_alerta("WARNING_GENERICO", {
            'titulo': titulo,
            'descripcion': descripcion,
            'seccion': seccion
        })
    
    def alerta_error_envio(self, destinatario: str, asunto: str, 
                          tipo_error: str, detalles: str = "") -> bool:
        """
        Envía alerta cuando el envío de emails falla.
        
        Args:
            destinatario (str): Destinatario(s) del email que falló
            asunto (str): Asunto del email que se intentaba enviar
            tipo_error (str): Tipo de error de envío
            detalles (str): Detalles adicionales del error
            
        Returns:
            bool: True si se envió correctamente
        """
        return self.enviar_alerta("EMAIL_ERROR_ENVIO", {
            'destinatario': destinatario,
            'asunto': asunto,
            'tipo_error': tipo_error,
            'detalles': detalles
        }, clave_unica=f"email_error_{destinatario}_{asunto[:20]}")
    
    def alerta_config_error(self, detalle: str, configuracion: str) -> bool:
        """Envía alerta de error de configuración."""
        return self.enviar_alerta("CONFIG_ERROR", {
            'detalle': detalle,
            'configuracion': configuracion,
            'valor_actual': 'N/A',
            'valor_esperado': 'N/A',
            'archivo_config': 'config.json'
        })
    
    def obtener_resumen_alertas(self) -> Dict[str, Any]:
        """Obtiene un resumen de las alertas enviadas."""
        return {
            'total_enviadas': len(self.alertas_enviadas),
            'alertas': self.alertas_enviadas,
            'contador': {k: len([a for a in self.alertas_enviadas if a['tipo'] == k]) 
                        for k in set(a['tipo'] for a in self.alertas_enviadas)}
        }


def crear_alert_service(config: dict, destinatario: str = "ivan.delgado@viveverde.es") -> AlertService:
    """
    Crea una instancia del AlertService.
    
    Args:
        config (dict): Configuración del sistema
        destinatario (str): Email destinatario para las alertas
        
    Returns:
        AlertService: Instancia inicializada del servicio de alertas
    """
    return AlertService(config, destinatario)


# ============================================================================
# FUNCIONES DE INTEGRACIÓN CON EL SISTEMA EXISTENTE
# ============================================================================

def integrar_alertas_main(main_func):
    """
    Decorador para integrar alertas en la función main.
    
    Usage:
        @integrar_alertas_main
        def main():
            # código principal
            pass
    """
    def wrapper(*args, **kwargs):
        from src.config_loader import cargar_configuracion
        
        # Cargar configuración
        config = cargar_configuracion()
        
        # Crear servicio de alertas
        alert_service = crear_alert_service(config)
        
        try:
            # Ejecutar función principal
            result = main_func(*args, **kwargs)
            
            # Si hay alertasenviadas, mostrar resumen
            resumen = alert_service.obtener_resumen_alertas()
            if resumen['total_enviadas'] > 0:
                logger.warning(f"Se enviaron {resumen['total_enviadas']} alertas durante la ejecución")
            
            return result
            
        except Exception as e:
            # Enviar alerta de excepción crítica
            alert_service.alerta_excepcion(e, seccion="main")
            raise
    
    return wrapper


if __name__ == "__main__":
    # Ejemplo de uso
    print("AlertService - Módulo de alertas y notificaciones")
    print("=" * 60)
    print("\nEscenarios de alertas disponibles:")
    for key, value in ESCENARIOS_ALERTAS.items():
        print(f"  - {key}: {value['asunto'][:50]}...")
    
    print("\n" + "=" * 60)
    print("Para usar en el proyecto:")
    print("  from src.alert_service import crear_alert_service")
    print("  alert_service = crear_alert_service(config)")
    print("  alert_service.enviar_alerta('ARCHIVO_NO_ENCONTRADO', {...})")



# ============================================================================
# HANDLER DE LOGGING PARA ALERTAS AUTOMÁTICAS
# ============================================================================

class AlertLoggingHandler(logging.Handler):
    """
    Handler de logging que automáticamente convierte advertencias y errores en alertas por email.
    
    Este handler se conecta al sistema de logging de Python. Cada vez que se hace:
    - logger.warning("mensaje") -> Se envía una alerta WARNING
    - logger.error("mensaje") -> Se envía una alerta ERROR
    - logger.critical("mensaje") -> Se envía una alerta CRITICAL
    
    Usage:
        # En tu código de inicialización (ej: main.py):
        alert_handler = AlertLoggingHandler(config)
        logging.getLogger().addHandler(alert_handler)
        
        # Opcional: pasar un AlertService existente
        alert_service = crear_alert_service(config)
        alert_handler = AlertLoggingHandler(config, alert_service=alert_service)
    """
    
    def __init__(self, config: dict, nivel_minimo: int = logging.WARNING, alert_service=None):
        """
        Inicializa el handler.
        
        Args:
            config: Configuración del sistema
            nivel_minimo: Nivel mínimo de logging que triggers alertas (default: WARNING)
            alert_service: AlertService existente (opcional)
        """
        super().__init__()
        self.config = config
        self.nivel_minimo = nivel_minimo
        self.alert_service = alert_service
        self.formato = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # Usar el servicio de alertas proporcionado o crear uno nuevo
        if self.alert_service is None:
            try:
                self.alert_service = crear_alert_service(config)
                logging.debug("AlertLoggingHandler inicializado correctamente")
            except Exception as e:
                logging.error(f"Error al inicializar AlertService: {e}")
    
    def emit(self, record: logging.LogRecord):
        """Envía una alerta cuando se registra un mensaje de nivel WARNING o superior."""
        # Solo procesar si el nivel es suficiente y tenemos el servicio activo
        if record.levelno < self.nivel_minimo or not self.alert_service:
            return
        
        # No procesar si las alertas están deshabilitadas
        if getattr(self.alert_service, '_alertas_deshabilitadas', False):
            return
        
        try:
            # Obtener el nombre del módulo
            modulo = record.name
            mensaje = record.getMessage()
            
            # Ignorar mensajes del propio alert_service para evitar bucles
            if 'alert_service' in modulo.lower() or 'alerta' in mensaje.lower():
                return
            
            # Determinar tipo de alerta según el nivel
            if record.levelno >= logging.CRITICAL:
                tipo_alerta = "EXCEPCION_NO_ESPERADA"
                seccion = "CRITICAL"
            elif record.levelno >= logging.ERROR:
                tipo_alerta = "PROCESAMIENTO_ERROR"
                seccion = modulo
            else:  # WARNING
                tipo_alerta = "WARNING_GENERICO"
                seccion = modulo
            
            # Contexto para la alerta
            contexto = {
                'titulo': f"{record.levelname} en {modulo}",
                'descripcion': mensaje,
                'seccion': seccion
            }
            
            # Enviar alerta (usar clave única basada en mensaje para evitar spam)
            import hashlib
            clave = hashlib.md5(mensaje.encode()).hexdigest()[:8]
            
            self.alert_service.enviar_alerta(tipo_alerta, contexto, clave_unica=f"{modulo}_{clave}")
            
        except Exception as e:
            # No dejar que el handler de alertas falle el proceso
            logging.debug(f"Error en AlertLoggingHandler: {e}")


def configurar_alertas_logging(config: dict, alert_service=None) -> AlertLoggingHandler:
    """
    Configura el sistema de logging para enviar alertas automáticamente.
    
    Args:
        config: Configuración del sistema
        alert_service: AlertService existente (opcional)
        
    Returns:
        AlertLoggingHandler: El handler configurado
    """
    # Crear el handler
    alert_handler = AlertLoggingHandler(config, nivel_minimo=logging.WARNING, alert_service=alert_service)
    
    # Añadir al logger raíz
    root_logger = logging.getLogger()
    root_logger.addHandler(alert_handler)
    
    logging.info("Sistema de alertas de logging configurado correctamente")
    
    return alert_handler


# ============================================================================
# INTEGRACIÓN CON sys.excepthook PARA EXCEPCIONES NO MANEJADAS
# ============================================================================

def configurar_excepthook(config: dict, alert_service=None):
    """
    Configura un hook global para capturar excepciones no manejadas.
    
    Args:
        config: Configuración del sistema
        alert_service: AlertService existente (opcional)
    """
    if alert_service is None:
        alert_service = crear_alert_service(config)
    
    def excepthook(tipo, valor, traza):
        """
        Hook global para excepciones no manejadas.
        """
        # Formatear la traza
        traza_formateada = ''.join(traceback.format_exception(tipo, valor, traza))
        
        # Enviar alerta crítica
        contexto = {
            'tipo_excepcion': tipo.__name__,
            'mensaje': str(valor),
            'seccion': 'global',
            'traza': traza_formateada
        }
        
        alert_service.enviar_alerta("EXCEPCION_NO_ESPERADA", contexto)
        
        # También imprimir la traza normalmente
        print(traza_formateada)
    
    # Instalar el hook
    sys.excepthook = excepthook
    
    logging.info("Excepthook configurado para capturar excepciones no manejadas")


# ============================================================================
# INTEGRACIÓN RÁPIDA EN main.py
# ============================================================================

def iniciar_sistema_alertas(config: dict):
    """
    Inicializa el sistema completo de alertas.
    
    Llama a esta función al inicio de main.py para activar:
    1. Logging handler para warnings/errors automáticos
    2. Excepthook para excepciones no manejadas
    
    Args:
        config: Configuración del sistema
    """
    logging.info("Inicializando sistema de alertas...")
    
    # Obtener destinatario de alertas desde configuración
    env_config = config.get('env_email', {})
    destinatario_alertas = env_config.get('destinatario_alertas', 'ivan.delgado@viveverde.es')
    
    # Crear un único AlertService para compartir
    alert_service = crear_alert_service(config, destinatario=destinatario_alertas)
    
    # Configurar handler de logging con el servicio existente
    configurar_alertas_logging(config, alert_service=alert_service)
    
    # Configurar excepthook con el servicio existente
    configurar_excepthook(config, alert_service=alert_service)
    
    logging.info("Sistema de alertas completamente configurado")
    
    logging.info("Sistema de alertas completamente configurado")
