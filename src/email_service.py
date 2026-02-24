#!/usr/bin/env python3
"""
Módulo EmailService - Servicio de envío de correos electrónicos
Este módulo gestiona el envío de correos electrónicos con archivos adjuntos
utilizando SMTP. Los destinatarios se leen EXCLUSIVAMENTE desde 
config/encargados.json (FUENTE ÚNICA).
Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-05
"""
import smtplib
import ssl
import os
import json
import logging
import unicodedata
import pandas as pd
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from typing import Optional, List, Dict, Any
from pathlib import Path

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


class EmailService:
    """
    Servicio de envío de correos electrónicos para el sistema de pedidos.
    
    Esta clase encapsula toda la lógica relacionada con el envío de correos:
    configuración SMTP, gestión de destinatarios, generación de mensajes
    y envío de adjuntos.
    
    Los destinatarios se leen EXCLUSIVAMENTE desde encargacos.json (FUENTE ÚNICA).
    Los nombres de los encargados se leen desde config/encargados.json.
    
    Attributes:
        config (dict): Configuración del sistema
        smtp_config (dict): Configuración del servidor SMTP
        remitente (dict): Información del remitente
        destinatarios (dict): Mapeo sección -> lista de destinatarios desde encargacos.json
        plantilla_asunto (str): Template del asunto del email
        plantilla_cuerpo (str): Template del cuerpo del email
        encargados_por_seccion (dict): Mapeo de secciones a nombres de encargados
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el EmailService con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.smtp_config = {}
        self.remitente = {}
        self.destinatarios = {}
        self.plantilla_asunto = ""
        self.plantilla_cuerpo = ""
        self.encargados_por_seccion = {}
        
        # Cargar configuración
        self._cargar_configuracion()
        
        # Cargar nombres de encargados desde archivo JSON
        self._cargar_encargados()
        
        logger.info("EmailService inicializado correctamente")
        logger.info(f"Remitente: {self.remitente.get('email', 'No configurado')}")
        logger.info(f"Secciones con destinatarios: {list(self.destinatarios.keys())}")
    
    def _cargar_configuracion(self):
        """
        Carga la configuración del email desde el diccionario config.
        Extrae configuración SMTP, remitente, destinatarios y templates.
        - SMTP y plantillas: Se leen desde email.json (fuente centralizada)
        - Destinatarios: Se leen desde encargados.json (FUENTE ÚNICA)
        """
        # Primero intentar leer desde email.json (fuente centralizada)
        email_json_config = {}
        try:
            import os
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            email_json_path = os.path.join(base_dir, 'config', 'email.json')
            if os.path.exists(email_json_path):
                with open(email_json_path, 'r', encoding='utf-8') as f:
                    email_json_config = json.load(f)
                logger.debug("Configuración SMTP cargada desde email.json")
        except Exception as e:
            logger.warning(f"No se pudo cargar email.json: {e}")
        
        # Obtener configuración desde config.json
        email_config = self.config.get('email', {})
        
        # Configuración SMTP -优先从 email.json 读取
        if 'smtp' in email_json_config:
            smtp = email_json_config['smtp']
            self.smtp_config = {
                'servidor': smtp.get('servidor', 'smtp.serviciodecorreo.es'),
                'puerto': smtp.get('puerto', 465),
                'usar_ssl': smtp.get('usar_ssl', True),
                'usar_tls': smtp.get('usar_tls', False)
            }
        else:
            self.smtp_config = {
                'servidor': email_config.get('smtp', {}).get('servidor', 'smtp.serviciodecorreo.es'),
                'puerto': email_config.get('smtp', {}).get('puerto', 465),
                'usar_ssl': email_config.get('smtp', {}).get('usar_ssl', True),
                'usar_tls': email_config.get('smtp', {}).get('usar_tls', False)
            }
        
        # Remitente -优先从 email.json 读取
        if 'remitente' in email_json_config:
            remitente = email_json_config['remitente']
            self.remitente = {
                'email': remitente.get('email', 'ivan.delgado@viveverde.es'),
                'nombre': remitente.get('nombre', 'Sistema de Pedidos VIVEVERDE')
            }
        else:
            self.remitente = {
                'email': email_config.get('remitente', {}).get('email', 'ivan.delgado@viveverde.es'),
                'nombre': email_config.get('remitente', {}).get('nombre', 'Sistema de Pedidos VIVEVERDE')
            }
        
        # Destinatarios - Desde encargacos.json (FUENTE ÚNICA)
        # Leer los encargado desde el archivo dedicado
        # Formato: {seccion: [{'email': 'x@x.com', 'nombre': 'Nombre'}, ...]}
        self.destinatarios = {}
        self.encargados_por_seccion = {}  # Para compatibilidad con funciones existentes
        try:
            import os
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            encargado_path = os.path.join(base_dir, 'config', 'encargados.json')
            if os.path.exists(encargado_path):
                with open(encargado_path, 'r', encoding='utf-8') as f:
                    encargodos_data = json.load(f)
                
                # Transformar encargado.json al formato de destinatarios
                # Puede ser un objeto (un encargado) o un array (múltiples encargado)
                for seccion, encargado in encargodos_data.get('encargados', {}).items():
                    if isinstance(encargado, list):
                        # Múltiples encargado
                        destinatarios_seccion = []
                        for e in encargado:
                            if e.get('email'):
                                destinatarios_seccion.append({
                                    'email': e.get('email', ''),
                                    'nombre': e.get('nombre', 'Encargado')
                                })
                        self.destinatarios[seccion] = destinatarios_seccion
                        # También guardar para compatibilidad
                        self.encargados_por_seccion[seccion] = [e.get('nombre', 'Encargado') for e in encargado if e.get('nombre')]
                    elif isinstance(encargado, dict):
                        # Un solo encargado
                        email = encargado.get('email', '')
                        nombre = encargado.get('nombre', 'Encargado')
                        if email:
                            self.destinatarios[seccion] = [{
                                'email': email,
                                'nombre': nombre
                            }]
                        # También guardar para compatibilidad
                        self.encargados_por_seccion[seccion] = nombre
                
                logger.debug(f"Destinatarios cargados desde encargados.json: {list(self.destinatarios.keys())}")
        except Exception as e:
            logger.warning(f"No se pudo cargar encargado.json: {e}")
            # Fallback a config.json si no existe encargacos.json
            self.destinatarios = email_config.get('destinatarios', {})
        
        # Plantillas -优先从 email.json 读取
        if 'plantillas' in email_json_config:
            plantillas = email_json_config['plantillas']
            self.plantilla_asunto = plantillas.get('asunto_pedido', 
                'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}')
            self.plantilla_cuerpo = plantillas.get('cuerpo_pedido', 
                'Buenos días {nombre_encargado}. Te adjunto el pedido de compra generado '
                'para la semana {semana} de la sección {seccion}. Atentamente, '
                'Sistema de Pedidos automáticos VIVEVERDE.')
        else:
            self.plantilla_asunto = email_config.get('plantillas', {}).get(
                'asunto', 
                'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}'
            )
            self.plantilla_cuerpo = email_config.get('plantillas', {}).get(
                'cuerpo', 
                'Buenos días {nombre_encargado}. Te adjunto el pedido de compra generado '
                'para la semana {semana} de la sección {seccion}. Atentamente, '
                'Sistema de Pedidos automáticos VIVEVERDE.'
            )
        
        logger.debug("Configuración de email cargada desde email.json + encargados.json")
    
    def _cargar_encargados(self):
        """
        Carga información de encargados desde archivo JSON.
        NOTA: Los nombres de los encargados ahora se leen desde config/encargados.json
        en lugar del archivo Excel. Esto permite cambios futuros sin modificar código.
        
        El archivo JSON tiene la siguiente estructura:
        {
            "version": "1.0",
            "descripcion": "Configuración de encargados por sección",
            "encargados": {
                "interior": "Iris",
                "mascotas_manufacturado": "María",
                ...
            }
        }
        """
        try:
            # Construir ruta al archivo de configuración de encargados
            directorio_base = self.config.get('rutas', {}).get('directorio_base', '.')
            ruta_archivo = Path(directorio_base) / 'config' / 'encargados.json'
            
            # Verificar que el archivo existe
            if not ruta_archivo.exists():
                logger.warning(f"Archivo de encargados no encontrado: {ruta_archivo}")
                logger.info("Se usarán nombres genéricos para los encargados")
                return
            
            # Leer el archivo JSON
            with open(ruta_archivo, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # Extraer el diccionario de encargados
            self.encargados_por_seccion = config_data.get('encargados', {})
            
            logger.info(f"Encargados cargados desde {ruta_archivo}: {len(self.encargados_por_seccion)} secciones")
            logger.debug(f"Mapping de encargados: {self.encargados_por_seccion}")
            
        except json.JSONDecodeError as e:
            logger.error(f"Error al parsear archivo de encargados (JSON inválido): {e}")
            logger.info("Se usarán nombres genéricos para los encargados")
        except Exception as e:
            logger.error(f"Error al leer archivo de encargados: {e}")
            logger.info("Se usarán nombres genéricos para los encargados")
    
    def enviar_resumen_gestion(self, semana: int, archivo_resumen: str) -> Dict[str, Any]:
        """
        Envía el resumen de pedidos de compra a los responsables de gestión.
        
        Destinatarios:
        - Sandra: ivan.delgado@viveverde.es
        - Ivan: ivan.delgado@viveverde.es
        - Pedro: ivan.delgado@viveverde.es
        
        Args:
            semana (int): Número de semana procesada
            archivo_resumen (str): Ruta al archivo de resumen consolidado
            
        Returns:
            Dict[str, Any]: Resultado del envío con estado y detalles
        """
        logger.info("\n" + "=" * 60)
        logger.info("ENVÍO DE RESUMEN A RESPONSABLES DE GESTIÓN")
        logger.info("=" * 60)
        
        # Destinatarios del resumen de gestión
        destinatarios_resumen = [
            {'nombre': 'Sandra', 'email': 'sandra.delgado@viveverde.es'},
            {'nombre': 'Ivan', 'email': 'ivan.delgado@viveverde.es'},
            {'nombre': 'Pedro', 'email': 'pedro.delgado@viveverde.es'}
        ]
        
        # Verificar que el archivo de resumen existe
        if not Path(archivo_resumen).exists():
            logger.warning(f"Archivo de resumen no encontrado: {archivo_resumen}")
            logger.info("Omitiendo envío de resumen a responsables de gestión")
            return {'enviado': False, 'razon': 'archivo_no_encontrado'}
        
        try:
            # Generar asunto y cuerpo del email
            asunto = f"Viveverde: Resumen de pedidos de compra de las secciones semana {semana}"
            
            resultados_envio = {}
            emails_enviados = 0
            emails_fallidos = 0
            
            for destinatario in destinatarios_resumen:
                nombre = destinatario['nombre']
                email = destinatario['email']
                
                # Generar cuerpo personalizado para cada destinatario
                cuerpo = (f"Buenos días {nombre}.\n\n"
                         f"Te adjunto el resumen de los pedidos de compra de cada sección "
                         f"de la semana {semana}.\n\n"
                         f"Atentamente,\n"
                         f"Sistema de Pedidos automáticos VIVEVERDE.")
                
                # Crear y enviar mensaje
                msg = self._crear_mensaje([email], asunto, cuerpo, [archivo_resumen])
                enviado = self._enviar_email(msg)
                
                if enviado:
                    logger.info(f"✓ Resumen enviado a {nombre} ({email})")
                    emails_enviados += 1
                    resultados_envio[nombre] = {'email': email, 'enviado': True}
                else:
                    logger.error(f"✗ Error al enviar resumen a {nombre} ({email})")
                    emails_fallidos += 1
                    resultados_envio[nombre] = {'email': email, 'enviado': False}
            
            logger.info("\n" + "=" * 60)
            logger.info("RESUMEN DE ENVÍO A RESPONSABLES DE GESTIÓN")
            logger.info("=" * 60)
            logger.info(f"Emails enviados: {emails_enviados}")
            logger.info(f"Emails fallidos: {emails_fallidos}")
            
            return {
                'enviado': emails_enviados > 0,
                'emails_enviados': emails_enviados,
                'emails_fallidos': emails_fallidos,
                'resultados': resultados_envio
            }
            
        except Exception as e:
            logger.error(f"Error al enviar resumen de gestión: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return {'enviado': False, 'error': str(e)}
    
    def _normalizar_seccion(self, seccion: str) -> str:
        """
        Normaliza el nombre de una sección para buscar en el mapeo de encargados.
        
        Args:
            seccion (str): Nombre de la sección a normalizar
            
        Returns:
            str: Nombre de sección normalizado
        """
        return seccion.strip().lower().replace(' ', '_')
    
    def _obtener_password(self) -> str:
        """
        Obtiene la contraseña del remitente desde variable de entorno.
        
        Returns:
            str: Contraseña del correo remitente
            
        Raises:
            ValueError: Si la variable de entorno no está configurada
        """
        email_config = self.config.get('email', {})
        password_var = email_config.get('password_var', 'EMAIL_PASSWORD')
        
        password = os.environ.get(password_var)
        
        if not password:
            error_msg = (f"Variable de entorno '{password_var}' no configurada. "
                        "Configure la contraseña antes de enviar emails.")
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        return password
    
    def _generar_asunto(self, semana: int, seccion: str) -> str:
        """
        Genera el asunto del email usando la plantilla configurada.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            
        Returns:
            str: Asunto formateado
        """
        return self.plantilla_asunto.format(semana=semana, seccion=seccion)
    
    def _generar_cuerpo(self, semana: int, seccion: str, nombre_encargado: str) -> str:
        """
        Genera el cuerpo del email usando la plantilla configurada.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            nombre_encargado (str): Nombre del encargado
            
        Returns:
            str: Cuerpo del mensaje formateado
        """
        return self.plantilla_cuerpo.format(
            semana=semana,
            seccion=seccion,
            nombre_encargado=nombre_encargado
        )
    
    def _crear_mensaje(self, destinatarios: List[str], asunto: str, 
                      cuerpo: str, archivos_adjuntos: List[str]) -> MIMEMultipart:
        """
        Crea el mensaje MIME con adjuntos.
        
        Args:
            destinatarios (List[str]): Lista de correos destinatarios
            asunto (str): Asunto del email
            cuerpo (str): Cuerpo del mensaje
            archivos_adjuntos (List[str]): Lista de rutas de archivos a adjuntar
            
        Returns:
            MIMEMultipart: Mensaje MIME listo para enviar
        """
        msg = MIMEMultipart()
        msg['From'] = f"{self.remitente['nombre']} <{self.remitente['email']}>"
        msg['To'] = ', '.join(destinatarios)
        msg['Subject'] = asunto
        
        # Adjuntar cuerpo en texto plano
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar archivos
        for archivo in archivos_adjuntos:
            if not Path(archivo).exists():
                logger.warning(f"Archivo no encontrado: {archivo}")
                continue
            
            try:
                # Determinar tipo de archivo
                extension = Path(archivo).suffix.lower()
                
                if extension in ['.xlsx', '.xls']:
                    mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                elif extension == '.csv':
                    mime_type = 'text/csv'
                elif extension == '.pdf':
                    mime_type = 'application/pdf'
                else:
                    mime_type = 'application/octet-stream'
                
                # Leer archivo
                with open(archivo, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                
                # Codificar
                encoders.encode_base64(part)
                
                # Añadir header
                filename = Path(archivo).name
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= "{filename}"'
                )
                part.add_header('Content-Type', mime_type)
                
                msg.attach(part)
                logger.debug(f"Adjunto añadido: {filename}")
                
            except Exception as e:
                logger.error(f"Error al adjuntar archivo {archivo}: {e}")
        
        return msg
    
    def _enviar_email(self, msg: MIMEMultipart) -> bool:
        """
        Envía el email a través del servidor SMTP.
        
        Args:
            msg (MIMEMultipart): Mensaje MIME a enviar
            
        Returns:
            bool: True si el envío fue exitoso, False en caso contrario
        """
        try:
            password = self._obtener_password()
            
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
                        msg['To'].split(', '),
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
                        msg['To'].split(', '),
                        msg.as_string()
                    )
            
            logger.info(f"Email enviado exitosamente a: {msg['To']}")
            return True
            
        except smtplib.SMTPException as e:
            logger.error(f"Error SMTP al enviar email: {e}")
            return False
        except Exception as e:
            logger.error(f"Error al enviar email: {e}")
            return False
    
    def obtener_destinatarios_seccion(self, seccion: str) -> List[Dict[str, str]]:
        """
        Obtiene la lista de destinatarios para una sección específica.
        
        Los correos electrónicos y nombres se obtienen EXCLUSIVAMENTE desde 
        config/encargados.json (FUENTE ÚNICA).
        
        Args:
            seccion (str): Nombre de la sección
            
        Returns:
            List[Dict[str, str]]: Lista de diccionarios con 'email' y 'nombre'
        """
        seccion_normalizada = self._normalizar_seccion(seccion)
        
        # Obtener destinatarios desde encargacos.json (FUENTE ÚNICA)
        # La estructura ahora es: [{'email': 'x@x.com', 'nombre': 'Nombre'}, ...]
        destinatarios_data = self.destinatarios.get(seccion_normalizada, [])
        
        if not destinatarios_data:
            logger.warning(f"No hay destinatarios configurados para la sección: {seccion}")
            return []
        
        # Normalizar al formato de salida esperado
        destinatarios = []
        for item in destinatarios_data:
            if isinstance(item, dict):
                # Nueva estructura con email y nombre
                destinatarios.append({
                    'email': item.get('email', '').strip(),
                    'nombre': item.get('nombre', 'Encargado')
                })
            elif isinstance(item, str):
                # Formato antiguo (solo email), usar nombre genérico
                destinatarios.append({
                    'email': item.strip(),
                    'nombre': 'Encargado'
                })
        
        logger.debug(f"Destinatarios para {seccion}: {destinatarios}")
        return destinatarios
    
    def enviar_pedido_por_seccion(self, semana: int, seccion: str, 
                                  archivos: List[str]) -> Dict[str, Any]:
        """
        Envía los archivos de pedido por email al responsable de la sección.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            archivos (List[str]): Lista de rutas de archivos a adjuntar
            
        Returns:
            Dict[str, Any]: Resultado del envío con estado y detalles
        """
        resultado = {
            'seccion': seccion,
            'semana': semana,
            'enviado': False,
            'destinatarios': [],
            'archivos_adjuntos': [],
            'error': None
        }
        
        # Obtener destinatarios
        destinatarios = self.obtener_destinatarios_seccion(seccion)
        
        if not destinatarios:
            resultado['error'] = "No hay destinatarios configurados"
            logger.warning(f"No se puede enviar email para {seccion}: {resultado['error']}")
            return resultado
        
        # Filtrar archivos existentes
        archivos_existentes = [f for f in archivos if Path(f).exists()]
        resultado['archivos_adjuntos'] = [Path(f).name for f in archivos_existentes]
        
        if not archivos_existentes:
            resultado['error'] = "No hay archivos para adjuntar"
            logger.warning(f"No se puede enviar email para {seccion}: {resultado['error']}")
            return resultado
        
        # Generar asunto y cuerpo
        asunto = self._generar_asunto(semana, seccion)
        cuerpo = self._generar_cuerpo(semana, seccion, destinatarios[0]['nombre'])
        
        # Extraer solo los correos
        lista_correos = [d['email'] for d in destinatarios]
        resultado['destinatarios'] = lista_correos
        
        # Crear y enviar mensaje
        msg = self._crear_mensaje(lista_correos, asunto, cuerpo, archivos_existentes)
        resultado['enviado'] = self._enviar_email(msg)
        
        if resultado['enviado']:
            logger.info(f"Email enviado para sección {seccion} (semana {semana})")
        else:
            resultado['error'] = "Error en el envío"
        
        return resultado
    
    def enviar_resumen_centralizado(self, semana: int, 
                                    archivos: Dict[str, List[str]]) -> Dict[str, Any]:
        """
        Envía todos los pedidos de forma centralizada al responsable principal.
        
        Args:
            semana (int): Número de semana
            archivos (Dict[str, List[str]]): Archivos por sección
            
        Returns:
            Dict[str, Any]: Resultado del envío
        """
        email_config = self.config.get('email', {})
        email_centralizado = email_config.get('email_centralizado')
        
        if not email_centralizado:
            logger.info("No hay configuración de email centralizado, omitiendo envío")
            return {'enviado': False, 'razon': 'No configurado'}
        
        # Recopilar todos los archivos
        todos_archivos = []
        for seccion, archivos_seccion in archivos.items():
            todos_archivos.extend(archivos_seccion)
        
        if not todos_archivos:
            return {'enviado': False, 'razon': 'No hay archivos'}
        
        # Generar cuerpo del resumen
        cuerpo = f"Resumen de pedidos para la semana {semana}:\n\n"
        for seccion in archivos.keys():
            cuerpo += f"- {seccion}\n"
        cuerpo += "\nArchivos adjuntos.\n"
        
        asunto = f"VIVEVERDE: Resumen de Pedidos - Semana {semana}"
        
        # Filtrar archivos existentes
        archivos_existentes = [f for f in todos_archivos if Path(f).exists()]
        
        # Crear y enviar mensaje
        msg = self._crear_mensaje([email_centralizado], asunto, cuerpo, archivos_existentes)
        enviado = self._enviar_email(msg)
        
        return {
            'enviado': enviado,
            'destinatario': email_centralizado,
            'archivos': len(archivos_existentes)
        }
    
    def verificar_configuracion(self) -> Dict[str, Any]:
        """
        Verifica que la configuración del email sea correcta.
        
        Returns:
            Dict[str, Any]: Resultado de la verificación
        """
        resultado = {
            'valido': True,
            'remitente': {},
            'smtp': {},
            'destinatarios': {},
            'problemas': []
        }
        
        # Verificar remitente
        if self.remitente.get('email'):
            resultado['remitente'] = {
                'email': self.remitente['email'],
                'nombre': self.remitente.get('nombre', ''),
                'configurado': True
            }
        else:
            resultado['problemas'].append("Email del remitente no configurado")
            resultado['valido'] = False
        
        # Verificar SMTP
        resultado['smtp'] = {
            'servidor': self.smtp_config.get('servidor', ''),
            'puerto': self.smtp_config.get('puerto', ''),
            'ssl': self.smtp_config.get('usar_ssl', True),
            'tls': self.smtp_config.get('usar_tls', False)
        }
        
        # Verificar destinatarios
        secciones_configuradas = []
        for seccion, datos in self.destinatarios.items():
            # La nueva estructura puede ser [{'email':..., 'nombre':...}, ...] o lista de strings
            correos = []
            nombres = []
            
            if isinstance(datos, str):
                correos = [c.strip() for c in datos.split(',') if c.strip()]
                nombres = ['Encargado'] * len(correos)
            elif isinstance(datos, list):
                for item in datos:
                    if isinstance(item, dict):
                        email = item.get('email', '').strip()
                        nombre = item.get('nombre', 'Encargado')
                        if email:
                            correos.append(email)
                            nombres.append(nombre)
                    elif isinstance(item, str):
                        email = item.strip()
                        if email:
                            correos.append(email)
                            nombres.append('Encargado')
            
            if correos:
                secciones_configuradas.append(seccion)
                resultado['destinatarios'][seccion] = {
                    'correos': correos,
                    'nombre_encargado': ', '.join(nombres) if nombres else 'No definido'
                }
        
        resultado['secciones_configuradas'] = len(secciones_configuradas)
        
        if not secciones_configuradas:
            resultado['problemas'].append("No hay destinatarios configurados en encargados.json")
            resultado['valido'] = False
        
        # Verificar contraseña
        try:
            password = self._obtener_password()
            resultado['password'] = {'configurado': True, 'variable': 'EMAIL_PASSWORD'}
        except ValueError:
            resultado['password'] = {'configurado': False}
            resultado['problemas'].append("Contraseña no configurada (variable EMAIL_PASSWORD)")
        
        return resultado


def crear_email_service(config: dict) -> EmailService:
    """
    Crea una instancia del EmailService.
    
    Args:
        config (dict): Configuración del sistema
        
    Returns:
        EmailService: Instancia inicializada del servicio de email
    """
    return EmailService(config)


def verificar_configuracion_email(config: dict) -> Dict[str, Any]:
    """
    Verifica la configuración de email sin crear el servicio.
    
    Args:
        config (dict): Configuración del sistema
        
    Returns:
        Dict[str, Any]: Resultado de la verificación
    """
    try:
        email_config = config.get('email', {})
        
        # Verificar destinatarios desde encargacos.json
        destinatarios = email_config.get('destinatarios', {})
        secciones_con_correo = []
        
        for seccion, correos in destinatarios.items():
            if correos:
                secciones_con_correo.append(seccion)
        
        return {
            'valido': len(secciones_con_correo) > 0,
            'secciones_configuradas': len(secciones_con_correo),
            'secciones': secciones_con_correo,
            'remitente': email_config.get('remitente', {}).get('email', 'No configurado'),
            'smtp': email_config.get('servidor', 'No configurado'),
            'mensaje': ('Configuración verificada. Los correos se leen desde encargacos.json. '
                       'Los nombres de encargados se leen desde config/encargados.json.')
        }
        
    except Exception as e:
        return {
            'valido': False,
            'error': str(e)
        }


if __name__ == "__main__":
    # Ejemplo de uso
    print("EmailService - Módulo de envío de correos")
    print("=" * 50)
    
    # Configurar logging
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración
    config_ejemplo = {
        'email': {
            'servidor': 'smtp.serviciodecorreo.es',
            'puerto': 465,
            'usar_ssl': True,
            'remitente': {
                'email': 'ivan.delgado@viveverde.es',
                'nombre': 'Sistema de Pedidos VIVEVERDE'
            },
            'destinatarios': {
                'maf': 'ivan.delgado@viveverde.es',
                'interior': 'ivan.delgado@viveverde.es',
                'mascotas_vivo': 'ivan.delgado@viveverde.es',
                'mascotas_manufacturado': 'ivan.delgado@viveverde.es',
                'deco_interior': 'ivan.delgado@viveverde.es',
                'deco_exterior': 'ivan.delgado@viveverde.es',
                'tierras_aridos': 'ivan.delgado@viveverde.es',
                'fitos': 'ivan.delgado@viveverde.es',
                'semillas': 'ivan.delgado@viveverde.es',
                'utiles_jardin': 'ivan.delgado@viveverde.es',
                'vivero': 'ivan.delgado@viveverde.es'
            },
            'plantillas': {
                'asunto': 'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}',
                'cuerpo': 'Buenos días {nombre_encargado}. Te adjunto el pedido de compra '
                         'generado para la semana {semana} de la sección {seccion}. '
                         'Atentamente, Sistema de Pedidos automáticos VIVEVERDE.'
            }
        },
        'archivos_entrada': {
            'archivo_encargados': None
        },
        'rutas': {
            'directorio_base': '.'
        }
    }
    
    # Verificar configuración
    print("\nVerificación de configuración:")
    verificacion = verificar_configuracion_email(config_ejemplo)
    for clave, valor in verificacion.items():
        print(f"  {clave}: {valor}")
    
    # Crear servicio
    print("\nCreando EmailService:")
    email_service = crear_email_service(config_ejemplo)
    
    # Verificar configuración completa
    print("\nVerificación completa:")
    verif = email_service.verificar_configuracion()
    for clave, valor in verif.items():
        print(f"  {clave}: {valor}")
    
    # Ejemplo de destinatarios
    print("\nEjemplo de destinatarios para sección 'maf':")
    destinatarios = email_service.obtener_destinatarios_seccion('maf')
    for d in destinatarios:
        print(f"  - Email: {d['email']}, Nombre: {d['nombre']}")