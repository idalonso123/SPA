#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
date_utils.py
Módulo de utilidades para determinar períodos y años dinámicamente
basándose en la fecha actual del sistema.

Definiciones de períodos (basadas en config/config_comun.json):
- P1: 1 de enero al 28 de febrero
- P2: 1 de marzo al 31 de mayo
- P3: 1 de junio al 31 de agosto
- P4: 1 de septiembre al 31 de diciembre
"""

from datetime import datetime
import json
from pathlib import Path


# ============================================================================
# DEFINICIONES DE PERÍODOS (desde config/config_comun.json)
# ============================================================================

# Períodos definidos por mes
PERIODOS = {
    "P1": {
        "mes_inicio": 1,
        "dia_inicio": 1,
        "mes_fin": 2,
        "dia_fin": 28,
        "nombre": "P1",
        "descripcion": "Enero - Febrero"
    },
    "P2": {
        "mes_inicio": 3,
        "dia_inicio": 1,
        "mes_fin": 5,
        "dia_fin": 31,
        "nombre": "P2",
        "descripcion": "Marzo - Mayo"
    },
    "P3": {
        "mes_inicio": 6,
        "dia_inicio": 1,
        "mes_fin": 8,
        "dia_fin": 31,
        "nombre": "P3",
        "descripcion": "Junio - Agosto"
    },
    "P4": {
        "mes_inicio": 9,
        "dia_inicio": 1,
        "mes_fin": 12,
        "dia_fin": 31,
        "nombre": "P4",
        "descripcion": "Septiembre - Diciembre"
    }
}


def get_semana_del_año(fecha=None):
    """
    Obtiene el número de semana del año para una fecha dada.
    
    Args:
        fecha: Objeto datetime. Si es None, usa la fecha actual.
    
    Returns:
        int: Número de semana (1-53)
    """
    if fecha is None:
        fecha = datetime.now()
    return fecha.isocalendar()[1]


def get_año_actual(fecha=None):
    """
    Obtiene el año actual.
    
    Args:
        fecha: Objeto datetime. Si es None, usa la fecha actual.
    
    Returns:
        int: Año actual
    """
    if fecha is None:
        fecha = datetime.now()
    return fecha.year


def get_periodo_actual(fecha=None):
    """
    Determina el período actual (P1, P2, P3 o P4) basándose en la fecha del sistema.
    
    La lógica es (basada en config/config_comun.json):
    - P1: 1 de enero al 28 de febrero
    - P2: 1 de marzo al 31 de mayo
    - P3: 1 de junio al 31 de agosto
    - P4: 1 de septiembre al 31 de diciembre
    
    Args:
        fecha: Objeto datetime. Si es None, usa la fecha actual.
    
    Returns:
        str: 'P1', 'P2', 'P3' o 'P4'
    """
    if fecha is None:
        fecha = datetime.now()
    
    mes = fecha.month
    
    # Determinar el período según el mes
    if mes <= 2:  # Enero - Febrero
        return "P1"
    elif mes <= 5:  # Marzo - Mayo
        return "P2"
    elif mes <= 8:  # Junio - Agosto
        return "P3"
    else:  # Septiembre - Diciembre
        return "P4"


def get_siguiente_periodo(periodo_actual):
    """
    Obtiene el siguiente período al actual.
    
    Args:
        periodo_actual: 'P1', 'P2', 'P3' o 'P4'
    
    Returns:
        str: Siguiente período en orden secuencial (P1->P2->P3->P4->P1)
    """
    orden_periodos = ["P1", "P2", "P3", "P4"]
    indice_actual = orden_periodos.index(periodo_actual)
    siguiente_indice = (indice_actual + 1) % len(orden_periodos)
    return orden_periodos[siguiente_indice]


def get_año_anterior(año_actual=None):
    """
    Obtiene el año anterior al actual.
    
    Args:
        año_actual: Año actual. Si es None, usa el año de la fecha actual.
    
    Returns:
        int: Año anterior
    """
    if año_actual is None:
        año_actual = get_año_actual()
    return año_actual - 1


def get_periodo_y_año_dinamico(tipo_calculo="siguiente", fecha=None):
    """
    Obtiene el período y año dinámicamente según el tipo de cálculo.
    
    Args:
        tipo_calculo: 
            - "siguiente": siguiente período con año anterior (para PRESENTACION e INFORME)
            - "actual": período actual con año anterior (para analisis_categoria_cd)
        fecha: Objeto datetime. Si es None, usa la fecha actual.
    
    Returns:
        dict: {'periodo': 'P1' o 'P2', 'año': integer}
    """
    if fecha is None:
        fecha = datetime.now()
    
    año_actual = get_año_actual(fecha)
    periodo_actual = get_periodo_actual(fecha)
    
    if tipo_calculo == "siguiente":
        # Para PRESENTACION e INFORME: siguiente período, año anterior
        periodo = get_siguiente_periodo(periodo_actual)
        año = get_año_anterior(año_actual)
    elif tipo_calculo == "actual":
        # Para analisis_categoria_cd: período actual, año anterior
        periodo = periodo_actual
        año = get_año_anterior(año_actual)
    else:
        raise ValueError(f"Tipo de cálculo desconocido: {tipo_calculo}")
    
    return {
        "periodo": periodo,
        "año": año
    }


def get_periodo_info_detallada(fecha=None):
    """
    Obtiene información detallada sobre el período actual y el cálculo.
    Útil para debugging y logs.
    
    Args:
        fecha: Objeto datetime. Si es None, usa la fecha actual.
    
    Returns:
        dict: Información detallada del período
    """
    if fecha is None:
        fecha = datetime.now()
    
    semana = get_semana_del_año(fecha)
    año_actual = get_año_actual(fecha)
    periodo_actual = get_periodo_actual(fecha)
    siguiente_periodo = get_siguiente_periodo(periodo_actual)
    año_anterior = get_año_anterior(año_actual)
    
    return {
        "fecha_actual": fecha.strftime("%Y-%m-%d"),
        "semana_actual": semana,
        "año_actual": año_actual,
        "periodo_actual": periodo_actual,
        "siguiente_periodo": siguiente_periodo,
        "año_anterior": año_anterior,
        # Para PRESENTACION e INFORME
        "siguiente_periodo_año_anterior": {
            "periodo": siguiente_periodo,
            "año": año_anterior
        },
        # Para analisis_categoria_cd
        "periodo_actual_año_anterior": {
            "periodo": periodo_actual,
            "año": año_anterior
        }
    }


def cargar_periodos_desde_config():
    """
    Carga las definiciones de períodos desde el archivo de configuración.
    Si no existe, usa las definiciones por defecto.
    
    Returns:
        dict: Definiciones de períodos
    """
    config_path = Path("config/config.json")
    
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # Verificar si hay definiciones de períodos en el config
            if "periodos" in config:
                return config["periodos"]
        except Exception as e:
            print(f"  AVISO: Error al cargar períodos desde config: {e}")
    
    # Usar definiciones por defecto
    return PERIODOS


# ============================================================================
# FUNCIONES DE PRUEBA
# ============================================================================

if __name__ == "__main__":
    # Prueba de las funciones
    print("=" * 60)
    print("PRUEBA DE FUNCIONES DE DATE_UTILS")
    print("=" * 60)
    
    info = get_periodo_info_detallada()
    
    print(f"\nFecha actual: {info['fecha_actual']}")
    print(f"Semana actual: {info['semana_actual']}")
    print(f"Año actual: {info['año_actual']}")
    print(f"Período actual: {info['periodo_actual']}")
    print(f"Siguiente período: {info['siguiente_periodo']}")
    print(f"Año anterior: {info['año_anterior']}")
    
    print("\n--- Para PRESENTACION.py e INFORME.py (siguiente período) ---")
    datos_siguiente = get_periodo_y_año_dinamico(tipo_calculo="siguiente")
    print(f"  Período: {datos_siguiente['periodo']}, Año: {datos_siguiente['año']}")
    print(f"  Archivo a buscar: CLASIFICACION_ABC+D_*_{datos_siguiente['periodo']}_{datos_siguiente['año']}.xlsx")
    
    print("\n--- Para analisis_categoria_cd.py (período actual) ---")
    datos_actual = get_periodo_y_año_dinamico(tipo_calculo="actual")
    print(f"  Período: {datos_actual['periodo']}, Año: {datos_actual['año']}")
    print(f"  Archivo a buscar: CLASIFICACION_ABC+D_*_{datos_actual['periodo']}_{datos_actual['año']}.xlsx")
