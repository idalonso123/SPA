#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para generar informe de artículos comprados fuera del pedido planificado.

Este script analiza los artículos que:
- No estaban en el stock inicial ni en el pedido de compra, pero están en el stock actual
- Estaban en el stock inicial, no estaban en el pedido, pero tienen más unidades en stock actual

Incluye registro histórico para evitar duplicados en semanas posteriores.

Autor: MiniMax Agent
Fecha: 2026-02-18
"""

import pandas as pd
import os
import glob
import json
from datetime import datetime
import argparse
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Configuración de rutas
# Usar el directorio del script como base para rutas absolutas
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_DATA_INPUT = os.path.join(SCRIPT_DIR, "data", "input")
RUTA_DATA_OUTPUT = os.path.join(SCRIPT_DIR, "data", "output")
RUTA_DATA = os.path.join(SCRIPT_DIR, "data")
ARCHIVO_HISTORICO = "compras_sin_pedido_historico.json"


def normalizar_clave(articulo, talla, color):
    """
    Normaliza una clave para que coincida entre archivos.
    """
    articulo = str(articulo)
    try:
        articulo_num = float(articulo)
        if not pd.isna(articulo_num):
            articulo = str(int(articulo_num))
    except:
        pass
    
    if pd.notna(talla):
        talla = str(talla).strip()
    else:
        talla = ""
    
    if pd.notna(color):
        color = str(color).strip()
    else:
        color = ""
    
    return f"{articulo}_{talla}_{color}"


def cargar_historico():
    """
    Carga el archivo histórico de compras sin pedido.
    
    Returns:
        Diccionario con el histórico de artículos por sección
    """
    ruta_historico = os.path.join(RUTA_DATA, ARCHIVO_HISTORICO)
    
    if os.path.exists(ruta_historico):
        try:
            with open(ruta_historico, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"  Warning: Error al cargar histórico: {e}")
            return {}
    return {}


def guardar_historico(historico):
    """
    Guarda el archivo histórico de compras sin pedido.
    
    Args:
        historico: Diccionario con el histórico de artículos por sección
    """
    ruta_historico = os.path.join(RUTA_DATA, ARCHIVO_HISTORICO)
    
    try:
        with open(ruta_historico, 'w', encoding='utf-8') as f:
            json.dump(historico, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"  Error al guardar histórico: {e}")


def cargar_clasificacion_abc(seccion):
    """
    Carga los artículos de una sección desde el archivo de clasificación ABC.
    """
    secciones_archivo = {
        "maf": "MAF",
        "deco_interior": "DECO_INTERIOR",
        "deco_exterior": "DECO_EXTERIOR",
        "semillas": "SEMILLAS",
        "mascotas_vivo": "MASCOTAS_VIVO",
        "mascotas_manufacturado": "MASCOTAS_MANUFACTURADO",
        "interior": "INTERIOR",
        "fitos": "FITOS",
        "vivero": "VIVERO",
        "utiles_jardin": "UTILES_JARDIN",
        "tierras_aridos": "TIERRA_ARIDOS"
    }
    
    nombre_archivo = secciones_archivo.get(seccion, seccion.upper())
    ruta_archivo = os.path.join(RUTA_DATA_INPUT, f"CLASIFICACION_ABC+D_{nombre_archivo}_ANUAL_2025.xlsx")
    
    try:
        df = pd.read_excel(ruta_archivo)
        df = df[['Artículo', 'Talla', 'Color']].copy()
        df = df.dropna(subset=['Artículo', 'Talla', 'Color'])
        df['clave'] = df.apply(lambda row: normalizar_clave(row['Artículo'], row['Talla'], row['Color']), axis=1)
        return set(df['clave'].unique())
    except Exception as e:
        print(f"  Warning: No se pudo cargar clasificación ABC para {seccion}: {e}")
        return None


def cargar_stock_periodo(ruta_stock, periodo):
    """
    Carga un archivo de stock de un periodo específico.
    """
    try:
        df = pd.read_excel(ruta_stock)
        df = df[df['Tipo registro'] == 'Detalle'].copy()
        df = df[['Artículo', 'Nombre artículo', 'Talla', 'Color', 'Unidades']].copy()
        df = df.dropna(subset=['Artículo', 'Talla', 'Color'])
        df['clave'] = df.apply(lambda row: normalizar_clave(row['Artículo'], row['Talla'], row['Color']), axis=1)
        return df
    except Exception as e:
        print(f"Error al cargar {ruta_stock}: {e}")
        return pd.DataFrame()


def cargar_stock_actual(ruta_stock_actual):
    """
    Carga el archivo de stock actual.
    """
    try:
        df = pd.read_excel(ruta_stock_actual)
        df = df[df['Tipo registro'] == 'Detalle'].copy()
        df = df[['Artículo', 'Nombre artículo', 'Talla', 'Color', 'Unidades']].copy()
        df = df.dropna(subset=['Artículo', 'Talla', 'Color'])
        df['clave'] = df.apply(lambda row: normalizar_clave(row['Artículo'], row['Talla'], row['Color']), axis=1)
        return df
    except Exception as e:
        print(f"Error al cargar {ruta_stock_actual}: {e}")
        return pd.DataFrame()


def cargar_pedido_semana(ruta_pedido):
    """
    Carga un archivo de pedido de semana.
    """
    try:
        df = pd.read_excel(ruta_pedido, header=1)
        df = df[['Código artículo', 'Talla', 'Color']].copy()
        df = df.rename(columns={'Código artículo': 'Artículo'})
        df = df.dropna(subset=['Artículo', 'Talla', 'Color'])
        df['clave'] = df.apply(lambda row: normalizar_clave(row['Artículo'], row['Talla'], row['Color']), axis=1)
        return df
    except Exception as e:
        print(f"Error al cargar {ruta_pedido}: {e}")
        return pd.DataFrame()


def buscar_pedido_semana(seccion, semana=None, directorio=None):
    """
    Busca el archivo de pedido de semana más reciente para una sección.
    """
    if directorio is None:
        directorio = RUTA_DATA_OUTPUT
    
    patron = os.path.join(directorio, f"Pedido_Semana*_{seccion}.xlsx")
    archivos = glob.glob(patron)
    
    if not archivos:
        return None
    
    if semana is not None:
        for archivo in archivos:
            if f"Semana_{semana}_" in archivo:
                return archivo
    
    archivos_ordenados = sorted(archivos, key=os.path.getmtime, reverse=True)
    return archivos_ordenados[0] if archivos_ordenados else None


def obtener_secciones():
    """
    Obtiene la lista de secciones activas del sistema.
    """
    return [
        "interior",
        "maf",
        "deco_interior",
        "deco_exterior",
        "semillas",
        "mascotas_vivo",
        "mascotas_manufacturado",
        "fitos",
        "vivero",
        "utiles_jardin",
        "tierras_aridos"
    ]


def analizar_compras_no_planificadas(seccion, semana=None, historico=None):
    """
    Analiza los artículos comprados sin estar en el pedido para una sección.
    
    Args:
        seccion: Nombre de la sección
        semana: Número de semana del pedido (opcional)
        historico: Diccionario con artículos ya detectados en semanas anteriores
    
    Returns:
        Tupla (DataFrame con resultados, conjunto de claves nuevas detectadas)
    """
    print(f"\n=== Analizando sección: {seccion} ===")
    
    # Obtener claves ya detectadas en el histórico
    claves_historico = set()
    if historico and seccion in historico:
        claves_historico = set(historico[seccion].keys())
    print(f"  Artículos en histórico: {len(claves_historico)}")
    
    # Cargar artículos de la sección desde clasificación ABC
    claves_seccion = cargar_clasificacion_abc(seccion)
    if claves_seccion is None:
        print(f"  Error: No se pudo obtener la lista de artículos para la sección {seccion}")
        return pd.DataFrame(), set(), set()
    print(f"  Artículos en sección (ABC): {len(claves_seccion)}")
    
    # Cargar stock P1 (periodo inicial)
    ruta_stock_p1 = os.path.join(RUTA_DATA_INPUT, "SPA_stock_P1.xlsx")
    stock_p1 = cargar_stock_periodo(ruta_stock_p1, "P1")
    stock_p1 = stock_p1[stock_p1['clave'].isin(claves_seccion)]
    print(f"  Stock P1 (sección): {len(stock_p1)} artículos")
    
    # Cargar stock actual
    ruta_stock_actual = os.path.join(RUTA_DATA_INPUT, "SPA_stock_actual.xlsx")
    stock_actual = cargar_stock_actual(ruta_stock_actual)
    stock_actual = stock_actual[stock_actual['clave'].isin(claves_seccion)]
    print(f"  Stock actual (sección): {len(stock_actual)} artículos")
    
    # Obtener claves que ya NO están en stock actual (consumidos)
    claves_actual = set(stock_actual['clave'].unique())
    claves_consumidas = claves_historico - claves_actual
    if claves_consumidas:
        print(f"  Artículos consumidos (se eliminarán del histórico): {len(claves_consumidas)}")
    
    # Cargar pedido de semana
    ruta_pedido = buscar_pedido_semana(seccion, semana)
    if ruta_pedido:
        pedido = cargar_pedido_semana(ruta_pedido)
        print(f"  Pedido semana: {len(pedido)} artículos")
        print(f"  Archivo: {os.path.basename(ruta_pedido)}")
    else:
        pedido = pd.DataFrame()
        print(f"  Pedido semana: No encontrado")
    
    # Obtener conjuntos de claves
    claves_p1 = set(stock_p1['clave'].unique())
    claves_pedido = set(pedido['clave'].unique()) if not pedido.empty else set()
    
    # Crear diccionarios
    stock_actual_dict = stock_actual.set_index('clave')['Unidades'].to_dict()
    stock_p1_dict = stock_p1.set_index('clave')['Unidades'].to_dict()
    nombre_articulo_dict = stock_actual.set_index('clave')['Nombre artículo'].to_dict()
    
    # Analizar cada artículo en stock actual
    resultados = []
    claves_nuevas = set()  # Nuevasdetections
    
    for clave in claves_actual:
        # Ignorar si ya está en el histórico (ya fue detectado antes)
        if clave in claves_historico:
            continue
        
        en_p1 = clave in claves_p1
        en_pedido = clave in claves_pedido
        unidades_actual = stock_actual_dict.get(clave, 0)
        unidades_p1 = stock_p1_dict.get(clave, 0)
        
        # Determinar si debe incluirse en el informe
        incluir = False
        razon = ""
        
        if not en_p1 and not en_pedido:
            incluir = True
            razon = "Nuevo artículo comprado sin estar en pedido"
        
        elif en_p1 and not en_pedido:
            if unidades_actual > unidades_p1:
                incluir = True
                razon = f"Stock aumentado de {unidades_p1} a {unidades_actual} sin estar en pedido"
            elif unidades_actual == unidades_p1:
                razon = "Stock sin cambios"
            else:
                razon = f"Stock reducido de {unidades_p1} a {unidades_actual} (ventas)"
        
        elif not en_p1 and en_pedido:
            razon = "Artículo en pedido"
        
        else:
            razon = "Artículo en pedido"
        
        if incluir:
            partes = clave.split('_')
            articulo = partes[0] if len(partes) > 0 else ""
            talla = partes[1] if len(partes) > 1 else ""
            color = partes[2] if len(partes) > 2 else ""
            nombre = nombre_articulo_dict.get(clave, "")
            
            resultados.append({
                'Artículo': articulo,
                'Nombre Artículo': nombre,
                'Talla': talla,
                'Color': color,
                'Stock': unidades_actual,
                'Observación': razon
            })
            
            # Añadir a nuevas detecciones
            claves_nuevas.add(clave)
    
    df_resultado = pd.DataFrame(resultados)
    print(f"  Artículos comprados sin estar en pedido (nuevos): {len(df_resultado)}")
    
    return df_resultado, claves_nuevas, claves_consumidas


def aplicar_formato_pedido(ws, df):
    """
    Aplica el formato exacto de los archivos de pedido al Excel.
    """
    # Definir estilos
    header_fill = PatternFill(start_color="008000", end_color="008000", fill_type="solid")  # Verde oscuro
    header_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal='center', vertical='center')
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Fila 1: Números de columna
    num_columnas = 5
    for col in range(1, num_columnas + 1):
        cell = ws.cell(row=1, column=col)
        cell.value = str(col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Fila 2: Encabezados (mismo formato que archivo de pedido)
    encabezados = ['Código artículo', 'Nombre Artículo', 'Talla', 'Color', 'Stock']
    for col, encabezado in enumerate(encabezados, 1):
        cell = ws.cell(row=2, column=col)
        cell.value = encabezado
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Ajustar anchos de columna
    anchos = [15, 40, 12, 12, 10]
    for col, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(col)].width = ancho
    
    # Escribir datos a partir de fila 3
    for idx, row in df.iterrows():
        ws.cell(row=idx+3, column=1, value=row['Artículo'])
        ws.cell(row=idx+3, column=2, value=row['Nombre Artículo'])
        ws.cell(row=idx+3, column=3, value=row['Talla'])
        ws.cell(row=idx+3, column=4, value=row['Color'])
        ws.cell(row=idx+3, column=5, value=row['Stock'])


def generar_informe_compras_no_planificadas(semana=None, archivo_salida=None, resetear_historico=False):
    """
    Genera el informe completo de compras no planificadas para todas las secciones.
    
    Args:
        semana: Número de semana del pedido (opcional)
        archivo_salida: Ruta del archivo de salida (opcional)
        resetear_historico: Si True, borra el histórico antes de generar
    """
    secciones = obtener_secciones()
    
    # Cargar o inicializar histórico
    if resetear_historico:
        print("\n*** REINICIANDO HISTÓRICO ***")
        historico = {}
    else:
        historico = cargar_historico()
        if historico:
            print(f"\nHistórico cargado: {sum(len(v) for v in historico.values())} artículos")
    
    if archivo_salida is None:
        fecha = datetime.now().strftime("%d%m%Y")
        archivo_salida = os.path.join(RUTA_DATA_OUTPUT, f"Compras_Sin_Pedido_{fecha}.xlsx")
    
    # Crear workbook
    wb = Workbook()
    
    # Eliminar la hoja por defecto
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Procesar cada sección
    for seccion in secciones:
        df_resultado, claves_nuevas, claves_consumidas = analizar_compras_no_planificadas(
            seccion, semana, historico
        )
        
        # Actualizar histórico: añadir nuevas detecciones
        if seccion not in historico:
            historico[seccion] = {}
        
        for clave in claves_nuevas:
            historico[seccion][clave] = {
                'fecha_deteccion': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
        
        # Actualizar histórico: eliminar consumidos
        for clave in claves_consumidas:
            if clave in historico[seccion]:
                del historico[seccion][clave]
        
        # Generar Excel
        if not df_resultado.empty:
            df_resultado = df_resultado.sort_values('Artículo')
            df_excel = df_resultado[['Artículo', 'Nombre Artículo', 'Talla', 'Color', 'Stock']].copy()
        else:
            df_excel = pd.DataFrame(columns=['Artículo', 'Nombre Artículo', 'Talla', 'Color', 'Stock'])
        
        # Crear hoja
        ws = wb.create_sheet(title=seccion)
        
        # Aplicar formato
        aplicar_formato_pedido(ws, df_excel)
    
    # Guardar archivo Excel
    wb.save(archivo_salida)
    
    # Guardar histórico actualizado
    guardar_historico(historico)
    
    # Mostrar resumen del histórico
    total_historico = sum(len(v) for v in historico.values())
    print(f"\n=== Resumen histórico ===")
    print(f"Total artículos en histórico: {total_historico}")
    for seccion in secciones:
        if seccion in historico:
            print(f"  {seccion}: {len(historico[seccion])} artículos")
    
    print(f"\n=== Informe generado: {archivo_salida} ===")
    return archivo_salida


def main():
    """Función principal del script."""
    parser = argparse.ArgumentParser(
        description='Genera informe de artículos comprados fuera del pedido planificado.'
    )
    parser.add_argument(
        '--semana', '-s',
        type=int,
        help='Número de semana del pedido (opcional)'
    )
    parser.add_argument(
        '--output', '-o',
        type=str,
        help='Ruta del archivo de salida (opcional)'
    )
    parser.add_argument(
        '--reset', '-r',
        action='store_true',
        help='Reiniciar el histórico (borra todos los registros anteriores)'
    )
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("INFORME DE COMPRAS SIN PEDIDO")
    print("=" * 60)
    print(f"Fecha de generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        archivo = generar_informe_compras_no_planificadas(
            semana=args.semana,
            archivo_salida=args.output,
            resetear_historico=args.reset
        )
        print(f"\n✓ Proceso completado exitosamente")
        print(f"✓ Archivo generado: {archivo}")
        return 0
    except Exception as e:
        print(f"\n✗ Error durante la ejecución: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
