#!/usr/bin/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para comparar dos archivos de análisis de categoría C y D.
Compara las métricas de resumen entre el archivo actual y el de la semana pasada.
"""

import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import glob


def extraer_metricas_de_excel(ruta_archivo):
    """
    Extrae las métricas de resumen de un archivo Excel de análisis.
    """
    from openpyxl import load_workbook
    
    metricas = {}
    
    try:
        wb = load_workbook(ruta_archivo, data_only=True)
        
        for nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
            
            # Buscar la sección de métricas
            metricas_seccion = {}
            
            # Buscar la fila donde dice "MÉTRICAS DE RESUMEN"
            fila_metricas = None
            for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if row[0] and "MÉTRICAS DE RESUMEN" in str(row[0]):
                    fila_metricas = i
                    break
            
            if fila_metricas:
                # Leer las métricas
                for i in range(fila_metricas + 1, fila_metricas + 10):
                    row = list(ws.iter_rows(min_row=i, max_row=i, values_only=True))[0]
                    if row[0] and row[1] is not None:
                        clave = str(row[0]).strip()
                        valor = row[1]
                        metricas_seccion[clave] = valor
                
                metricas[nombre_hoja] = metricas_seccion
        
        return metricas
    
    except Exception as e:
        print(f"Error al leer {ruta_archivo}: {e}")
        return None


def comparar_metricas(metricas_actual, metricas_anterior):
    """
    Compara las métricas entre dos semanas.
    """
    comparacion = {}
    
    # Obtener todas las secciones
    todas_secciones = set(metricas_actual.keys()) | set(metricas_anterior.keys())
    
    for seccion in todas_secciones:
        metricas_seccion_actual = metricas_actual.get(seccion, {})
        metricas_seccion_anterior = metricas_anterior.get(seccion, {})
        
        comparacion_seccion = {}
        
        # Obtener todas las métricas
        todas_claves = set(metricas_seccion_actual.keys()) | set(metricas_seccion_anterior.keys())
        
        for clave in todas_claves:
            valor_actual = metricas_seccion_actual.get(clave, 0)
            valor_anterior = metricas_seccion_anterior.get(clave, 0)
            
            # Convertir a número
            try:
                if isinstance(valor_actual, str):
                    valor_actual = float(valor_actual.replace('%', '').replace(',', '.'))
                else:
                    valor_actual = float(valor_actual) if valor_actual is not None else 0
                
                if isinstance(valor_anterior, str):
                    valor_anterior = float(valor_anterior.replace('%', '').replace(',', '.'))
                else:
                    valor_anterior = float(valor_anterior) if valor_anterior is not None else 0
                
                diferencia = valor_actual - valor_anterior
                
                # Determinar si es positivo o negativo
                if diferencia > 0:
                    tendencia = "↑"
                elif diferencia < 0:
                    tendencia = "↓"
                else:
                    tendencia = "="
                
                comparacion_seccion[clave] = {
                    'actual': valor_actual,
                    'anterior': valor_anterior,
                    'diferencia': diferencia,
                    'tendencia': tendencia
                }
            except:
                comparacion_seccion[clave] = {
                    'actual': valor_actual,
                    'anterior': valor_anterior,
                    'diferencia': 0,
                    'tendencia': '='
                }
        
        comparacion[seccion] = comparacion_seccion
    
    return comparacion


def generar_excel_comparacion(comparacion, archivo_salida):
    """
    Genera un archivo Excel con la comparación.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    wb = Workbook()
    
    # Eliminar hoja por defecto
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # Estilos
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="FF008000", end_color="FF008000", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    section_font = Font(bold=True, size=12, color="FF008000")
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Crear hoja de resumen
    ws = wb.create_sheet(title="RESUMEN COMPARATIVO")
    
    # Título
    ws['A1'] = "COMPARACIÓN DE ARTÍCULOS C Y D - EVOLUCIÓN SEMANAL"
    ws['A1'].font = Font(bold=True, size=14, color="FF008000")
    ws.merge_cells('A1:F1')
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    
    ws['A2'] = f"Fecha de comparación: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws['A2'].font = Font(italic=True, size=10)
    ws.merge_cells('A2:F2')
    
    # Encabezados
    headers = ['Sección', 'Métrica', 'Semana Anterior', 'Semana Actual', 'Diferencia', 'Tendencia']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Escribir datos
    fila_actual = 5
    
    for seccion, metricas in comparacion.items():
        # Título de sección
        ws.cell(row=fila_actual, column=1, value=seccion)
        ws.cell(row=fila_actual, column=1).font = section_font
        ws.merge_cells(f'A{fila_actual}:F{fila_actual}')
        ws.cell(row=fila_actual, column=1).fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        fila_actual += 1
        
        for metrica, datos in metricas.items():
            # Convertir diferencia a formato legible
            diferencia = datos['diferencia']
            tendencia = datos['tendencia']
            
            # Formatear valores
            if 'unidades' in metrica.lower() or 'artículos' in metrica.lower():
                actual_str = f"{datos['actual']:.0f}"
                anterior_str = f"{datos['anterior']:.0f}"
                diferencia_str = f"{diferencia:+.0f}"
            elif '%' in metrica:
                actual_str = f"{datos['actual']:.1f}%"
                anterior_str = f"{datos['anterior']:.1f}%"
                diferencia_str = f"{diferencia:+.1f}%"
            else:
                actual_str = str(datos['actual'])
                anterior_str = str(datos['anterior'])
                diferencia_str = str(diferencia)
            
            # Color de tendencia
            if tendencia == "↓":
                color_tendencia = "FF008000"  # Verde (descenso)
            elif tendencia == "↑":
                color_tendencia = "FFFF0000"  # Rojo (aumento)
            else:
                color_tendencia = "FF808080"  # Gris (sin cambio)
            
            ws.cell(row=fila_actual, column=1, value="").border = thin_border
            ws.cell(row=fila_actual, column=2, value=metrica).border = thin_border
            ws.cell(row=fila_actual, column=3, value=anterior_str).border = thin_border
            ws.cell(row=fila_actual, column=4, value=actual_str).border = thin_border
            ws.cell(row=fila_actual, column=5, value=diferencia_str).border = thin_border
            
            cell_tendencia = ws.cell(row=fila_actual, column=6, value=tendencia)
            cell_tendencia.border = thin_border
            cell_tendencia.font = Font(bold=True, color=color_tendencia)
            
            fila_actual += 1
        
        fila_actual += 1  # Espacio entre secciones
    
    # Calcular totales
    ws.cell(row=fila_actual, column=1, value="TOTALES")
    ws.cell(row=fila_actual, column=1).font = section_font
    ws.merge_cells(f'A{fila_actual}:B{fila_actual}')
    ws.cell(row=fila_actual, column=1).fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    
    # Calcular totales
    total_anterior = 0
    total_actual = 0
    
    for seccion, metricas in comparacion.items():
        for metrica, datos in metricas.items():
            if 'artículos' in metrica.lower() and 'en stock' in metrica.lower():
                total_anterior += datos['anterior']
                total_actual += datos['actual']
            elif 'unidades totales' in metrica.lower():
                total_anterior += datos['anterior']
                total_actual += datos['actual']
    
    diferencia_total = total_actual - total_anterior
    tendencia_total = "↓" if diferencia_total < 0 else ("↑" if diferencia_total > 0 else "=")
    
    ws.cell(row=fila_actual, column=3, value=f"{total_anterior:.0f}").border = thin_border
    ws.cell(row=fila_actual, column=4, value=f"{total_actual:.0f}").border = thin_border
    ws.cell(row=fila_actual, column=5, value=f"{diferencia_total:+.0f}").border = thin_border
    
    cell_tendencia_total = ws.cell(row=fila_actual, column=6, value=tendencia_total)
    cell_tendencia_total.font = Font(bold=True, size=12)
    if tendencia_total == "↓":
        cell_tendencia_total.font = Font(bold=True, size=12, color="FF008000")
    elif tendencia_total == "↑":
        cell_tendencia_total.font = Font(bold=True, size=12, color="FFFF0000")
    
    # Ajustar anchos
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 12
    
    # Guardar archivo
    wb.save(archivo_salida)
    return archivo_salida


def comparar_archivos(archivo_actual, archivo_anterior, archivo_salida=None):
    """
    Compara dos archivos de análisis de categoría C y D.
    """
    print("=" * 60)
    print("COMPARANDO ANÁLISIS DE ARTÍCULOS C Y D")
    print("=" * 60)
    
    print(f"\n📊 Archivo actual: {archivo_actual}")
    print(f"📊 Archivo anterior: {archivo_anterior}")
    
    # Extraer métricas
    print("\n📥 Extrayendo métricas...")
    metricas_actual = extraer_metricas_de_excel(archivo_actual)
    metricas_anterior = extraer_metricas_de_excel(archivo_anterior)
    
    if metricas_actual is None or metricas_anterior is None:
        print("❌ Error al leer los archivos")
        return None
    
    # Comparar métricas
    print("📊 Comparando métricas...")
    comparacion = comparar_metricas(metricas_actual, metricas_anterior)
    
    # Generar archivo Excel
    if archivo_salida is None:
        fecha = datetime.now().strftime("%d%m%Y")
        archivo_salida = f"data/output/Comparacion_Categoria_CD_{fecha}.xlsx"
    
    print(f"\n📝 Generando archivo de comparación: {archivo_salida}")
    generar_excel_comparacion(comparacion, archivo_salida)
    
    # Mostrar resumen en consola
    print("\n" + "=" * 60)
    print("RESUMEN DE COMPARACIÓN")
    print("=" * 60)
    
    for seccion, metricas in comparacion.items():
        print(f"\n📦 {seccion}:")
        
        for metrica, datos in metricas.items():
            tendencia = datos['tendencia']
            diferencia = datos['diferencia']
            
            if 'unidades' in metrica.lower() or 'artículos' in metrica.lower():
                print(f"   {metrica}: {datos['anterior']:.0f} → {datos['actual']:.0f} ({diferencia:+.0f}) {tendencia}")
            elif '%' in metrica:
                print(f"   {metrica}: {datos['anterior']:.1f}% → {datos['actual']:.1f}% ({diferencia:+.1f}%) {tendencia}")
    
    # Calcular totales
    print("\n" + "=" * 60)
    print("EVOLUCIÓN TOTAL")
    print("=" * 60)
    
    total_articulos_anterior = 0
    total_articulos_actual = 0
    total_unidades_anterior = 0
    total_unidades_actual = 0
    
    for seccion, metricas in comparacion.items():
        for metrica, datos in metricas.items():
            if 'artículos' in metrica.lower() and 'en stock' in metrica.lower():
                total_articulos_anterior += datos['anterior']
                total_articulos_actual += datos['actual']
            elif 'unidades totales' in metrica.lower():
                total_unidades_anterior += datos['anterior']
                total_unidades_actual += datos['actual']
    
    print(f"\n📊 Total artículos en stock:")
    print(f"   Anterior: {total_articulos_anterior:.0f}")
    print(f"   Actual: {total_articulos_actual:.0f}")
    print(f"   Diferencia: {total_articulos_actual - total_articulos_anterior:+.0f}")
    
    print(f"\n📊 Total unidades en stock:")
    print(f"   Anterior: {total_unidades_anterior:.0f}")
    print(f"   Actual: {total_unidades_actual:.0f}")
    print(f"   Diferencia: {total_unidades_actual - total_unidades_anterior:+.0f}")
    
    if total_articulos_actual < total_articulos_anterior:
        print("\n✅ EVOLUCIÓN POSITIVA:减少 artículos de categoría C y D")
    elif total_articulos_actual > total_articulos_anterior:
        print("\n⚠️ EVOLUCIÓN NEGATIVA: Aumentan artículos de categoría C y D")
    else:
        print("\n➡️ SIN CAMBIOS: Mismo número de artículos")
    
    return archivo_salida


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("Uso: python comparar_analisis_cd.py <archivo_actual> <archivo_anterior> [archivo_salida]")
        print("\nEjemplo:")
        print("  python comparar_analisis_cd.py Analisis_Categoria_CD_20022026.xlsx Analisis_Categoria_CD_13022026.xlsx")
        sys.exit(1)
    
    archivo_actual = sys.argv[1]
    archivo_anterior = sys.argv[2]
    archivo_salida = sys.argv[3] if len(sys.argv) > 3 else None
    
    try:
        resultado = comparar_archivos(archivo_actual, archivo_anterior, archivo_salida)
        if resultado:
            print(f"\n✅ Comparación completada: {resultado}")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
