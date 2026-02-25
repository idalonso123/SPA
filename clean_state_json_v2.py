#!/usr/bin/env python3
"""
Script para limpiar rutas absolutas del archivo state.json (versión mejorada)
"""

import json
import re

# Leer el archivo state.json
with open('/workspace/SPA/data/state.json', 'r', encoding='utf-8') as f:
    state_data = json.load(f)

# Contador de cambios
cambios = 0

def clean_path(path):
    """Limpia una ruta absoluta de Windows"""
    global cambios
    
    # Si es "Sin archivo", dejar como está
    if path == "Sin archivo":
        return path
    
    # Si es una ruta absoluta de Windows (D:\...)
    if re.match(r'^[A-Z]:\\', path):
        # Extraer solo el nombre del archivo
        filename = path.split('\\')[-1]
        # Determinar el tipo de archivo basado en el nombre
        if 'Pedido_Semana' in filename:
            cleaned = f"./data/output/Pedidos_semanales/{filename}"
        elif 'Resumen_Pedidos' in filename:
            cleaned = f"./data/output/Pedidos_semanales_resumen/{filename}"
        elif 'INFORME_FINAL' in filename:
            cleaned = f"./data/output/Informes/{filename}"
        elif 'PRESENTACION' in filename:
            cleaned = f"./data/output/Presentaciones/{filename}"
        elif 'Articulos_no_comprados' in filename:
            cleaned = f"./data/output/Articulos_no_comprados/{filename}"
        elif 'Compras_sin_autorizacion' in filename:
            cleaned = f"./data/output/Compras_sin_autorizacion/{filename}"
        elif 'Comparacion' in filename:
            cleaned = f"./data/output/Comparacion_categoria_C_y_D/{filename}"
        elif 'Analisis_Categorias' in filename:
            cleaned = f"./data/output/Analisis_categoria_C_y_D/{filename}"
        else:
            cleaned = f"./data/output/{filename}"
        
        if cleaned != path:
            print(f"  Limpiado: {path[:50]}... -> {cleaned}")
            cambios += 1
        return cleaned
    
    # Normalizar rutas que ya son relativas (.\ o ./)
    if path.startswith('.\\'):
        return path.replace('\\', '/')
    
    return path

# Limpiar sección ejecuciones
if 'ejecuciones' in state_data:
    print("\nLimpiando sección 'ejecuciones'...")
    for exec_entry in state_data['ejecuciones']:
        if 'archivo_generado' in exec_entry and exec_entry['archivo_generado']:
            exec_entry['archivo_generado'] = clean_path(exec_entry['archivo_generado'])

# Limpiar sección pedidos_generados  
if 'pedidos_generados' in state_data:
    print("\nLimpiando sección 'pedidos_generados'...")
    for pedido in state_data['pedidos_generados']:
        if 'archivo' in pedido and pedido['archivo']:
            pedido['archivo'] = clean_path(pedido['archivo'])

# Guardar el archivo limpio
output_path = '/workspace/SPA/data/state.json'
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(state_data, f, indent=4, ensure_ascii=False)

print(f"\n✅ Limpieza completada. Total de cambios: {cambios}")
print(f"✅ Archivo guardado en: {output_path}")
