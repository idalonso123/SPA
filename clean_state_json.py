#!/usr/bin/env python3
"""
Script para limpiar rutas absolutas antiguas del archivo state.json
"""

import json
import re

# Leer el archivo state.json
with open('/workspace/SPA/data/state.json', 'r', encoding='utf-8') as f:
    state_data = json.load(f)

# Contador de cambios
cambios = 0

def clean_absolute_paths_in_list(entries, key='archivo'):
    """Limpia rutas absolutas en una lista de entradas"""
    global cambios
    
    for entry in entries:
        if key in entry and entry[key]:
            original = entry[key]
            
            # Si es "Sin archivo", dejar como está
            if original == "Sin archivo":
                continue
            
            # Si es una ruta absoluta de Windows (D:\...)
            if re.match(r'^[A-Z]:\\', original):
                # Extraer solo el nombre del archivo
                filename = original.split('\\')[-1]
                # Convertir a ruta relativa
                entry[key] = f"./data/output/Pedidos_semanales/{filename}"
                cambios += 1
                print(f"  Cambiado: {original[:50]}... -> {entry[key]}")
            
            # Si tiene .\ o .\/, normalizar a formato consistente
            elif original.startswith('.\\') or original.startswith('./'):
                # Normalizar: convertir backslashes a forward slashes y quitar .\ o .\/ 
                cleaned = original.replace('\\', '/').replace('././', './')
                if cleaned != original:
                    entry[key] = cleaned
                    cambios += 1
                    print(f"  Normalizado: {original} -> {cleaned}")

# Limpiar ejecuciones
if 'ejecuciones' in state_data:
    print("\nLimpiando sección 'ejecuciones'...")
    for exec_entry in state_data['ejecuciones']:
        if 'archivo_generado' in exec_entry:
            original = exec_entry['archivo_generado']
            
            # Si es una ruta absoluta de Windows
            if re.match(r'^[A-Z]:\\', original):
                filename = original.split('\\')[-1]
                # Determinar el tipo de archivo basado en el nombre
                if 'Pedido_Semana' in filename:
                    exec_entry['archivo_generado'] = f"./data/output/Pedidos_semanales/{filename}"
                elif 'Resumen_Pedidos' in filename:
                    exec_entry['archivo_generado'] = f"./data/output/Pedidos_semanales_resumen/{filename}"
                elif 'INFORME_FINAL' in filename:
                    exec_entry['archivo_generado'] = f"./data/output/Informes/{filename}"
                elif 'PRESENTACION' in filename:
                    exec_entry['archivo_generado'] = f"./data/output/Presentaciones/{filename}"
                else:
                    exec_entry['archivo_generado'] = f"./data/output/{filename}"
                cambios += 1
                print(f"  Cambiado: {original[:60]}... -> {exec_entry['archivo_generado']}")
            
            # Normalizar rutas que ya son relativas
            elif original.startswith('.\\'):
                exec_entry['archivo_generado'] = original.replace('\\', '/')
                cambios += 1

# Limpiar pedidos_generados
if 'pedidos_generados' in state_data:
    print("\nLimpiando sección 'pedidos_generados'...")
    clean_absolute_paths_in_list(state_data['pedidos_generados'], 'archivo')

# Guardar el archivo limpio
output_path = '/workspace/SPA/data/state.json'
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(state_data, f, indent=4, ensure_ascii=False)

print(f"\n✅ Limpieza completada. Total de cambios: {cambios}")
print(f"✅ Archivo guardado en: {output_path}")
