#!/usr/bin/env python3
"""
Script para limpiar rutas absolutas de múltiples archivos en el proyecto
"""

import json
import re
import os

def clean_file_content(content, file_path):
    """Limpia rutas absolutas del contenido de un archivo"""
    changes = 0
    
    # Patrón para rutas absolutas de Windows
    absolute_path_pattern = r'[A-Z]:\\[^"\s]+'
    
    # Reemplazar con ruta relativa genérica
    def replace_path(match):
        nonlocal changes
        original = match.group(0)
        # Extraer nombre de archivo
        filename = original.split('\\')[-1]
        changes += 1
        return f"./data/output/{filename}"
    
    new_content = re.sub(absolute_path_pattern, replace_path, content)
    return new_content, changes

def clean_json_file(file_path):
    """Limpia rutas absolutas en archivos JSON"""
    print(f"\nLimpiando {file_path}...")
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Buscar rutas absolutas
    matches = re.findall(r'[A-Z]:\\[^"\s]+', content)
    
    if not matches:
        print(f"  ✓ No hay rutas absolutas")
        return 0
    
    # Intentar cargar como JSON
    try:
        data = json.loads(content)
        changes = clean_json_recursive(data)
        
        # Guardar archivo limpio
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        
        print(f"  ✓ Limpiadas {changes} rutas absolutas")
        return changes
    except json.JSONDecodeError:
        # Si no es JSON válido, intentar limpiar como texto
        new_content, changes = clean_file_content(content, file_path)
        if changes > 0:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print(f"  ✓ Limpiadas {changes} rutas absolutas")
        return changes

def clean_json_recursive(obj):
    """Limpia rutas absolutas en un objeto JSON recursivamente"""
    changes = 0
    
    if isinstance(obj, dict):
        for key, value in obj.items():
            if isinstance(value, str):
                # Si es una ruta absoluta de Windows
                if re.match(r'^[A-Z]:\\', value):
                    # Determinar tipo de archivo por el nombre
                    filename = value.split('\\')[-1]
                    if 'Pedido_Semana' in filename:
                        obj[key] = f"./data/output/Pedidos_semanales/{filename}"
                    elif 'Resumen_Pedidos' in filename:
                        obj[key] = f"./data/output/Pedidos_semanales_resumen/{filename}"
                    elif 'INFORME_FINAL' in filename:
                        obj[key] = f"./data/output/Informes/{filename}"
                    elif 'PRESENTACION' in filename:
                        obj[key] = f"./data/output/Presentaciones/{filename}"
                    elif 'Articulos_no_comprados' in filename:
                        obj[key] = f"./data/output/Articulos_no_comprados/{filename}"
                    elif 'Compras_sin_autorizacion' in filename:
                        obj[key] = f"./data/output/Compras_sin_autorizacion/{filename}"
                    elif 'Comparacion' in filename:
                        obj[key] = f"./data/output/Comparacion_categoria_C_y_D/{filename}"
                    elif 'Analisis_Categorias' in filename:
                        obj[key] = f"./data/output/Analisis_categoria_C_y_D/{filename}"
                    else:
                        obj[key] = f"./data/output/{filename}"
                    changes += 1
            else:
                changes += clean_json_recursive(value)
    elif isinstance(obj, list):
        for item in obj:
            changes += clean_json_recursive(item)
    
    return changes

def clean_log_file(file_path):
    """Limpia rutas absolutas en archivos de log"""
    print(f"\nLimpiando {file_path}...")
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Buscar rutas absolutas
    matches = re.findall(r'[A-Z]:\\[^"\s]+', content)
    
    if not matches:
        print(f"  ✓ No hay rutas absolutas")
        return 0
    
    # Reemplazar rutas absolutas en stack traces
    # Patrón para rutas en stack traces de Python
    new_content = content
    
    # Reemplazar rutas E:\ y D:\ en stack traces
    for match in set(re.findall(r'[A-Z]:\\[^"\s]+', content)):
        filename = match.split('\\')[-1]
        # Reemplazar solo la ruta, mantener el resto de la línea
        new_content = new_content.replace(match, f"./{filename}")
    
    changes = len(set(re.findall(r'[A-Z]:\\[^"\s]+', content)))
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"  ✓ Limpiadas {changes} rutas absolutas")
    return changes

# Limpiar archivos
total_changes = 0

# 1. state.json.backup
total_changes += clean_json_file('/workspace/SPA/data/state.json.backup')

# 2. logs/sistema.log
total_changes += clean_log_file('/workspace/SPA/logs/sistema.log')

print(f"\n{'='*50}")
print(f"✅ Limpieza completada. Total de cambios: {total_changes}")
