# Implementación del Cálculo de Tendencia de Ventas

## Resumen de Cambios

Este documento describe la implementación del sistema de cálculo de tendencia de ventas, que permite comparar las ventas objetivo de la semana anterior con las ventas reales actuales para detectar tendencias de aumento o disminución de demanda.

---

## Archivos Modificados

### 1. `src/correction_data_loader.py`

Se añadieron las siguientes funciones al final del archivo:

#### `encontrar_archivo_semana_anterior(directorio_base, semana_actual)`
Busca el archivo de pedido de la semana anterior (semana actual - 1) basándose en el patrón de nombre `Pedido_Semana_NN_*.xlsx`.

**Retorna:** Ruta del archivo encontrado o `None` si no existe.

#### `leer_archivo_ventas_reales(directorio_entrada)`
Lee el archivo `SPA_ventas_reales.xlsx` del ERP. Este archivo tiene nombre fijo porque el ERP no permite nombres dinámicos.

**Comportamiento especial:** Si el archivo no existe, genera una **ADVERTENCIA** (no un error) y continúa con el proceso. Esto permite que el sistema genere pedidos incluso sin datos de ventas reales.

**Retorna:** Tupla `(DataFrame, Boolean)` donde el Boolean indica si el archivo existía.

#### `leer_archivo_stock_actual(directorio_entrada)`
Lee el archivo `SPA_stock_actual.xlsx` del ERP.

**Retorna:** DataFrame con el stock actual o `None` si hay error.

#### `normalizar_datos_historicos(df_pedido_anterior)`
Normaliza el DataFrame del pedido de la semana anterior para extraer solo las columnas necesarias para el cálculo de tendencia.

**Retorna:** DataFrame con columnas `Clave_Articulo` y `Ventas_Objetivo_Semana_Pasada`.

#### `fusionar_datos_tendencia(pedidos_df, df_ventas_reales, df_stock_actual, df_ventas_objetivo_anterior)`
Fusiona los datos históricos con el DataFrame de pedidos actual, añadiendo las 3 nuevas columnas:

- `Ventas_Objetivo_Semana_Pasada`: Ventas objetivo de la semana anterior (del archivo Pedido_Semana)
- `Ventas_Reales`: Ventas reales de la semana actual (del archivo SPA_ventas_reales)
- `Stock_Real`: Stock actual (del archivo SPA_stock_actual)

---

### 2. `main.py`

#### Importaciones añadidas
```python
from src.correction_data_loader import (
    encontrar_archivo_semana_anterior,
    leer_archivo_ventas_reales,
    leer_archivo_stock_actual,
    normalizar_datos_historicos,
    fusionar_datos_tendencia
)
```

#### Carga de datos de tendencia (antes del loop de secciones)
Antes de procesar las secciones, el sistema ahora:

1. Determina los directorios de entrada y salida
2. Lee el archivo `SPA_ventas_reales.xlsx` (con advertencia si no existe)
3. Lee el archivo `SPA_stock_actual.xlsx`
4. Busca y lee el archivo de pedido de la semana anterior
5. Normaliza los datos históricos para obtener `Ventas_Objetivo_Semana_Pasada`

#### Fusión de datos (después de `aplicar_stock_minimo`)
Después de calcular el pedido y aplicar el stock mínimo, se llama a `fusionar_datos_tendencia()` para añadir las 3 nuevas columnas al DataFrame.

---

### 3. `src/order_generator.py`

Se actualizó la generación del archivo Excel para incluir las nuevas columnas:

#### COLUMN_WIDTHS actualizado
```python
COLUMN_WIDTHS = {
    ...
    'Q': 14.00,  # Ventas Obj. Semana Pasada (NUEVO)
    'R': 11.00,  # Ventas Reales
    'S': 11.00,  # Stock Real (NUEVO)
    ...
}
```

#### COLUMN_HEADERS actualizado
```python
COLUMN_HEADERS = [
    ...
    'Pedido Corregido Stock',    # P
    'Ventas Obj. Semana Pasada', # Q (NUEVO)
    'Ventas Reales',             # R
    'Stock Real',                # S (NUEVO)
    'Tendencia Consumo',         # T
    'Pedido Final'               # U
]
```

#### COLUMN_MAPPING actualizado
```python
COLUMN_MAPPING = {
    ...
    'Ventas_Objetivo_Semana_Pasada': 'Ventas Obj. Semana Pasada',
    'Ventas_Reales': 'Ventas Reales',
    'Stock_Real': 'Stock Real',
    ...
}
```

#### Formato numérico actualizado
Se incluyeron las nuevas columnas en el formato numérico entero.

---

## Flujo de Datos

```
┌─────────────────────────────────────────────────────────────────┐
│                    main.py                                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│  1. Cargar configuración                                        │
│         │                                                      │
│         ▼                                                      │
│  2. Cargar datos de entrada                                     │
│     ├── SPA_ventas_reales.xlsx (opcional)                       │
│     ├── SPA_stock_actual.xlsx                                   │
│     └── Pedido_Semana_{semana-1}_*.xlsx                         │
│         │                                                      │
│         ▼                                                      │
│  3. Procesar cada sección                                       │
│     ├── Calcular pedido teórico (forecast_engine)               │
│     ├── Aplicar stock mínimo                                    │
│     └── Fusionar datos de tendencia ◄──── NUEVO                 │
│         │                                                      │
│         ▼                                                      │
│  4. Generar archivo Excel                                       │
│     └── Incluir nuevas columnas                                 │
│         │                                                      │
│         ▼                                                      │
│  5. Envío de notificaciones                                     │
│      (las advertencias se enviarán por email)                   │
└─────────────────────────────────────────────────────────────────┘
```

---

## Ejemplo de Salida

El archivo Excel generado incluirá las siguientes columnas de tendencia:

| Código | Artículo | Talla | Color | ... | Pedido Corregido Stock | Ventas Obj. Semana Pasada | Ventas Reales | Stock Real | Tendencia Consumo | Pedido Final |
|--------|----------|-------|-------|-----|------------------------|---------------------------|---------------|------------|-------------------|--------------|
| ART001 | Producto A | U | UNICA | ... | 26 | 20 | 25 | 1 | 0 | 38 |
| ART002 | Producto B | U | UNICA | ... | 33 | 26 | 20 | 6 | 0 | 33 |

---

## Manejo de Errores

### Archivo SPA_ventas_reales.xlsx no existe
```
ADVERTENCIA: No se encontró el archivo de ventas reales
  Archivo: SPA_ventas_reales.xlsx
  Directorio: data/input
  El sistema continuará generando pedidos sin el dato de ventas reales.
  NOTA: Esta advertencia será enviada por email en una futura actualización.
```

### Archivo de semana anterior no existe
```
No se encontró archivo de pedido para la semana 8
  Patrón buscado: Pedido_Semana_8_*.xlsx
  Directorio: data/output
```

---

## Notas Técnicas

1. **Nombre fijo del archivo:** `SPA_ventas_reales.xlsx` tiene nombre fijo porque el ERP no permite generar nombres dinámicos con la semana actual.

2. **Búsqueda de semana anterior:** El sistema busca el archivo de la semana anterior usando el patrón `Pedido_Semana_{semana-1}_*.xlsx`. Si hay múltiples archivos coincidentes, selecciona el más reciente por fecha de modificación.

3. **Continuidad del servicio:** El sistema está diseñado para continuar generando pedidos incluso si falta el archivo de ventas reales, mostrando solo una advertencia.

4. **Preparación para notificaciones por email:** La estructura de advertencia está preparada para que, en una futura actualización, todas las advertencias se envíen automáticamente por email al responsable.

---

## Próximos Pasos (Futuras Actualizaciones)

1. Implementar sistema de envío de advertencias por email
2. Añadir cálculo automático de tendencia en el forecast_engine
3. Generar alertas cuando se detecten tendencias significativas
4. Añadir historial de tendencias para análisis predictivo
