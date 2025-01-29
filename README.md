# Sistema de Extracción y Procesamiento de Datos Sears

Este proyecto consiste en tres scripts principales que trabajan en conjunto para extraer y procesar información de diferentes fuentes y consolidarla en un archivo Excel concentrador.

## Estructura del Proyecto

```
proyecto/
├── input/            # Carpeta con los PDFs a procesar
├── input_csv/        # Carpeta con los CSVs de reportes
├── output/           # Carpeta donde se guarda sears_extractions.xlsx
├── cruce1/          
│   ├── Concentrado Sears.xlsx    # Archivo concentrador principal
│   ├── reporte_merge.xlsx        # Reporte de operaciones de merge PDF
│   ├── reporte_merge_csv.xlsx    # Reporte de operaciones de merge CSV
│   └── backups/                  # Respaldos automáticos
└── scripts/         
    ├── extract.py                # Script de extracción de PDFs
    ├── merge_data.py            # Script de merge de datos de PDFs
    └── merge_csv_data.py        # Script de merge de datos de CSVs
```

## Pasos de Instalación

### 1. Crear entorno virtual
En macOS/Linux:
```bash
python3.11 -m venv venv
```

En Windows:
```bash
python -m venv venv
```

### 2. Activar el entorno virtual
En macOS/Linux:
```bash
source venv/bin/activate
```

En Windows:
```bash
venv\Scripts\activate
```

### 3. Instalar las dependencias
```bash
pip install -r requirements.txt
```

## Flujo de Trabajo

### 1. Extracción de PDFs (extract.py)
```bash
python scripts/extract.py
```
- Lee PDFs desde la carpeta `input/`
- Extrae información relevante como números de pedido, fechas, montos
- Genera `sears_extractions.xlsx` en la carpeta `output/`
- Mantiene un log detallado en `extraction.log`

### 2. Merge de Datos PDF (merge_data.py)
```bash
python scripts/merge_data.py
```
- Lee datos de `sears_extractions.xlsx`
- Actualiza `Concentrado Sears.xlsx`
- Maneja casos especiales como pedidos duplicados
- Suma montos automáticamente cuando corresponde
- Preserva formatos del Excel
- Genera respaldos automáticos
- Mantiene un log en `merge.log`

### 3. Merge de Datos CSV (merge_csv_data.py)
```bash
python scripts/merge_csv_data.py
```
- Lee CSVs desde la carpeta `input_csv/`
- Actualiza columnas específicas en `Concentrado Sears.xlsx` (a partir de la columna AB)
- Verifica y valida cada campo antes de actualizar
- Preserva formatos del Excel
- Genera respaldos automáticos
- Mantiene un log en `merge_csv.log`

## Características Principales

### Sistema de Respaldos
- Genera respaldos automáticos antes de cada operación
- Mantiene historial de cambios
- Preserva formatos y estilos del Excel

### Manejo de Duplicados
- Detecta pedidos duplicados
- Suma montos automáticamente
- Registra operaciones en observaciones

### Preservación de Datos
- No destructivo, solo acumulativo
- Verifica datos existentes
- Mantiene integridad de la información

### Logging y Reportes
- Logs detallados de cada operación
- Reportes de operaciones realizadas
- Tracking de cambios y actualizaciones

## Notas Importantes

1. **Procesamiento de PDFs**
   - Puede procesar uno o múltiples PDFs
   - Mantiene consistencia en extracciones
   - Detecta y maneja duplicados

2. **Procesamiento de CSVs**
   - Verifica cada campo individualmente
   - No salta archivos procesados previamente
   - Valida datos existentes

3. **Formato del Excel**
   - Preserva todos los formatos existentes
   - Actualiza celda por celda
   - Mantiene integridad del archivo

4. **Seguridad**
   - Sistema de respaldos automáticos
   - Validación de datos
   - Logs detallados para auditoría

## Logs y Monitoreo

Los logs se pueden revisar en:
```bash
# En macOS/Linux:
cat extraction.log
cat merge.log
cat merge_csv.log

# En Windows:
type extraction.log
type merge.log
type merge_csv.log
```

## Mantenimiento

Para limpiar archivos temporales (opcional):
```bash
# En macOS/Linux:
rm -rf output/*
rm *.log

# En Windows:
del output\*
del *.log
```