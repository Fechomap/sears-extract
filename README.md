# Extractor y Procesador de Datos Sears

Este proyecto consiste en dos scripts principales que trabajan en conjunto para extraer información de PDFs y procesarla en un archivo Excel concentrador.

## Estructura del Proyecto

```
proyecto/
├── input/            # Carpeta con los PDFs a procesar
├── output/           # Carpeta donde se guarda sears_extractions.xlsx
├── cruce1/          
│   ├── Concentrado Sears.xlsx    # Archivo concentrador principal
│   ├── reporte_merge.xlsx        # Reporte de operaciones de merge
│   └── backups/                  # Respaldos automáticos
└── scripts/         
    ├── extract.py    # Script de extracción de PDFs
    ├── merge_data.py # Script de merge con concentrador
    └── run_all.py    # Script para ejecutar todo el proceso
```

## Pasos de Instalación y Uso

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

### 4. Proceso de Extracción y Merge

#### a. Ejecutar solo la extracción
```bash
python scripts/extract.py
```
Este script procesará los PDFs en la carpeta `input/` y generará el archivo `sears_extractions.xlsx` en la carpeta `output/`.

#### b. Ejecutar solo el merge
```bash
python scripts/merge_data.py
```
Este script tomará los datos de `sears_extractions.xlsx` y los incorporará al archivo `Concentrado Sears.xlsx`.

#### c. Ejecutar el proceso completo
```bash
python scripts/run_all.py
```
Este comando ejecutará ambos scripts en secuencia.

### 5. Revisar los logs (opcional)
En macOS/Linux:
```bash
cat extraction.log
cat merge.log
```

En Windows:
```bash
type extraction.log
type merge.log
```

## Funcionalidades Principales

### Extracción (extract.py)
- Lee PDFs de la carpeta input/
- Extrae información clave como números de pedido, fechas, montos
- Genera un archivo Excel con los datos extraídos
- Mantiene un log detallado del proceso

### Merge (merge_data.py)
- Procesa la información extraída y la incorpora al archivo concentrador
- Maneja casos especiales como pedidos duplicados (suma montos automáticamente)
- Preserva el formato del archivo concentrador
- Genera respaldos automáticos antes de cada operación
- Mantiene un reporte detallado de las operaciones realizadas

### Características del Merge
- Suma automática de pedidos duplicados
- Registro de operaciones en la columna de observaciones
- Preservación de formatos y estilos del Excel
- Sistema de respaldos automáticos
- Proceso gradual con delays para estabilidad
- Logging detallado de todas las operaciones

## Notas Importantes
- El script de merge preserva todos los formatos del archivo concentrador
- Se crean respaldos automáticos antes de cada operación de merge
- Los pedidos duplicados se suman automáticamente y se registra en observaciones
- Se pueden procesar tanto archivos individuales como lotes grandes
- El sistema está diseñado para ser incremental y no destructivo