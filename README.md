# Sistema de Extracción y Procesamiento de Datos Sears

Sistema automatizado para procesar PDFs y CSVs de Sears, consolidando la información en un archivo Excel central.

## Instalación Rápida

```bash
# Crear y activar entorno virtual
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # Unix/Mac

# Instalar dependencias
pip install -r requirements.txt
npm install
```

## Estructura de Carpetas

```
extract-sears/
├── CSVreporte/        # Coloca aquí los archivos CSV a procesar
├── PDFSEARS/         # Coloca aquí los PDFs a procesar
├── EXCELPDFSEARS/    # Archivo de salida de PDFs procesados
├── RESULTADOFINAL/   # Archivo concentrador final y reportes
└── scripts/          # Scripts de procesamiento
```

## Uso

1. **Procesar PDFs:**
   - Coloca los PDFs en la carpeta `PDFSEARS`
   - Ejecuta: `python scripts/extract.py`
   - El resultado se guarda en `EXCELPDFSEARS/sears_extractions.xlsx`

2. **Procesar datos de PDFs:**
   - Ejecuta: `python scripts/merge_data.py`
   - Actualiza el concentrador con datos de PDFs

3. **Procesar CSVs:**
   - Coloca los CSVs en la carpeta `CSVreporte`
   - Ejecuta: `python scripts/merge_csv_data.py`
   - Actualiza el concentrador con datos de CSVs

4. **Ejecutar todo el proceso:**
   ```bash
   node run.js
   ```

## Notas Importantes

- El sistema genera respaldos automáticos antes de cada operación
- Los archivos de log se crean en la carpeta raíz
- Se mantiene registro de todas las operaciones realizadas
- Los archivos duplicados se procesan sumando los montos automáticamente

## Limpieza del Sistema

Para limpiar archivos temporales y cachés:

```bash
# Windows
del /s /q *.log
del /s /q *.pyc

# Unix/Mac
rm -f *.log
find . -name "*.pyc" -delete
```