## Pasos de Instalación y Uso

# 1. Crear entorno virtual
# En macOS/Linux:
python3.11 -m venv venv

# En Windows:
python -m venv venv

# 2. Activar el entorno virtual
# En macOS/Linux:
source venv/bin/activate

# En Windows:
venv\Scripts\activate

# 3. Instalar las dependencias
pip install -r requirements.txt

# 4. Ejecutar el script
python scripts/extract.py

# 5. Revisar los logs (opcional)
# En macOS/Linux:
cat extraction.log

# En Windows:
type extraction.log

# 6. Limpiar y Desactivar (cuando termines)

# Eliminar archivos generados (opcional)
# En macOS/Linux:
rm -rf output/*
rm extraction.log

# En Windows:
del output\*
del extraction.log

# Desactivar el entorno virtual
deactivate

# Eliminar el entorno virtual (opcional)
# En macOS/Linux:
rm -rf venv

# En Windows:
rmdir /s /q venv