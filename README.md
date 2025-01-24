# 1. Activar el entorno virtual
# En macOS/Linux:
source venv/bin/activate

# En Windows:
venv\Scripts\activate

# 2. Instalar las dependencias (si no están instaladas):
pip install -r requirements.txt

# 3. Ejecutar el script
python scripts/extract.py

# 4. Revisar los logs (opcional):
# En macOS/Linux:
cat extraction.log

# En Windows:
type extraction.log