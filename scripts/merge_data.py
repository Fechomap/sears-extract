import os
import pandas as pd
import logging
from datetime import datetime
import time
from openpyxl import load_workbook

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('merge.log'),
        logging.StreamHandler()
    ]
)

class SearsMerger:
    def __init__(self):
        self.output_file = os.path.join('output', 'sears_extractions.xlsx')
        self.concentrado_file = os.path.join('cruce1', 'Concentrado Sears.xlsx')
        self.backup_dir = os.path.join('cruce1', 'backups')
        self.report_file = os.path.join('cruce1', 'reporte_merge.xlsx')
        
    def create_backup(self):
        """Crea una copia de respaldo del archivo concentrado antes de modificarlo"""
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
            
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = os.path.join(self.backup_dir, f'Concentrado_Sears_backup_{timestamp}.xlsx')
        
        if os.path.exists(self.concentrado_file):
            # Copiar el archivo preservando el formato
            wb = load_workbook(self.concentrado_file)
            wb.save(backup_file)
            logging.info(f"Backup creado: {backup_file}")

    def process_duplicates(self, extractions_df):
        """Procesa los pedidos duplicados, sumando sus totales"""
        # Identificar duplicados
        duplicados = extractions_df[extractions_df.duplicated(['Numero_Pedido'], keep=False)]
        pedidos_duplicados = duplicados['Numero_Pedido'].unique()
        
        # Diccionario para almacenar los resultados procesados
        processed_data = {}
        
        for pedido in pedidos_duplicados:
            registros = extractions_df[extractions_df['Numero_Pedido'] == pedido]
            
            # Pequeño delay para estabilidad
            time.sleep(0.1)
            
            # Sumar los totales
            total_sumado = registros['Total'].sum()
            
            # Tomar los datos del primer registro
            primer_registro = registros.iloc[0].to_dict()
            
            # Actualizar el total y agregar nota sobre la suma
            primer_registro['Total'] = total_sumado
            primer_registro['documentos_sumados'] = ', '.join(registros['Numero_Documento'].astype(str))
            
            # Guardar en el diccionario de procesados
            processed_data[pedido] = primer_registro
            
            logging.info(f"""
            Pedido duplicado procesado: {pedido}
            Documentos sumados: {primer_registro['documentos_sumados']}
            Total sumado: {total_sumado}
            """)
        
        return processed_data, pedidos_duplicados

    def update_concentrado_cell(self, wb, sheet_name, row_idx, col_name, value):
        """Actualiza una celda específica preservando el formato"""
        ws = wb[sheet_name]
        # Encontrar el índice de la columna
        header_row = 1  # Asumiendo que el encabezado está en la primera fila
        for idx, cell in enumerate(ws[header_row], 1):
            if cell.value == col_name:
                col_idx = idx
                break
        else:
            return
        
        # Preservar el formato existente
        target_cell = ws.cell(row=row_idx, column=col_idx)
        old_format = target_cell._style
        target_cell.value = value
        target_cell._style = old_format
        
    def merge_data(self):
        try:
            # Crear backup antes de comenzar
            self.create_backup()
            time.sleep(0.5)  # Delay para asegurar que el backup se complete
            
            # Leer el archivo de extracciones
            logging.info("Leyendo archivo de extracciones...")
            extractions_df = pd.read_excel(self.output_file)
            
            # Leer el archivo concentrado preservando el formato
            logging.info("Leyendo archivo concentrado...")
            wb = load_workbook(self.concentrado_file)
            ws = wb.active
            concentrado_df = pd.read_excel(self.concentrado_file)
            
            # Asegurar tipos de datos correctos
            extractions_df['Numero_Pedido'] = extractions_df['Numero_Pedido'].astype(str)
            concentrado_df['ORDEN SEARS '] = concentrado_df['ORDEN SEARS '].astype(str)
            
            # Procesar duplicados
            processed_duplicates, pedidos_duplicados = self.process_duplicates(extractions_df)
            
            # Contador para seguimiento
            updates = 0
            no_matches = 0
            
            # Iterar sobre las filas del archivo de extracciones
            for idx, row in extractions_df.iterrows():
                pedido = row['Numero_Pedido']
                
                time.sleep(0.05)  # Pequeño delay entre actualizaciones
                
                # Si es un duplicado y ya lo procesamos, saltarlo
                if pedido in pedidos_duplicados and pedido not in processed_duplicates:
                    continue
                
                # Buscar coincidencia en el archivo concentrado
                mask = concentrado_df['ORDEN SEARS '] == pedido
                
                if mask.any():
                    row_idx = mask.idxmax() + 2  # +2 porque Excel usa 1-based indexing y tiene encabezado
                    
                    if pedido in processed_duplicates:
                        datos = processed_duplicates[pedido]
                        # Actualizar campos preservando formato
                        self.update_concentrado_cell(wb, ws.title, row_idx, 'Total', datos['Total'])
                        self.update_concentrado_cell(wb, ws.title, row_idx, 'OBSERVACIONES ', 
                                                  f"SUMA DE PRODUCTOS - Documentos: {datos['documentos_sumados']}")
                        
                        updates += 1
                        logging.info(f"""
                        Actualizado pedido duplicado: {pedido}
                        Total sumado: {datos['Total']}
                        Documentos: {datos['documentos_sumados']}
                        """)
                    else:
                        # Actualizar campos preservando formato
                        for campo in ['Total', 'Fecha_Pedido', 'Fecha_Vencimiento', 
                                    'Numero_Documento', 'Tipo_Docto', 'Descripcion', 
                                    'Cheque', 'Proveedor']:
                            self.update_concentrado_cell(wb, ws.title, row_idx, campo, row[campo])
                        updates += 1
                else:
                    no_matches += 1
                    logging.warning(f"No se encontró coincidencia para el pedido: {pedido}")
            
            # Guardar el archivo actualizado preservando formato
            logging.info("Guardando archivo actualizado...")
            wb.save(self.concentrado_file)
            
            # Resumen del proceso
            logging.info(f"""
            Resumen del proceso de merge:
            - Total de registros procesados: {len(extractions_df)}
            - Registros actualizados: {updates}
            - Registros sin coincidencia: {no_matches}
            """)
            
            logging.info("Proceso de merge completado exitosamente")
            
        except Exception as e:
            logging.error(f"Error durante el proceso de merge: {str(e)}")
            raise

if __name__ == "__main__":
    merger = SearsMerger()
    merger.merge_data()