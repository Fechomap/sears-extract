import os
import pandas as pd
import logging
from datetime import datetime
import time
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('merge_csv.log'),
        logging.StreamHandler()
    ]
)

class SearsCsvMerger:
    def __init__(self):
        self.input_dir = 'input_csv'
        self.concentrado_file = os.path.join('cruce1', 'Concentrado Sears.xlsx')
        self.backup_dir = os.path.join('cruce1', 'backups')
        self.report_file = os.path.join('cruce1', 'reporte_merge_csv.xlsx')
        
        # Mapeo de columnas del CSV a columnas del Excel (comenzando en AB)
        self.column_mapping = {
            'Pedido': 'AB',            # Columna AB
            'Marketplace': 'AC',        # Columna AC
            'Seller': 'AD',            # Columna AD
            'Monto': 'AE',             # Columna AE
            'Nombre_producto': 'AF',    # Columna AF
            'Precio': 'AG',            # Columna AG
            'sku': 'AH',               # Columna AH
            'Estatus_pedido': 'AI',    # Columna AI
            'Estatus_partida': 'AJ',   # Columna AJ
            'Fecha_Pedido': 'AK',      # Columna AK
            'IdFulfillment': 'AL',     # Columna AL
            'NoGuia': 'AM',            # Columna AM
            'Tipo_envio': 'AN'         # Columna AN
        }
        
    def create_backup(self):
        """Crea una copia de respaldo del archivo concentrado antes de modificarlo"""
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
            
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = os.path.join(self.backup_dir, f'Concentrado_Sears_backup_csv_{timestamp}.xlsx')
        
        if os.path.exists(self.concentrado_file):
            wb = load_workbook(self.concentrado_file)
            wb.save(backup_file)
            logging.info(f"Backup creado: {backup_file}")

    def get_cell_value(self, ws, row_idx, col_letter):
        """Obtiene el valor de una celda de manera segura"""
        cell = ws[f"{col_letter}{row_idx}"]
        return cell.value if cell else None

    def update_concentrado_cell(self, wb, sheet_name, row_idx, col_letter, value, csv_col):
        """Actualiza una celda específica preservando el formato"""
        ws = wb[sheet_name]
        
        # Obtener valor actual
        current_value = self.get_cell_value(ws, row_idx, col_letter)
        
        # Convertir ambos valores a string para comparación
        str_current = str(current_value) if current_value is not None else ""
        str_value = str(value) if value is not None else ""
        
        # Verificar si el valor actual está vacío o es diferente
        if current_value is None or str_current.strip() == "" or str_current != str_value:
            target_cell = ws[f"{col_letter}{row_idx}"]
            old_format = target_cell._style
            target_cell.value = value
            target_cell._style = old_format
            return True, f"{csv_col}: {current_value} -> {value}"
            
        return False, None

    def merge_csv_data(self, csv_file):
        try:
            # Crear backup antes de comenzar
            self.create_backup()
            time.sleep(0.5)
            
            csv_filename = os.path.basename(csv_file)
            logging.info(f"Procesando archivo: {csv_filename}")
            
            # Leer el archivo CSV
            csv_df = pd.read_csv(csv_file, encoding='utf-8')
            csv_df['Pedido'] = csv_df['Pedido'].astype(str)
            
            # Leer el archivo concentrado
            logging.info("Leyendo archivo concentrado...")
            wb = load_workbook(self.concentrado_file)
            ws = wb.active
            concentrado_df = pd.read_excel(self.concentrado_file)
            concentrado_df['ORDEN SEARS '] = concentrado_df['ORDEN SEARS '].astype(str)
            
            # Contadores y reporte
            updates = 0
            no_matches = 0
            report_data = {
                'Fecha_Proceso': [],
                'Archivo_Origen': [],
                'Numero_Pedido': [],
                'Estado': [],
                'Detalles': []
            }
            
            # Procesar cada fila del CSV
            for idx, row in csv_df.iterrows():
                time.sleep(0.05)
                pedido = str(row['Pedido'])
                
                # Buscar coincidencia
                mask = concentrado_df['ORDEN SEARS '] == pedido
                
                if mask.any():
                    excel_row_idx = mask.idxmax() + 2
                    changes = []
                    updates_in_row = 0
                    
                    # Verificar y actualizar cada campo
                    for csv_col, excel_col in self.column_mapping.items():
                        try:
                            value = row[csv_col]
                            if pd.notna(value):  # Solo procesar valores no nulos
                                updated, change_msg = self.update_concentrado_cell(
                                    wb, ws.title, excel_row_idx, excel_col, value, csv_col
                                )
                                if updated:
                                    updates_in_row += 1
                                    if change_msg:
                                        changes.append(change_msg)
                        except Exception as e:
                            logging.warning(f"Error en columna {csv_col}, pedido {pedido}: {str(e)}")
                    
                    if updates_in_row > 0:
                        updates += 1
                        logging.info(f"Pedido {pedido}: {updates_in_row} campos actualizados")
                        if changes:
                            logging.info("Cambios: " + ", ".join(changes))
                    
                    # Agregar al reporte
                    report_data['Fecha_Proceso'].append(datetime.now())
                    report_data['Archivo_Origen'].append(csv_filename)
                    report_data['Numero_Pedido'].append(pedido)
                    report_data['Estado'].append('Actualizado' if updates_in_row > 0 else 'Sin cambios')
                    report_data['Detalles'].append(
                        f"{updates_in_row} campos actualizados: {', '.join(changes)}" if changes 
                        else "Datos ya actualizados"
                    )
                else:
                    no_matches += 1
                    logging.warning(f"No se encontró coincidencia para el pedido: {pedido}")
                    
                    # Agregar al reporte
                    report_data['Fecha_Proceso'].append(datetime.now())
                    report_data['Archivo_Origen'].append(csv_filename)
                    report_data['Numero_Pedido'].append(pedido)
                    report_data['Estado'].append('No encontrado')
                    report_data['Detalles'].append('Pedido no existe en Concentrado Sears')
            
            # Guardar archivo actualizado
            logging.info("Guardando archivo actualizado...")
            wb.save(self.concentrado_file)
            
            # Actualizar reporte
            report_df = pd.DataFrame(report_data)
            if os.path.exists(self.report_file):
                existing_report = pd.read_excel(self.report_file)
                report_df = pd.concat([existing_report, report_df], ignore_index=True)
            report_df.to_excel(self.report_file, index=False)
            
            # Resumen
            logging.info(f"""
            Resumen del proceso de merge CSV:
            Archivo: {csv_filename}
            - Total de registros: {len(csv_df)}
            - Registros actualizados: {updates}
            - Sin coincidencia: {no_matches}
            
            El reporte detallado se ha guardado en: {self.report_file}
            """)
            
        except Exception as e:
            logging.error(f"Error durante el proceso de merge CSV: {str(e)}")
            raise

    def process_all_csvs(self):
        """Procesa todos los CSVs en la carpeta input_csv"""
        if not os.path.exists(self.input_dir):
            os.makedirs(self.input_dir)
            logging.info(f"Carpeta {self.input_dir} creada")
            return
        
        csv_files = [f for f in os.listdir(self.input_dir) if f.endswith('.csv')]
        if not csv_files:
            logging.info("No se encontraron archivos CSV para procesar")
            return
        
        for csv_file in sorted(csv_files):
            file_path = os.path.join(self.input_dir, csv_file)
            self.merge_csv_data(file_path)
            time.sleep(1)  # Delay entre archivos

if __name__ == "__main__":
    merger = SearsCsvMerger()
    merger.process_all_csvs()