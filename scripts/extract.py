import os
import pandas as pd
import pdfplumber
import logging
import traceback
from datetime import datetime

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join('EXCELPDFSEARS', 'extraction.log')),
        logging.StreamHandler()
    ]
)

class SearsExtractor:
    def __init__(self):
        self.output_file = os.path.join('EXCELPDFSEARS', 'sears_extractions.xlsx')
        self.processed_data = []
        # Diccionario para mapear tipos de documento
        self.doc_types = {
            'NP': 'PEDIDO ENTREGADO',
            'NT': 'PEDIDO ENTREGADO',
            'ND': 'DEVOLUCIÓN DE PEDIDO',
            'DR': 'RETENCIÓN ISR E IVA',
            'DV': 'DESCUENTO SERVICIO DE REPARTO'
        }

    def format_date(self, date_str):
        """
        Formatea correctamente una cadena de fecha para asegurar que se reconozca como fecha.
        Maneja diferentes formatos de entrada posibles.
        """
        try:
            # Si ya es un objeto datetime, devolverlo tal cual
            if isinstance(date_str, datetime) or isinstance(date_str, pd.Timestamp):
                return date_str
                
            # Limpia la cadena de fecha
            date_str = date_str.strip()
            
            # Intenta formatear según el patrón DD/MM/YYYY
            try:
                return datetime.strptime(date_str, "%d/%m/%Y")
            except ValueError:
                pass
                
            # Intenta con formato alternativo MM/DD/YYYY
            try:
                return datetime.strptime(date_str, "%m/%d/%Y")
            except ValueError:
                pass
                
            # Intenta con formato YYYY-MM-DD
            try:
                return datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                pass
                
            # Si ninguno funciona, intentar con pandas que es más flexible
            return pd.to_datetime(date_str, errors='coerce')
                
        except Exception as e:
            logging.warning(f"No se pudo formatear la fecha '{date_str}': {str(e)}")
            return date_str  # Devuelve la cadena original si falla

    def extract_data_from_pdf(self, pdf_path):
        logging.info(f"Procesando archivo: {pdf_path}")
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Registrar información sobre páginas totales
                num_pages = len(pdf.pages)
                logging.info(f"Archivo {pdf_path} contiene {num_pages} páginas")
                
                # Variables para acumular datos por archivo
                cheque_global = ""
                proveedor_global = ""
                total_lines_processed = 0
                
                # Primer recorrido para extraer información de cheque y proveedor del documento
                # (tomando el primero que encontremos ya que generalmente es un dato por documento)
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if not page_text:
                        continue
                        
                    lines = page_text.split('\n')
                    for line in lines:
                        if 'Cheque' in line and not cheque_global:
                            cheque_parts = line.split(':')
                            if len(cheque_parts) > 1:
                                cheque_global = cheque_parts[1].strip()
                        elif 'Proveedor' in line and not proveedor_global:
                            proveedor_parts = line.split(':')
                            if len(proveedor_parts) > 1:
                                proveedor_global = proveedor_parts[1].strip()
                            
                    # Si ya encontramos ambos datos, podemos salir del primer escaneo
                    if cheque_global and proveedor_global:
                        break
                
                # Segundo recorrido para extraer líneas de datos de todas las páginas
                for page_num, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if not page_text:
                        logging.warning(f"Página {page_num+1} de {pdf_path} está vacía o no contiene texto extraíble")
                        continue
                    
                    logging.info(f"Procesando página {page_num+1} de {pdf_path}")
                    
                    lines = page_text.split('\n')
                    data_lines = []
                    
                    # Identificar líneas que contienen datos de pedidos (formato de 8 dígitos)
                    for line in lines:
                        parts = line.split()
                        if len(parts) > 0 and parts[0].isdigit() and len(parts[0]) == 8:
                            data_lines.append(line)
                    
                    # Registrar líneas encontradas por página
                    page_lines = len(data_lines)
                    total_lines_processed += page_lines
                    logging.info(f"Encontradas {page_lines} líneas de datos en página {page_num+1} de {pdf_path}")
                    
                    # Procesar cada línea de datos
                    for line in data_lines:
                        parts = line.split()
                        if len(parts) >= 6:
                            try:
                                # Formatear correctamente las fechas para asegurar que se reconozcan
                                fecha_pedido = self.format_date(parts[1])
                                fecha_vencimiento = self.format_date(parts[2])
                                
                                # Normalizar formato de valores numéricos
                                total = parts[5].replace('$', '').replace(',', '')
                                
                                order_data = {
                                    'Numero_Pedido': parts[0],
                                    'Fecha_Pedido': fecha_pedido,
                                    'Fecha_Vencimiento': fecha_vencimiento,
                                    'Numero_Documento': parts[3],
                                    'Tipo_Docto': parts[4],
                                    'Total': total,
                                    'Descripcion': self.doc_types.get(parts[4], 'OTRO'),
                                    'Cheque': cheque_global,
                                    'Proveedor': proveedor_global,
                                    'Pagina_PDF': page_num+1,  # Referencia a la página de origen
                                    'Archivo_PDF': os.path.basename(pdf_path)  # Referencia al archivo
                                }
                                self.processed_data.append(order_data)
                                
                            except Exception as e:
                                logging.error(f"Error procesando línea {line}: {str(e)}")
                
            # Resumen final del procesamiento
            logging.info(f"Finalizado procesamiento de {pdf_path}: {total_lines_processed} líneas en {num_pages} páginas")
                
        except Exception as e:
            logging.error(f"Error procesando {pdf_path}: {str(e)}")
            # Añadir trazabilidad del error
            logging.error(traceback.format_exc())

    def generate_analysis_from_df(self, df):
        """Genera análisis estadístico a partir del DataFrame combinado."""
        if df.empty:
            return None
       
        # Análisis por tipo de documento
        doc_analysis = df.groupby(['Tipo_Docto', 'Descripcion']).agg({
            'Numero_Pedido': 'count',
            'Total': 'sum'
        }).reset_index()
        
        total_docs = doc_analysis['Numero_Pedido'].sum()
        # Convertir a porcentaje (si deseas que sea 0-100, multiplica por 100, o déjalo entre 0 y 1)
        doc_analysis['Porcentaje'] = (doc_analysis['Numero_Pedido'] / total_docs).round(2)
        doc_analysis = doc_analysis.sort_values('Numero_Pedido', ascending=False)
        
        return doc_analysis

    def process_all_pdfs(self):
        input_dir = 'PDFSEARS'
        for filename in os.listdir(input_dir):
            if filename.endswith('.pdf'):
                pdf_path = os.path.join(input_dir, filename)
                self.extract_data_from_pdf(pdf_path)

    def generate_excel(self):
        # Si ya existe el Excel acumulado, lo leemos
        if os.path.exists(self.output_file):
            try:
                existing_df = pd.read_excel(self.output_file, sheet_name='Pedidos')
                logging.info("Archivo Excel existente leído correctamente.")
            except Exception as e:
                logging.error(f"Error al leer el archivo existente: {str(e)}")
                existing_df = pd.DataFrame()
        else:
            existing_df = pd.DataFrame()

        # Crear DataFrame de los nuevos datos procesados
        new_df = pd.DataFrame(self.processed_data)
        
        # Convertir columnas numéricas
        new_df['Total'] = pd.to_numeric(new_df['Total'], errors='coerce')
        new_df['Numero_Pedido'] = pd.to_numeric(new_df['Numero_Pedido'], errors='coerce')
        new_df['Numero_Documento'] = pd.to_numeric(new_df['Numero_Documento'], errors='coerce')
        
        # Convertir columnas de fecha explícitamente
        date_columns = ['Fecha_Pedido', 'Fecha_Vencimiento']
        for col in date_columns:
            # Utilizar pd.to_datetime para convertir las columnas a fecha
            new_df[col] = pd.to_datetime(new_df[col], errors='coerce')
            # Registrar información sobre fechas procesadas
            if not new_df.empty:
                valid_dates = new_df[col].notna().sum()
                total_rows = len(new_df)
                logging.info(f"Columna {col}: {valid_dates} de {total_rows} fechas válidas ({valid_dates/total_rows*100:.1f}%)")
                # Registrar ejemplos de fechas para diagnóstico
                if col in new_df.columns:
                    sample_dates = new_df[col].dropna().head(3).tolist()
                    sample_str = ', '.join([str(d) for d in sample_dates])
                    logging.info(f"Ejemplos de {col}: {sample_str}")

        # Combinar los datos existentes con los nuevos
        if not existing_df.empty:
            # Asegurar que las columnas de fecha también se convierten en el DataFrame existente
            for col in date_columns:
                if col in existing_df.columns:
                    existing_df[col] = pd.to_datetime(existing_df[col], errors='coerce')
            
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        else:
            combined_df = new_df

        # Eliminar registros duplicados (usando como clave Numero_Pedido y Numero_Documento)
        combined_df = combined_df.drop_duplicates(subset=['Numero_Pedido', 'Numero_Documento'], keep='first')

        # Verificar estado final de las fechas antes de escribir
        for col in date_columns:
            if col in combined_df.columns:
                valid_count = combined_df[col].notna().sum()
                total_count = len(combined_df)
                logging.info(f"Final {col}: {valid_count}/{total_count} fechas válidas")

        # Generar análisis basado en el DataFrame combinado
        analysis_df = self.generate_analysis_from_df(combined_df)

        # Crear Excel con múltiples hojas (se sobrescribe el archivo acumulado)
        with pd.ExcelWriter(self.output_file, engine='xlsxwriter', date_format='dd/mm/yyyy') as writer:
            # Hoja de datos principales
            combined_df.to_excel(writer, sheet_name='Pedidos', index=False)
            
            # Hoja de análisis
            if analysis_df is not None:
                analysis_df.to_excel(writer, sheet_name='Análisis', index=False)
            
            # Obtener el libro y las hojas para aplicar formatos
            workbook = writer.book
            worksheet = writer.sheets['Pedidos']
            
            # Formatos
            money_format = workbook.add_format({'num_format': '$#,##0.00'})
            date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            
            # Aplicar formatos a la hoja principal
            worksheet.set_column('A:A', 15)  # Numero Pedido
            worksheet.set_column('B:C', 15, date_format)  # Fechas
            worksheet.set_column('D:D', 15)  # Numero Documento
            worksheet.set_column('E:E', 10)  # Tipo Docto
            worksheet.set_column('F:F', 15, money_format)  # Total
            worksheet.set_column('G:I', 20)  # Descripción, Cheque, Proveedor
            worksheet.set_column('J:K', 15)  # Campos de origen (Página y Archivo)
            
            # Dar formato a la hoja de análisis si existe
            if analysis_df is not None:
                analysis_sheet = writer.sheets['Análisis']
                analysis_sheet.set_column('A:B', 20)  # Tipo y Descripción
                analysis_sheet.set_column('C:C', 15)  # Conteo
                analysis_sheet.set_column('D:D', 15, money_format)  # Total
                analysis_sheet.set_column('E:E', 12)  # Porcentaje
                
                # Agregar formato de porcentaje
                percent_format = workbook.add_format({'num_format': '0.00%'})
                analysis_sheet.set_column('E:E', 12, percent_format)
        
        logging.info(f"Excel generado exitosamente: {self.output_file}")

if __name__ == "__main__":
    extractor = SearsExtractor()
    extractor.process_all_pdfs()
    extractor.generate_excel()