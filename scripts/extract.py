import os
import pandas as pd
import pdfplumber
import logging

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

    def extract_data_from_pdf(self, pdf_path):
        logging.info(f"Procesando archivo: {pdf_path}")
        try:
            with pdfplumber.open(pdf_path) as pdf:
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                lines = text.split('\n')
                data_lines = []
                
                # Extraer información de cheque y proveedor
                cheque_info = ""
                proveedor_info = ""
                for line in lines:
                    if 'Cheque' in line:
                        cheque_info = line.split(':')[1].strip() if ':' in line else ""
                    elif 'Proveedor' in line:
                        proveedor_info = line.split(':')[1].strip() if ':' in line else ""
                
                for line in lines:
                    parts = line.split()
                    if len(parts) > 0 and parts[0].isdigit() and len(parts[0]) == 8:
                        data_lines.append(line)

                for line in data_lines:
                    parts = line.split()
                    if len(parts) >= 6:
                        order_data = {
                            'Numero_Pedido': parts[0],
                            'Fecha_Pedido': parts[1],
                            'Fecha_Vencimiento': parts[2],
                            'Numero_Documento': parts[3],
                            'Tipo_Docto': parts[4],
                            'Total': parts[5].replace('$', '').replace(',', ''),
                            'Descripcion': self.doc_types.get(parts[4], 'OTRO'),
                            'Cheque': cheque_info,
                            'Proveedor': proveedor_info
                        }
                        self.processed_data.append(order_data)
                
        except Exception as e:
            logging.error(f"Error procesando {pdf_path}: {str(e)}")

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
        new_df['Total'] = pd.to_numeric(new_df['Total'], errors='coerce')
        # Convertir Numero_Pedido y Numero_Documento a numérico
        new_df['Numero_Pedido'] = pd.to_numeric(new_df['Numero_Pedido'], errors='coerce')
        new_df['Numero_Documento'] = pd.to_numeric(new_df['Numero_Documento'], errors='coerce')

        # Combinar los datos existentes con los nuevos
        if not existing_df.empty:
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        else:
            combined_df = new_df

        # Eliminar registros duplicados (usando como clave Numero_Pedido y Numero_Documento)
        combined_df = combined_df.drop_duplicates(subset=['Numero_Pedido', 'Numero_Documento'], keep='first')

        # Generar análisis basado en el DataFrame combinado
        analysis_df = self.generate_analysis_from_df(combined_df)

        # Crear Excel con múltiples hojas (se sobrescribe el archivo acumulado)
        with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
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