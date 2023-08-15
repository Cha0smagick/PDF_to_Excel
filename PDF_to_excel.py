import PyPDF2
import camelot
import tabula
import pandas as pd

def extract_tables_using_camelot(pdf_path):
    # Utiliza Camelot para extraer tablas del PDF usando el modo 'stream'
    # y lee todas las páginas disponibles
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
    return [table.df for table in tables]

def extract_tables_using_tabula(pdf_path):
    # Utiliza Tabula para extraer tablas del PDF leyendo todas las páginas disponibles
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    return tables

def export_tables_to_excel(tables, excel_path):
    # Exporta las tablas extraídas a un archivo Excel usando la biblioteca pandas
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, table_df in enumerate(tables):
            table_df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

if __name__ == "__main__":
    pdf_path = input("Ingrese la ruta del archivo PDF: ")
    excel_path = input("Ingrese la ruta del archivo Excel de salida: ")
    
    extracted_tables = []
    
    try:
        # Intenta extraer tablas utilizando Camelot
        extracted_tables = extract_tables_using_camelot(pdf_path)
    except Exception as e:
        print("Error al utilizar Camelot:", e)
        
        try:
            # Si falla Camelot, intenta extraer tablas utilizando Tabula
            extracted_tables = extract_tables_using_tabula(pdf_path)
        except Exception as e:
            print("Error al utilizar Tabula:", e)
            print("No se pudieron extraer tablas del PDF.")
    
    if extracted_tables:
        # Si se extrajeron tablas con éxito, exporta a un archivo Excel
        export_tables_to_excel(extracted_tables, excel_path)
        print("Extracción y exportación completadas.")
