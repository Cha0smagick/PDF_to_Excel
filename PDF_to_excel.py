import PyPDF2
import camelot
import tabula
import pandas as pd

def extract_tables_using_camelot(pdf_path):
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
    return [table.df for table in tables]

def extract_tables_using_tabula(pdf_path):
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    return tables

def export_tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, table_df in enumerate(tables):
            table_df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

if __name__ == "__main__":
    pdf_path = input("Ingrese la ruta del archivo PDF: ")
    excel_path = input("Ingrese la ruta del archivo Excel de salida: ")
    
    method_choice = input("Elija el método de extracción de tablas ('camelot' o 'tabula'): ")
    
    if method_choice == "camelot":
        try:
            extracted_tables = extract_tables_using_camelot(pdf_path)
        except Exception as e:
            print("Error al utilizar Camelot:", e)
            extracted_tables = []
    elif method_choice == "tabula":
        try:
            extracted_tables = extract_tables_using_tabula(pdf_path)
        except Exception as e:
            print("Error al utilizar Tabula:", e)
            extracted_tables = []
    else:
        print("Método no reconocido. No se pudieron extraer tablas del PDF.")
        extracted_tables = []
    
    if extracted_tables:
        export_tables_to_excel(extracted_tables, excel_path)
        print("Extracción y exportación completadas.")

