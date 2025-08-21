import pdfplumber
import pandas as pd

def extract_tables_from_pdf(pdf_file, excel_file):

    with pdfplumber.open(pdf_file) as pdf:
        all_tables = []  
        for page_number, page in enumerate(pdf.pages, start=1):

            tables = page.extract_tables()
            for table_index, table in enumerate(tables, start=1):

                df = pd.DataFrame(table)
                
                df.insert(0, "Page", page_number)
                df.insert(1, "Table", table_index)
                all_tables.append(df)

    if all_tables:

        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
            for i, table_df in enumerate(all_tables):
                sheet_name = f"Page{table_df['Page'][0]}_Table{table_df['Table'][0]}"
                table_df.drop(columns=["Page", "Table"], inplace=True)
                table_df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
        print(f"Tables extracted and saved to '{excel_file}' successfully.")
    else:
        print("No tables found in the PDF.")


pdf_path = "target.pdf"  
excel_path = "output.xlsx"  
extract_tables_from_pdf(pdf_path, excel_path)

