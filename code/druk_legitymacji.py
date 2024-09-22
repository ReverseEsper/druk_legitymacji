import pandas as pd
import xlwings as xw
from datetime import datetime

def get_from_xls(file_path):
    # Odczytanie pliku .xlsx za pomocą pandas
    data = pd.read_excel(file_path)
    return data


def num_to_excel_column(n):
    column = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        column = chr(65 + remainder) + column
    return column




def work_on_excel(template_path,output_path,input_data):
    wb = xw.Book(template_path)
    sheet = wb.sheets['Sheet1']
    # W ten sposób wpisujemy : 
    # sheet.range('B2').value = 'Janus Excelowy'

    dzis = datetime.today()
    formatted_date = dzis.strftime("%d.%m.%Y")

    # Wypełniamy dane od drugiego wiersza (pierwszy wiersz to nagłówki)
    for index, row in input_data.iterrows():
        row_index = 2 + index  # Wiersz, w którym będziemy dodawać dane
        # Uzupełnienie danych w odpowiednich kolumnach
        sheet.range(num_to_excel_column(2)+str(row_index)).value = row['Nazwisko']       
        sheet.range(num_to_excel_column(6)+str(row_index)).value = row['Imię']           
        d_imie = row['Drugie imię']
        if isinstance(d_imie,str):
            sheet.range(num_to_excel_column(8)+str(row_index)).value = d_imie     
        sheet.range(num_to_excel_column(10)+str(row_index)).value = row['Data urodzenia'] 
        sheet.range(num_to_excel_column(11)+str(row_index)).value = row['PESEL']         
        sheet.range(num_to_excel_column(13)+str(row_index)).value = row['Numer w księdze uczniów']         
        sheet.range(num_to_excel_column(15)+str(row_index)).value = formatted_date       
        sheet.range(num_to_excel_column(16)+str(row_index)).value  = "Technikum Samochodowe nr 2"        
        sheet.range(num_to_excel_column(18)+str(row_index)).value  = "im. Czesława Orłowskiego"        
        sheet.range(num_to_excel_column(20)+str(row_index)).value  = "al. Jana Pawła II 69 01-138 Warszawa"        
        sheet.range(num_to_excel_column(22)+str(row_index)).value  = "Piotr Zając"        
        sheet.range(num_to_excel_column(24)+str(row_index)).value = row['Nazwisko']+' '+row['Imię']+' '+row['Dane oddziału']+".jpg"   

    wb.save(output_path)
    wb.close()