import robin_stocks.robinhood as robin
import os
from openpyxl import Workbook, worksheet, load_workbook
from csv import writer
from datetime import datetime

def createXLSX(path, file_name=None):
    if file_name == None:
        file_name = 'myDividends.xlsx'
    workbook = Workbook()
    worksheet = workbook.active
    row_data = [
    'symbol',
    'record_date',
    'payable_date',
    'quantity',
    'withholding',
    'state',
    'amount',
    'rate'
    ]
    worksheet.append(row_data)
    workbook.save(path + "\\" + file_name)
              
@robin.helper.login_required
def export_dividends(dir_path, file_name=None):
    if file_name == None:
        file_name = 'myDividends.xlsx'

    file_path = dir_path + "\\" + file_name
    workbook = load_workbook(filename=file_path)
    worksheet = workbook.active
    all_dividends = robin.get_dividends()
    
    last_row = worksheet.max_row
    for row in range(2, last_row + 1):
        for column in range(1, worksheet.max_column + 1):
           worksheet.cell(row=row, column=column).value = None

    row_index = 2 #start appending at row 2
    for dividend in all_dividends:         
        if(dividend['state'] == 'voided'):  
            continue #don't need voided dividends to be counted
        row_data = [
            robin.get_symbol_by_url(dividend['instrument']),
            datetime.strptime(dividend['record_date'], "%Y-%m-%d"), #as date
            datetime.strptime(dividend['payable_date'], "%Y-%m-%d"), #as date
            float(dividend['position']),
            float(dividend['withholding']),
            dividend['state'],
            float(dividend['amount']),
            float(dividend['rate'])
        ]
        for col, value in enumerate(row_data, start=1):
            worksheet.cell(row=row_index, column=col).value = value
        row_index += 1

    workbook.save(file_path)
        
file_name = input("enterName of your dividend excel doc (including .xlsx) or hit enter for default name: ")
if file_name == "":
    file_name = "myDividends.xlsx"

login = robin.login()

if not os.path.exists(f"C:\\Users\\{os.getlogin()}\\OneDrive\\Desktop\\" + file_name):
    print("Creating Excel Document")
    createXLSX(f"C:\\Users\\{os.getlogin()}\\OneDrive\\Desktop", file_name)

export_dividends(f"C:\\Users\\{os.getlogin()}\\OneDrive\\Desktop", file_name)
