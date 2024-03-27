import robin_stocks.robinhood as robin
import os
from csv import writer

@robin.helper.login_required
def export_dividends(dir_path, file_name=None):
    file_path = robin.export.create_absolute_csv(dir_path, file_name, 'dividends')
    all_dividends = robin.get_dividends()
    with open(file_path, 'w', newline='') as f:
        csv_writer = writer(f)
        csv_writer.writerow([
            'symbol',
            'record_date',
            'payable_date',
            'quantity',
            'withholding',
            'state',
            'amount',
            'rate'
        ])
        for dividend in all_dividends:           
                    csv_writer.writerow([
                        robin.get_symbol_by_url(dividend['instrument']),
                        dividend['record_date'],
                        dividend['payable_date'],
                        dividend['position'],
                        dividend['withholding'],
                        dividend['state'],
                        dividend['amount'],
                        dividend['rate']
                    ])
        f.close()
        

login = robin.login()
export_dividends(f"C:\\Users\\{os.getlogin()}\\OneDrive\\Desktop", file_name=None)

