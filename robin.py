import robin_stocks.robinhood as robin
import os

login = robin.login()
robin.export.export_dividends(f"C:\\Users\\{os.getlogin()}\\OneDrive\\Desktop", file_name=None)