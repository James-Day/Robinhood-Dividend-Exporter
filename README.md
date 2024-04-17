This Python project leverages the Robinhood API through the robin_stocks library to retrieve dividend information. The program prompts the user for their Robinhood username and password, as well as a filename for the Excel document where the dividend data will be exported.

Using the obtained dividend information, the program utilizes the openpyxl library to export the data to an Excel document. The dependencies required for running the project are listed in the included requirements.txt file, which can be easily installed using the command pip install -r requirements.txt.

This project provides a convenient way to retrieve dividend data from Robinhood and export it to an Excel document for further analysis or record-keeping purposes.


**To run: (only tested for Windows)**

git clone https://github.com/James-Day/Robinhood-Dividend-Exporter.git

cd Robinhood-Dividend-Exporter

pip install -r requirements.txt (best to use a virtual env)

python robin.py

(the excel file should now be on your desktop)
