# stock_data_api

Uses a REST API to query Quandl's database for select stocks (the ones that are not premium-restricted), writes the data to a SQLite database, and generates charts in Excel. You must use Python 3; the code is intended for Mac/Linux and may encounter issues in Windows.


My recommendation is to type the following into Terminal before running the program (name venv whatever you want):

virtualenv --no-site-packages venv

source venv/bin/activate

pip install -r requirements.txt
