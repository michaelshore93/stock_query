# stock_data_api

Uses a REST API to query Quandl's database for select stocks (the ones that are not premium-restricted), writes the data to a SQLite database, and generates charts in Excel. You must use Python 3. The free API only gives you a limited date range. If you want to use this API for professional/academic purposes, you will need to purchase the premium API here: https://www.quandl.com/data/EOD-End-of-Day-US-Stock-Prices.


My recommendation is to type the following into Terminal before running the program (name venv whatever you want or use no virtual environment at all):

virtualenv --no-site-packages venv

source venv/bin/activate

pip install -r requirements.txt
