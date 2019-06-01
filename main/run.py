"""Module to run to program.

Operates a series of queries from JSON to SQLite to Excel.
"""

import stock_query
import sqlite3

if __name__ == "__main__":
    stock_query.Last_Ran_Check.last_row_pr_key()
    stock_query.JSON_Connection.connect_json()
    stock_query.DB_Session.json_to_db()
    stock_query.Workbook.closeprice()
    stock_query.Workbook.open_close_delta()
    sqlite3.connect('stocks.db').close()
    print('NOTE: All prices are adjusted using CRSP Methodology.')
    print('LINK: http://www.crsp.com/products/documentation/crsp-calculations')
