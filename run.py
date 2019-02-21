"""Module to run to program.

Operates a series of queries from JSON to SQLite to Excel.
"""

import stock_query

if __name__ == "__main__":
    print("""I only have a free API (not paying $35/mo for premium, so \
you can only use my key for certain stocks. \
Apple (AAPL), Microsoft (MSFT), and Goldman Sachs (GS) are some \
that work.\n\nIf you want to check, type \
https://www.quandl.com/api/v3/datasets/EOD/{STOCK TICKER HERE}.json?api_key=EPaRw4wZxHqisMT9gMdX \
into your browser.""")
   stock_query.JSON_Connection.connect_json()
   stock_query.DB_Session.json_to_db()
   stock_query.Workbook.closeprice()
   stock_query.Workbook.open_close_delta()
    print("NOTE: All prices are adjusted using CRSP Methodology")
    print("http://www.crsp.com/products/documentation/crsp-calculations")
