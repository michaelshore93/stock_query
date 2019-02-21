"""This program uses Quandl's API for Microsoft's stock data, putting that data
    in a JSON format. Then, that JSON is transferred to a SQLite database via
    SQLAlchemy ORM. From there, the data is parsed into Excel, where it is
    utilized to plot charts to provide insights on close price, high/low
    spreads, open/clode deltas, and trading volumes."""

import stock_query

if __name__ == "__main__":
    stock_query.JSON_Connection.connect_json()
    stock_query.DB_Session.json_to_db()
    stock_query.Workbook.closeprice()
    stock_query.Workbook.open_close_delta()
    print("NOTE: All prices are adjusted using CRSP Methodology")
    print("http://www.crsp.com/products/documentation/crsp-calculations")
