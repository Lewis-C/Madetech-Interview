# Python script to extract data from excel sheet and load to database file
# Some processing is applied to data to allow easier analysis


import pandas as p
import sqlite3

sql_connection = sqlite3.connect("sales.db")
source = p.read_excel("sales.xlsx", header=1)

# Conversion of customer ID field to integer to remove uneccesary float
source.CustomerID = source.CustomerID.astype("Int64")

# Convert to datetime from string
source.InvoiceDate = p.to_datetime(source.InvoiceDate, yearfirst=True)

# Conversion and splitting of invoice number to seperate prefix character for analysis
source.InvoiceNo = source.InvoiceNo.astype("string")
# Regex finds intitial, single letter (regardless of case) optionally then matches multiple digits
invoice_split = (source.InvoiceNo.str.split(pat="^([A-Za-z]?)(\d+)$", regex=True, expand=True))
source.InvoiceNo = invoice_split[2]
source.insert(1,"InvoicePrefix",value=invoice_split[1])

# Replace native language country name with english for consistency
source.Country = source.Country.replace("EIRE","Ireland")

# Add Invoice Cost, total cost of an invoices items
source.loc[:, 'TotalCost'] = value=source.Quantity * source.UnitPrice

# Sets empty strings to Null
source = source.replace(r'^\s*$', None, regex=True)

source.to_sql(name="SALES", con=sql_connection, if_exists="replace", index=True, index_label="SaleID")

dim_item = source[["StockCode","Description","UnitPrice"]].drop_duplicates()
dim_item.to_sql(name="DIM_ITEM", con=sql_connection, if_exists="replace", index=False)

fact_invoice = source[["StockCode","InvoiceNo","Quantity","TotalCost"]].drop_duplicates()
fact_invoice.to_sql(name="FACT_INVOICE", con=sql_connection, if_exists="replace", index=True, index_label="SaleID")

dim_order = source[["InvoiceNo","InvoicePrefix","CustomerID","Country","InvoiceDate"]].drop_duplicates()
dim_order.to_sql(name="DIM_ORDER", con=sql_connection, if_exists="replace", index=False)

sql_connection.close()
