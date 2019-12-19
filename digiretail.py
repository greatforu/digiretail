import xlrd
import mysql.connector

book = xlrd.open_workbook("beginner_assignment01.xlsx")
sheet1 = book.sheet_by_name("product_listing")
sheet2 = book.sheet_by_name("group_listing")

database = mysql.connector.connect (host="localhost", user = "root", passwd = "", db = "mysqlPython")
# Get the cursor, which is used to traverse the database, line by line
cursor = database.cursor()

# Create the INSERT INTO sql query
query1 = """INSERT INTO digiretail (product_name, model_name, product_serial_number, group_associated, product_mrp) VALUES (%s, %s, %s, %s, %s)"""
query2 = """INSERT INTO digiretailgroup (group_name, group_desc, isactive) VALUES (%s, %s, %s)"""

for r in range(1, sheet1.nrows):
    product_name = sheet1.cell(r,0).value
    model_name	= sheet1.cell(r,1).value
    product_serial_number = sheet1.cell(r,2).value
    group_associated = sheet1.cell(r,3).value
    product_mrp	= sheet1.cell(r,4).value

    # Assign values from each row
    values = (product_name, model_name, product_serial_number, group_associated, product_mrp)

    # Execute sql Query
    cursor.execute(query1, values)

for r in range(1, sheet2.nrows):
    group_name = sheet1.cell(r,0).value
    group_desc	= sheet1.cell(r,1).value
    isactive = sheet1.cell(r,2).value

    values = (group_name, group_desc, isactive)
    cursor.execute(query2, values)

# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()

# print( results
print( "")
print( "All Done!")
print( "")
columns = str(sheet1.ncols)
rows = str(sheet1.nrows)
# print( "I just imported " %2B  columns %2B " columns and " %2B rows %2B " rows to MySQL!")