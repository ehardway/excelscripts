# Import the module to read Execl Spreadsheets
import xlrd

# Import the module to write Execl Spreadsheets
from openpyxl import Workbook

# Import module to list files in directory
import glob

# Creating list of sales number spread sheets
excel_files = glob.glob("sales_numbers_*.xlsx")

# Declaring a list for sales data
sales_data = []

# Opening and iterating through workbooks
for excel_file in excel_files:
    workbook = xlrd.open_workbook(excel_file)
    # Iterating through the sheets
    for sheet_name in workbook.sheet_names():
        worksheet = workbook.sheet_by_name(sheet_name)
        # Iterating through the rows
        for row_number in range(1, 13):
            # building a data set from the rows
            row = worksheet.row(row_number)
            sales_data.append(
                {'product': row[0].value, 'month': row[1].value, 'year': row[2].value, 'sales': row[3].value})

# Declaring a dictionary for the yearly sales
yearly_sales = {}

# Iterating over the sales_data
for sale_data in sales_data:
    # Checking if there is an value in the dictionary for a product
    if sale_data['product'] in yearly_sales:
        # if there is value add the next month to it
        yearly_sales[sale_data['product']] = sale_data['sales'] + yearly_sales[sale_data['product']]
    else:
        # if there is not a value create the initial value
        yearly_sales[sale_data['product']] = sale_data['sales']

# calculate yearly average
yearly_average = {}
for product_name, yearly_sale in yearly_sales.items():
    yearly_average[product_name] = yearly_sale/12

# Create the workbook
yearly_book = Workbook()

# Create the Sheet
sheet1 = yearly_book.active
sheet1.title = "Yearly_Sales"

# Create the Header
header = [
    {'row': 1, 'col': 1, 'value': 'Product'},
    {'row': 1, 'col': 2, 'value': 'Yearly Sales'},
    {'row': 1, 'col': 3, 'value': 'Monthly Average'}
]
for rows in header:
    sheet1.cell(row=rows['row'], column=rows['col'], value=rows['value'])

# Create the data set in excel
row_start = 2
for product_name, yearly_amount in yearly_sales.items():
    sheet1["a" + str(row_start)] = product_name

    sheet1["b" + str(row_start)] = yearly_amount
    sheet1["b" + str(row_start)].number_format = '$#,##0.00'

    sheet1["c" + str(row_start)] = yearly_average[product_name]
    sheet1["c" + str(row_start)].number_format = '$#,##0.00'

    row_start += 1

# Save the spreadsheet
yearly_book.save('Yearly_Sales.xlsx')
