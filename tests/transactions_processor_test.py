#!/usr/bin/python
import os
import sys
#get the transactions_processor directory to add to syspath, 
#took little hit&trial to come up with this piece of code
transactions_processor_path = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "../"))
sys.path.append(transactions_processor_path)

from transactions_processor import *

assert(parse_year_month("Fiscal Year 2000 - 2001","July",2001) == 2001)
assert(parse_year_month("Fiscal Year 2000 - 2001","July",2000) == 2000)
assert(parse_year_month("Fiscal Year 2000 - 2001","September",2000) == 2000)
assert(parse_year_month("Fiscal Year 2000 - 2001","January",2000) == 2001)

xlfilename = "testdata.xls"
sheet_name = "ampdata"
assert(validate_sheet_name(sheet_name, xlfilename) == True)

workbook = xlrd.open_workbook(xlfilename)
worksheet = workbook.sheet_by_name(sheet_name)
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1

(desired_cells, header_year_row, header_month_row) = get_data_cells_and_header(worksheet, "Commitments", "Total")
assert(len(desired_cells) == 22)
assert(len(header_year_row) == len(desired_cells)+2)

(desired_cells, header_year_row, header_month_row) = get_data_cells_and_header(worksheet, "Commitments", "Technical Assistance")
assert(len(desired_cells) == 18)

