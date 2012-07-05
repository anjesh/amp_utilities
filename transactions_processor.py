#!/usr/bin/python
"""
AMP Transactions Data Processor

author: Anjesh Tuladhar <anjesh@yipl.com.np>

Reads AMP xls file 
and produces the disburement data for each project along with year and month segregation

Following table is read

---------------------------------------------------------------------------------------------
|							Fiscal Year 2000 - 2001											|
|    								January													|  
|		Actual Commitments				|		Actual Disbursements						|
|	Technical Assistance	| 	Total	|	Technical Assistance | Grant Aid	|	Total 	|
---------------------------------------------------------------------------------------------
|			1000			|	1000	|			2000		 |	500			|	2200	|
---------------------------------------------------------------------------------------------

Extracts the column number of data we are interested in, 



"""
import xlrd
import csv
from optparse import OptionParser
import os

#followings constants used in getDataCellsAndHeader
HEADER_YEAR_TEXT_ROW = 6
HEADER_MONTH_ROW = 7
HEADER_TRANSACTION_TYPE_ROW = 8
HEADER_GRANT_TYPE_ROW = 9

#used while reading data rows
START_DATA_ROW = 11



def parse_year_month(year_text, month, previous_year):
	"""
	parses the year_text and gets the appropriate year to the given month as per fiscal year of Nepal
	year_text = Fiscal Year 2000 - 2001
	previous_year is used to handle for the month of July as it may lie in either of two english years
	"""
	#first half of the nepal fiscal year
	MONTHS1 = ["August", "September", "October", "November", "December"]
	#second half of the nepal fiscal year, note missing July in both as it is handled separately
	MONTHS2 = ["January", "February", "March", "April", "May", "June"]
	month = month.strip()
	year_text = year_text.strip()
	if year_text.find("-") != -1:
		year_dash_pos = year_text.index("-")
		#AMP year header contains "Fiscal Year 2000 - 2001"
		#in that case, year1 = 2000 and year2 = 2001
		year1 = int(year_text[year_dash_pos-5:year_dash_pos])
		year2 = int(year_text[year_dash_pos+1:year_dash_pos+6])
	else:
		return False
	if month in MONTHS1:
		#if passed month present in MONTHS1 array, then year must be year1 of the year_text		
		year = year1
	elif month in MONTHS2:
		#similarly for year2
		year = year2
	else: #month is "July":	
		#july is special case as it may appear in both year1 and year2
		#the only logic (simple) employed is to check for the previous year, there might still be error here
		#if previous month/year = Jun 2011, then july must belong to 2011		
		#for fiscal year 2011-2012, july may belong to both 2011 or 2012
		#from above processing, if previous date=May 2012, then July belongs to 2012
		#@TODO: still need to get good documentation here, very confusing while trying to document :(
		if previous_year == year1:
			year = year1
		else:
			year = year2
	return year


def get_data_cells_and_header(worksheet, transaction_type_value, grant_type_value):
	"""
	read headers and prepare rows for year and month header and also get the cells numbers to read
	"""
	num_cells = worksheet.ncols - 1
	previous_year = year = month = transaction_type = grant_type  = year_text = ""
	#desired cells contains the cell cols which we are interested in depending upon the passed
	# params, transaction_type_value and grant_type_value	
	desired_cells = []
	curr_cell = 1
	#first two columns of year and month rows will contain donor and project title
	header_year_row = ["",""]
	header_month_row = ["",""]
	while curr_cell < num_cells:
		curr_cell += 1
		year_text = worksheet.cell_value(HEADER_YEAR_TEXT_ROW, curr_cell) if worksheet.cell_value(HEADER_YEAR_TEXT_ROW, curr_cell) else year_text
		month = worksheet.cell_value(HEADER_MONTH_ROW, curr_cell) if worksheet.cell_value(HEADER_MONTH_ROW, curr_cell) else month
		previous_year = year = parse_year_month(year_text, month, previous_year)		
		transaction_type = worksheet.cell_value(HEADER_TRANSACTION_TYPE_ROW, curr_cell) if worksheet.cell_value(HEADER_TRANSACTION_TYPE_ROW, curr_cell) else transaction_type
		grant_type = worksheet.cell_value(HEADER_GRANT_TYPE_ROW, curr_cell) if worksheet.cell_value(HEADER_GRANT_TYPE_ROW, curr_cell) else grant_type
		if year and transaction_type.find(transaction_type_value) != -1 and grant_type.find(grant_type_value) != -1:
			desired_cells.append(curr_cell)
			header_year_row.append(year)
			header_month_row.append(month)
			#print curr_cell, year_text, year, month, transaction_type, grant_type
	return (desired_cells, header_year_row, header_month_row)


def main(worksheet, transaction_type_value, grant_type_value, csv_file):
	"""
	main function which reads the data rows and prepares the csv file with needed content
	"""	
	(desired_cells, header_year_row, header_month_row) = get_data_cells_and_header(worksheet, transaction_type_value, grant_type_value)
	datawriter = csv.writer(open(csv_file, "w"))
	datawriter.writerow(header_year_row)
	datawriter.writerow(header_month_row)

	#we know from the file, data row starts from row 11
	#while reading csv, first row = 0
	curr_row = START_DATA_ROW - 1
	while curr_row < num_rows:
		row = worksheet.row(curr_row)
		column1 = worksheet.cell_value(curr_row, 0)
		project_name = worksheet.cell_value(curr_row, 1)
		#add organization name and project name in the first 2 columns of each row
		data_row = [column1.encode("utf-8"), project_name.encode("utf-8")]
		for curr_cell in desired_cells:
			data_row.append(worksheet.cell_value(curr_row, curr_cell))
		datawriter.writerow(data_row)	
		curr_row += 1

def parse_command_line_params():
    usage = 'usage %prog [options]'
    version = 'Version: %prog 0.1'
    parser = OptionParser(usage=usage, version=version)
    parser.add_option("-i", "--input", dest="input_filename", default=False)
    parser.add_option("-s", "--sheetname", dest="sheet_name", default=False)
    parser.add_option("-o", "--output", dest="output_csv_filename", default="output.csv")
    parser.add_option("-t", "--transaction_type", dest="transaction_type_value", default="Commitment")
    parser.add_option("-g", "--grant_type", dest="grant_type_value", default="Total")
    (opts, args) = parser.parse_args()
    return parser, opts, args

def validate_input_filename(input_filename):	
	if not input_filename:
		print "--input filename must be mentioned"
		return False
	else:
		if not os.path.isfile(input_filename):
			print input_filename, " doesn't seem to exist. Please check file"
			return False
	return True

def validate_sheet_name(sheet_name, input_filename):
	tmpworksheet = xlrd.open_workbook(input_filename)
	worksheet_names = tmpworksheet.sheet_names()
	if sheet_name in worksheet_names:
		return True
	else:
		if not sheet_name:
			print "--sheetname must be mentioned"
		else:
			print sheet_name, " doesn't exist in the file"
			print "Available worksheets are ", 
			for worksheet_name in worksheet_names:
				print "\n - ",worksheet_name
		return False

if __name__ == "__main__":
	parser, opts, args = parse_command_line_params()
	if validate_input_filename(opts.input_filename) and validate_sheet_name(opts.sheet_name, opts.input_filename):		
		workbook = xlrd.open_workbook(opts.input_filename)
		worksheet = workbook.sheet_by_name(opts.sheet_name)
		num_rows = worksheet.nrows - 1
		num_cells = worksheet.ncols - 1
		main(worksheet, opts.transaction_type_value, opts.grant_type_value, opts.output_csv_filename)
