## AMP Transactions Data Processor

### Usage
usage transactions_processor.py [options]

Options:

  -i INPUT_FILENAME, --input=INPUT_FILENAME
  
  -s SHEET_NAME, --sheetname=SHEET_NAME
  
  -o OUTPUT_CSV_FILENAME, --output=OUTPUT_CSV_FILENAME
  
  -t TRANSACTION_TYPE_VALUE, --transaction_type=TRANSACTION_TYPE_VALUE
  
  -g GRANT_TYPE_VALUE, --grant_type=GRANT_TYPE_VALUE


### Description:

Reads AMP xls file and produces the transactions data for each project along with year and month segregation. The AMP xls file contains tabular data. Please see tests/testdata.xls for the format of xls file.


### Examples

`./transactions_processor.py -i inputamp.xls -s sheetname -t Commitment -g Total -o outputcommitments.csv`

Extracts "Total" of "Committments" from the input xls file, segregated by year/month. See tests/commitments_total.csv.

`./transactions_processor.py -i inputamp.xls -s sheetname -t Disbursement -g "Technical Assistance" -o outputdisbursements.csv`

Extracts "Technical Assistance" of "Disbursements" from the input xls file, segregated by year/month
