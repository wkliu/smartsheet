# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os.path
from myconfig import *
import re
# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"

_dir = os.path.dirname(os.path.abspath(__file__))

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here

column_r_map = {}
datatype = ("RRR", "Hot Issues", "Others")

# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)

print("Starting ...")

# Initialize client
proxies = {'http': 'http://proxy.esl.cisco.com:80/', 'https':'http://proxy.esl.cisco.com:80/'}
ss = smartsheet.Smartsheet(access_token=access_token, proxies=proxies)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Import the sheet
#result = ss.Sheets.import_xlsx_sheet(_dir + '/Sample Sheet.xlsx', header_row_index=0)

# Load destination sheet

destination_sheetId = "1914561801545604"
destination_sheet = ss.Sheets.get_sheet(destination_sheetId)

#print ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

# Build column map for later reference - translates column names to column id
columnToAdd = []
for column in destination_sheet.columns:
    column_r_map[column.title] = column.id

#print('Column R Map: ' + str(column_r_map))    
# Load source sheets using Loop
rowsToAddRRR = []
rowsToAddHotIssues = []
rowsToAddOthers = []
source_sheetIds = ["6986082806982532", "2587280381634436", "8904432097224580", "3099560458381188",
"831680287139716", "4676809888425860", "400534290098052", "2395220181575556", "7602507250722692"]
#source_sheetIds = ["7319980007024516"]
row_num = 0 
for source_sheetId in source_sheetIds:
    source_sheet = ss.Sheets.get_sheet(source_sheetId)
    column_map = {}
    for column in source_sheet.columns:
        column_map[column.id] = column.title

    # Accumulate rows needing update here

    for row in source_sheet.rows:
        dtype = ''
        #print(row.row_number)
        if row.cells[0].display_value == None:
            continue
        rowObject = ss.models.Row()
        rowObject.to_bottom = True    
        #row_num = row_num + 1
        #print(get_cell_by_column_name(row, "Deal ID"))
        for cell in row.cells:
            if (cell.display_value != None):
                cell.column_id = column_r_map[column_map[cell.column_id]]
                #print("Cell value:" + str(column_map[cell.column_id]) + ":" + repr(cell.display_value))
                if cell.value in datatype:
                    dtype = cell.value
                    print('type:' + cell.value)
                if cell.column_id == column_r_map['HTTP Link']: 
                    if dtype == "RRR":
                       row_num = row_num + 1
                    formula = ''
                    cell.value = ''   
                    formula = re.findall(r'(.*Type)\d*(.*\[Engagement ID\])\d*', str(cell.formula))
                    cell.formula = str(formula[0][0]) + str(row_num) + str(formula[0][1]) + str(row_num)
                rowObject.cells.append(cell) 
        #print("Row: " + str(rowObject))        
        if dtype == 'RRR':      
            rowsToAddRRR.append(rowObject)
        elif dtype == 'Hot Issues':
            rowsToAddHotIssues.append(rowObject)
        elif dtype == 'Others':
            rowsToAddOthers.append(rowObject)    
        #rowToUpdate = evaluate_row_and_build_updates(row)
        #if (rowToUpdate != None):
        #    rowsToUpdate.append(rowToUpdate)

    # Finally, write updated cells back to Smartsheet
    #if rowsToUpdate:
    #    print("Writing " + str(len(rowsToUpdate)) + " rows back to sheet id " + str(sheet.id))
    #    result = ss.Sheets.update_rows(result.data.id, rowsToUpdate)
    #else:
    #   print("No updates required")
#print("Top5 Rows:" + str(rowsToAddTop5))
#print("Win Case Rows:" + str(rowsToAddWin))
#print("Loss Case Rows:" + str(rowsToAddLoss))
#print("column_id:" + str(column_r_map['Architectural Plays']))
#print('HTTP: ' + str(column_r_map['HTTP Link']))
newline = ss.models.Row()
#newline.to_bottom = True
#print('Length of Top5: ' + str(len(rowsToAddTop5)))
#print('Length of Win Case: ' + str(len(rowsToAddWin)))
#print('Length of Loss Case: ' + str(len(rowsToAddLoss)))
print(destination_sheet.add_rows(rowsToAddRRR))
print(destination_sheet.add_rows(newline))
destination_sheet.add_rows(rowsToAddHotIssues)
destination_sheet.add_rows(newline)  
destination_sheet.add_rows(rowsToAddOthers)        
print ("Done")
