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
datatype = ("RRR", "CAP", "Hot Issues", "Others")

# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)

print("Starting ...")

# Initialize client
proxies = {'http': 'http://proxy.esl.cisco.com:80/', 'https':'http://proxy.esl.cisco.com:80/'}
#ss = smartsheet.Smartsheet(access_token=access_token, proxies=proxies)
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)
formula_string = "=IF(OR(Type29 = \"RRR\", Type29 = \"CAP\"), \"https://clpsvs.cloudapps.cisco.com/services/clip/main/transaction/\" + [Engagement ID]29}]"
# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Import the sheet
#result = ss.Sheets.import_xlsx_sheet(_dir + '/Sample Sheet.xlsx', header_row_index=0)

# Load destination sheet

destination_sheetId = destination_sheetIds["RRR"] 
destination_sheet = ss.Sheets.get_sheet(destination_sheetId)

#print ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

# Build column map for later reference - translates column names to column id
columnToAdd = []
for column in destination_sheet.columns:
    column_r_map[column.title] = column.id

#print('Column R Map: ' + str(column_r_map))    
# Load source sheets using Loop
rowsToAddRRR = []
rowsToAddCAP = []
rowsToAddHotIssues = []
rowsToAddOthers = []
source_sheetIds = [allentseng["RRR"], andrewyang["RRR"], angelalin["RRR"],
                   barryhuang["RRR"], davidtai["RRR"], jerrylin["RRR"],
                   jimcheng["RRR"], karlhsieh["RRR"], rickywang["RRR"],
                   stanhuang["RRR"], tonyhsieh["RRR"], vanhsieh["RRR"],
                   vinceliu["RRR"], vincenthsu["RRR"], willyhuang["RRR"]]
RRR_row_num = 0
CAP_row_num = 0
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
                    print(dtype)
                    if dtype == 'RRR' or dtype == 'CAP':
                       cell.value = ''   
                    #formula = re.findall(r'(.*Type)\d*(.*Type)\d*(.*\[Engagement ID\])\d*', str(cell.formula))
                    #cell.formula = str(formula[0][0]) + str(RRR_row_num) + str(formula[0][1]) + str(RRR_row_num) + str(formula[0][2]) + str(RRR_row_num)
                rowObject.cells.append(cell) 
        print("Row: " + str(rowObject))        
        if dtype == 'RRR':      
            rowsToAddRRR.append(rowObject)
        elif dtype == 'CAP':
            rowsToAddCAP.append(rowObject)
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
newline = ss.models.Row()
#newline.to_bottom = True
updated_rows =[]
deleted_rows =[]
final_rows = []
for row in rowsToAddRRR:
    final_rows.append(row)
#final_rows.append(newline)
for row in rowsToAddCAP:
    final_rows.append(row)
#final_rows.append(newline)
for row in rowsToAddHotIssues:
    final_rows.append(row)
#final_rows.append(newline)
for row in rowsToAddOthers:
    final_rows.append(row)

print(len(final_rows))
num = 0
while len(final_rows)> num :
    print(num)
    if hasattr(final_rows[num], 'cells'):
        for cell in final_rows[num].cells:
            if (cell.display_value != None):
                #cell.column_id = column_r_map[column_map[cell.column_id]]
                #print("Cell value:" + str(column_map[cell.column_id]) + ":" + repr(cell.display_value))
                if cell.value in datatype:
                    dtype = cell.value
                    print('type:' + cell.value)
                if cell.column_id == column_r_map['HTTP Link']: 
                    print(dtype)
                    if dtype == 'RRR' or dtype == 'CAP':
                       cell.value = ''   
                       formula = re.findall(r'(.*Type)\d*(.*Type)\d*(.*\[Engagement ID\])\d*', str(cell.formula))
                       cell.formula = str(formula[0][0]) + str(num+1) + str(formula[0][1]) + str(num+1) + str(formula[0][2]) + str(num+1)
               
    
    print(final_rows[num])
    num = num +1

destination_sheet.add_rows(final_rows)

