# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os.path
from myconfig import *

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"

_dir = os.path.dirname(os.path.abspath(__file__))

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here

column_r_map = {}
quarter = ("FY19Q1", "FY19Q2", "FY19Q3", "FY19Q4")

# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)

print("Starting ...")

# Initialize client
proxies = {'http': 'http://proxy.esl.cisco.com:80/', 'https':'http://proxy.esl.cisco.com:80'}
#ss = smartsheet.Smartsheet(access_token=access_token, proxies=proxies)
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Import the sheet
#result = ss.Sheets.import_xlsx_sheet(_dir + '/Sample Sheet.xlsx', header_row_index=0)

# Load destination sheet

destination_sheetId = destination_sheetIds["KSO"] 
destination_sheet = ss.Sheets.get_sheet(destination_sheetId)

#print ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

# Build column map for later reference - translates column names to column id
columnToAdd = []
for column in destination_sheet.columns:
    column_r_map[column.title] = column.id

# Load source sheets using Loop
rowsToAddFY19Q1 = []
rowsToAddFY19Q2 = []
rowsToAddFY19Q3 = []
rowsToAddFY19Q4 = []
source_sheetIds = {"Vince Liu": vinceliu["KSO"],"Angela Lin": angelalin["KSO"],"Jim Cheng": jimcheng["KSO"], "Karl Hsieh":karlhsieh["KSO"], "Allen Tseng": allentseng["KSO"], "Andrew Yang":andrewyang["KSO"], "Barry Huang":barryhuang["KSO"], "David Tai": davidtai["KSO"], "Jerry Lin": jerrylin["KSO"], "Ricky Wang": rickywang["KSO"], "Stan Huang": stanhuang["KSO"], "Tony Hsieh": tonyhsieh["KSO"], "Van Hsieh": vanhsieh["KSO"],"Vincent Hsu": vincenthsu["KSO"], "Willy Huang": willyhuang["KSO"]}

for owner in source_sheetIds:
    source_sheetId = source_sheetIds[owner]
    source_sheet = ss.Sheets.get_sheet(source_sheetId)
    column_map = {}
    for column in source_sheet.columns:
        column_map[column.id] = column.title

    # Accumulate rows needing update here

    for row in source_sheet.rows:
        #print(str(row))
        if row.cells[0].display_value in quarter:
            qtype = row.cells[0].display_value 
        if row.cells[1].display_value == None:
            continue
        rowObject = ss.models.Row()
        rowObject.to_bottom = True    
        for cell in row.cells:
            if (cell.display_value != None):
                #print("Cell value:" + str(column_map[cell.column_id]) + ":" + repr(cell.display_value))
                cell.column_id = column_r_map[column_map[cell.column_id]]
                rowObject.cells.append(cell) 
        rowObject.cells.append({"value": owner, "columnId": column_r_map['Owner'], "displayValue": owner})
        #print("Row: " + str(rowObject))        
        if qtype == 'FY19Q1':      
            rowsToAddFY19Q1.append(rowObject)
        elif qtype == 'FY19Q2':
            rowsToAddFY19Q2.append(rowObject)
        elif qtype == 'FY19Q3':
            rowsToAddFY19Q3.append(rowObject)    
        elif qtype == 'FY19Q4':
            rowsToAddFY19Q4.append(rowObject)
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
newline = ss.models.Row()
newline.to_bottom = True
destination_sheet.add_rows(rowsToAddFY19Q1)
destination_sheet.add_rows(newline)
destination_sheet.add_rows(rowsToAddFY19Q2)
destination_sheet.add_rows(newline) 
destination_sheet.add_rows(rowsToAddFY19Q3) 
destination_sheet.add_rows(newline) 
destination_sheet.add_rows(rowsToAddFY19Q4)        
print ("Done")
