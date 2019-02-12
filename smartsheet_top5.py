# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os.path
from myconfig import *

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"

_dir = os.path.dirname(os.path.abspath(__file__))

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here

column_r_map = {}
datatype = ("Top5", "Win Case", "Loss Case")

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

# Load destination sheet

destination_sheetId_all = destination_sheetIds["TOP5_ALL"] 
destination_sheet = ss.Sheets.get_sheet(destination_sheetId_all)

#print ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

# Build column map for later reference - translates column names to column id
columnToAdd = []
for column in destination_sheet.columns:
    column_r_map[column.title] = column.id

#print('Column R Map: ' + str(column_r_map))    
# Load source sheets using Loop
rowsToAddWin = []
rowsToAddLoss = []
rowsToAddTop5 = []
source_sheetIds_SP = [davidtai["Top5"], andrewyang["Top5"], stanhuang["Top5"]]
source_sheetIds_RMT = [jimcheng["Top5"], karlhsieh["Top5"], vincenthsu["Top5"]]
source_sheetIds_FSI = [barryhuang["Top5"], angelalin["Top5"], vanhsieh["Top5"]]
source_sheetIds_PS = [rickywang["Top5"], willyhuang["Top5"]]
source_sheetIds_COM = [jerrylin["Top5"], allentseng["Top5"], vinceliu["Top5"]]
#source_sheetIds = ["7319980007024516"]
for source_sheetId in source_sheetIds:
    source_sheet = ss.Sheets.get_sheet(source_sheetId)
    column_map = {}
    for column in source_sheet.columns:
        column_map[column.id] = column.title

    # Accumulate rows needing update here

    for row in source_sheet.rows:
        #print(row.row_number)
        if row.cells[0].display_value in datatype:
            dtype = row.cells[0].display_value
            #print(dtype)
            continue
        elif row.cells[0].display_value == None:
            continue
        rowObject = ss.models.Row()
        rowObject.to_bottom = True    
        #print(get_cell_by_column_name(row, "Deal ID"))
        for cell in row.cells:
            if (cell.display_value != None):
                #print("Cell value:" + str(column_map[cell.column_id]) + ":" + repr(cell.display_value))
                cell.column_id = column_r_map[column_map[cell.column_id]]
                rowObject.cells.append(cell) 
        #print("Row: " + str(rowObject))        
        if dtype == 'Top5':      
            rowsToAddTop5.append(rowObject)
        elif dtype == 'Win Case':
            rowsToAddWin.append(rowObject)
        elif dtype == 'Loss Case':
            rowsToAddLoss.append(rowObject)    
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
top5row = ss.models.Row()
top5row.to_bottom = True
top5row.cells.append({'columnId': column_r_map['Architectural Plays'], 'displayValue': 'Top5', 'value': 'Top5' })
winrow = ss.models.Row()
winrow.to_bottom = True
winrow.cells.append({'columnId': column_r_map['Architectural Plays'], 'displayValue': 'Win Case', 'value': 'Win Case' })
lossrow = ss.models.Row()
lossrow.to_bottom = True
lossrow.cells.append({'columnId': column_r_map['Architectural Plays'], 'displayValue': 'Loss Case', 'value': 'Loss Case' })
destination_sheet.add_rows(top5row) 
#print('Length of Top5: ' + str(len(rowsToAddTop5)))
#print('Length of Win Case: ' + str(len(rowsToAddWin)))
#print('Length of Loss Case: ' + str(len(rowsToAddLoss)))
destination_sheet.add_rows(rowsToAddTop5)
destination_sheet.add_rows(newline)
destination_sheet.add_rows(winrow) 
destination_sheet.add_rows(rowsToAddWin)
destination_sheet.add_rows(newline) 
destination_sheet.add_rows(lossrow) 
destination_sheet.add_rows(rowsToAddLoss)        
print ("Done")
