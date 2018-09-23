# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os.path
import json
from myconfig import *

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"

_dir = os.path.dirname(os.path.abspath(__file__))

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
column_map = {}

# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)


# TODO: Replace the body of this function with your code
# This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
#
# Return a new Row with updated cell values, else None to leave unchanged
def evaluate_row_and_build_updates(source_row):
    # Find the cell and value we want to evaluate
    status_cell = get_cell_by_column_name(source_row, "Status")
    print("==================================")
    print(status_cell)
    status_value = status_cell.display_value
    if (status_value == "Complete"):
        remaining_cell = get_cell_by_column_name(source_row, "Remaining")
        if (remaining_cell.display_value != "0"):                           # Skip if already 0
            print("Need to update row #" + str(source_row.row_number))

            # Build new cell value
            newCell = ss.models.Cell()
            newCell.column_id = column_map["Remaining"]
            newCell.value = 0

            # Build the row to update
            newRow = ss.models.Row()
            newRow.id = source_row.id
            newRow.cells.append(newCell)

            return newRow

    return None



print("Starting ...")

# Initialize client
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)
#print(vars(ss))
# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Import the sheet
#result = ss.Sheets.import_xlsx_sheet(_dir + '/Sample Sheet.xlsx', header_row_index=0)
#print(result.data.id)
# Load entire sheet
dashboardid = "8324841494341508"
#chartid = "1090278086993796"
sight = ss.Sights.get_sight(dashboardid)
print(sight)
print("\n\n")
#print(chart)

#print("Dashboard:" + sheet)

#print ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)
# Build column map for later reference - translates column names to column id
for widget in sight.widgets:
    #column_map[column.title] = column.id
    #print("Column: " + str(widget.id) + " " + widget.title )
    print(type(widget))
    #w_obj = json.loads(widget)
    print(widget)
    print("\n\n")

# Accumulate rows needing update here
rowsToUpdate = []

#for row in sheet.rows:
#    print("row" + str(row))
    #rowToUpdate = evaluate_row_and_build_updates(row)
    #if (rowToUpdate != None):
    #    rowsToUpdate.append(rowToUpdate)

# Finally, write updated cells back to Smartsheet
#if rowsToUpdate:
#    print("Writing " + str(len(rowsToUpdate)) + " rows back to sheet id " + str(sheet.id))
#    result = ss.Sheets.update_rows(result.data.id, rowsToUpdate)
#else:
#    print("No updates required")
        
print ("Done")
