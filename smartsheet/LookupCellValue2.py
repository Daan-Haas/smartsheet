# --- built ins ---
# --- internals ---
from base import init_sheet, init_client

# --- externals ---

sheet_name = 'Addressbook'
workspace_name = 'API TEST'
result_column = 'Street Name'
lookup_value = 'Karel'
lookup_column = 'Name'

smart = init_client()
sheet_data, _ = init_sheet(workspace_name, sheet_name, smart)

result_column_id = sheet_data.get_column_by_title(result_column).id_
lookup_column_id = sheet_data.get_column_by_title(lookup_column).id_

lookup_row = None
for row in sheet_data.rows:
    for cell in row.cells:
        if cell.column_id == lookup_column_id and cell.value == lookup_value:
            lookup_row = row
            break
    if lookup_row:
        break

if lookup_row:
    for cell in lookup_row.cells:
        if cell.column_id == result_column_id:
            result_value = cell.value
            print(f"The value in the '{result_column}' column for the row with '{lookup_column}' equal to '{lookup_value}' is: '{result_value}'")
else:
    print(f"No row found where '{lookup_column}' is equal to '{lookup_value}'.")