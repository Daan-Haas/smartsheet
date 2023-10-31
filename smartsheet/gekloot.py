# --- built ins ---
# --- internals ---
from base import *

# --- externals ---

workspace_name = 'API TEST'
sheet_name = 'Machines'
lookup_value = 'A'
write_value = 'b'

write_column = 'Parts'
lookup_column = 'Machine'
count_column = 'Machine count'

client = init_client()
sheet_data, sheet_id = init_sheet(workspace_name, sheet_name, client)

write_column = sheet_data.get_column_by_title(write_column)
write_id = write_column.id_
lookup_column_id = sheet_data.get_column_by_title(lookup_column).id_
# count_column_id = sheet_data.get_column_by_title(count_column).id_

updated_row = None
for row in sheet_data.rows:
    for cell in row.cells:
        if cell.column_id == lookup_column_id and cell.value == lookup_value:
            updated_row = row
            break
    if updated_row:
        break

for cell in updated_row.cells:
    if cell.column_id == write_id:
        cell.value = None
        cell.object_value = write_column.options[0]

print(updated_row)
updated_row.created_at = None
updated_row.modified_at = None
updated_row.row_number = None
updated_row.below = True
write_column.id_ = None
write_column.version = None
write_column.type = "MULTI_PICKLIST"

write_column.level = '3'
write_column.include = "objectValue"

client.Sheets.update_rows(sheet_id, updated_row)
# client.Sheets.update_column(sheet_id, write_id, write_column)
