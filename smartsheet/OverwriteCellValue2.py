# --- built ins ---
# --- internals ---
from base import init_sheet, init_client

# --- externals ---

sheet_name = 'Addressbook'
workspace_name = 'API TEST'
write_value = 'Hoogstraat'
write_column = 'Street Name'
lookup_value = 'Kees'
lookup_column = 'Name'

smart = init_client()
sheet_data, sheet_id = init_sheet(workspace_name, sheet_name, smart)

write_column_id = sheet_data.get_column_by_title(write_column).id_

lookup_column_id = sheet_data.get_column_by_title(lookup_column).id_


updated_row = None
for row in sheet_data.rows:
    for cell in row.cells:
        if cell.column_id == lookup_column_id and cell.value == lookup_value:
            updated_row = row
            break
    if updated_row:
        break

for cell in updated_row.cells:
    if cell.column_id == write_column_id:
        cell.value = write_value
        break


updated_row.created_at = None
updated_row.modified_at = None
updated_row.row_number = None
updated_row.below = True

smart.Sheets.update_rows(sheet_id, [updated_row])

