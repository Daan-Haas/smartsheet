# --- built ins ---
# --- internals ---
from base import init_sheet, init_client, lookup_cellvalue

# --- externals ---

sheet_name = 'Addressbook'
workspace_name = 'API TEST'
result_column = 'Street Name'
lookup_value = 'Karel'
lookup_column = 'Name'

smart = init_client()

lookup_cellvalue(workspace_name, sheet_name, lookup_column, lookup_value, result_column, smart)