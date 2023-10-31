# --- built ins ---
# --- internals ---
from base import write_cellvalue, init_client

# --- externals ---

sheet_name = 'Parts'
workspace_name = 'API TEST'
write_value = 'test'
write_column = 'Machine count'
lookup_value = 'c'
lookup_column = 'Part'

smart = init_client()

write_cellvalue(workspace_name, sheet_name, lookup_column, lookup_value, write_column, write_value, smart)

