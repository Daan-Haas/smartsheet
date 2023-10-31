# --- built ins ---
# --- internals ---
from base import *

# --- externals ---

target_sheet = 'Parts'
origin_sheet = 'Machines'
lookup_column = 'Machines'
target_column = 'Part'
workspace_name = 'API TEST'

# Initialize the smartsheet client and sheets
smart = init_client()

column, value = lookup_cellvalue(workspace_name, target_sheet, target_column, 'a', lookup_column, smart)

