# --- built ins ---
# --- internals ---
from base import *

# --- externals ---

workspace = 'API TEST'
first_sheet = 'Machines'
second_sheet = 'Parts'
first_key_column = 'Machine'
first_values_column = 'Parts'
second_key_column = 'Part'
second_values_column = 'Machines'

if __name__ == '__main__':
    compare_dicts(workspace, first_sheet, first_key_column, first_values_column,
                  second_sheet, second_key_column, second_values_column)
