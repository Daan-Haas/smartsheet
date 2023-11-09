# --- built ins ---
# --- internals ---
from base import *
# --- externals ---

# Set up logging
logging.basicConfig(filename='errorlog.log', filemode='w', level=logging.WARNING)

# Read in key from file or user input
KEY = check_key()


""" These are the names of the workspace, sheets and columns that you want to perform actions on"""

# Work space. All your sheets need to be in the same workspace
workspace = 'Workspace name'

# Source sheets & columns for dropdown menus
dropdown_sheet = 'Dropdown Sheet Name'  # This is the sheet in which you would like to populate your dropdown
dropdown_col = 'Dropdown column name'  # This is the column that needs to become your dropdown

origin_sheet = 'Origin Sheet Name'  # This is the sheet containing the dropdown options
origin_col = 'Origin Column Name'  # This is the column containing the dropdown options
col_type = 'MULTI_PICKLIST'  # The type of picklist you want 'MULTI_PICKLIST' If you want multiple selectable
                             # 'PICKLIST' If you want only one option


"""These are the names of the co-dependent sheets and columns"""
frst_sheet = 'First Sheet Name'  # Name of the sheet containing one of the co-dependent columns
scnd_sheet = 'Second Sheet Name'  # Name of the sheet containing the other co-dependent column
frst_key_col = 'First Search Column'  # Name of the search column in the first sheet
frst_val_col = 'First Co-dependent Column'  # Name of the first co-dependent column
scnd_key_col = 'Second Search Column'  # Name of the search column in the second sheet
scnd_val_col = 'Second Co-dependent Column'  # Name of the second co-dependent column

if __name__ == '__main__':

    # Make and fill dropdown menus
    make_dropdown(KEY, workspace, origin_sheet, origin_col, dropdown_sheet, dropdown_col, col_type)

    # Update linked columns
    compare_dicts(KEY, workspace, frst_sheet, frst_key_col, frst_val_col, scnd_sheet, scnd_key_col, scnd_val_col)
