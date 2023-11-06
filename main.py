# --- built ins ---
import logging
# --- internals ---
from base import *

# --- externals ---


# Set up logging
logging.basicConfig(filename='errorlog.log', filemode='w', level=logging.WARNING)

# Work space
workspace = 'Standardisation Prime API TEST'

# Source sheets & columns for dropdown menus
enigines = 'Engine Models'
enigines_col = 'Engine Models'
enigines_target_col = 'Engine Models'

job_codes = 'Job Codes'
job_codes_col = 'Job Code'
Naxt_JC = 'Naxt Job Code'
Actual_JC = 'Actual Job Code'

component_codes = 'Component Codes'
component_codes_col = 'Component Code'
Naxt_CC = 'Naxt Component Code'
Actual_CC = 'Actual Component Code'

# Sheets to compare
SP = 'Sales Packages'
SJ = 'Standard Jobs/Building Blocks'

# Key and value columns
SP_key = 'SP Number (Naxt)'
SJ_key = 'Naxt Reference'
SP_val = 'Standard Jobs'
SJ_val = 'Sales Packages'

if __name__ == '__main__':

    # Make and fill dropdown menus
    # make_dropdown(workspace, enigines, enigines_col, SJ, enigines_target_col, 'MULTI_PICKLIST')
    # make_dropdown(workspace, job_codes, job_codes_col, SJ, Naxt_JC, 'PICKLIST')
    # make_dropdown(workspace, job_codes, job_codes_col, SJ, Actual_JC, 'PICKLIST')
    # make_dropdown(workspace, component_codes, component_codes_col, SJ, Naxt_CC, 'PICKLIST')
    # make_dropdown(workspace, component_codes, component_codes_col, SJ, Actual_CC, 'PICKLIST')
    # make_dropdown(workspace, SJ, SJ_key, SP, SP_val, 'MULTI_PICKLIST')
    # Update linked columns
    compare_dicts(workspace, SJ, SJ_key, SJ_val, SP, SP_key, SP_val)
