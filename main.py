# --- built ins ---
# --- internals ---
from base import *
from datetime import datetime
# --- externals ---

# Set up logging
logging.basicConfig(filename='errorlog.log', filemode='w', level=logging.WARNING)

# Check datetime to see if new Key required
fmt = "%Y-%m-%d %H:%M:%S"
last_run = ""
current_time = datetime.now()

try:  # Read in the last run time, then overwrite with current time
    with open("last run.txt", mode="r") as file:
        last_run = datetime.strptime(file.read(), fmt)
    with open('last run.txt', mode='w') as file:
        file.write(current_time.strftime(fmt))

except FileNotFoundError:  # If no last run file found
    with open('last run.txt', mode='w') as file:  # Write one
        file.write(current_time.strftime(fmt))
    KEY = input('Please Input API key')  # And ask for API key
    with open('API key.txt', 'w') as key_file:
        key_file.write(KEY)

# Check Time delta
if isinstance(last_run, datetime):
    td = current_time - last_run
else:
    td = 8

if td.days > 7:  # If too long ago, ask for new key
    KEY = input('Please Input API key')
    with open('API key.txt', 'w') as key_file:
        key_file.write(KEY)
else:
    with open('API key.txt') as key_file:
        KEY = key_file.readline()

# Work space
workspace = 'Standardisation Prime'

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
SP_key = 'Name of Standard'
SJ_key = 'Naxt Reference'
SP_val = 'Standard Jobs'
SJ_val = 'Sales Packages'

if __name__ == '__main__':

    # Make and fill dropdown menus
    make_dropdown(KEY, workspace, enigines, enigines_col, SJ, enigines_target_col, 'MULTI_PICKLIST')
    make_dropdown(KEY, workspace, job_codes, job_codes_col, SJ, Naxt_JC, 'PICKLIST')
    make_dropdown(KEY, workspace, job_codes, job_codes_col, SJ, Actual_JC, 'PICKLIST')
    make_dropdown(KEY, workspace, component_codes, component_codes_col, SJ, Naxt_CC, 'PICKLIST')
    make_dropdown(KEY, workspace, component_codes, component_codes_col, SJ, Actual_CC, 'PICKLIST')
    make_dropdown(KEY, workspace, SJ, SJ_key, SP, SP_val, 'MULTI_PICKLIST')

    # Update linked columns
    compare_dicts(KEY, workspace, SJ, SJ_key, SJ_val, SP, SP_key, SP_val)
