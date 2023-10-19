# --- built ins ---
# --- internals ---
from base import init_sheet, init_client, makedropdown

# --- externals ---

gender_sheet = 'Genders'
addressbook_sheet = 'Addressbook'
workspace_name = 'API TEST'
lookup_column = 'Genders'

# Initialize the smartsheet client and sheets
smart = init_client()

makedropdown(gender_sheet, lookup_column, addressbook_sheet, 'Gender', smart, workspace_name)
