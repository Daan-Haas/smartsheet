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
workspaces_data = smart.Workspaces.list_workspaces().to_dict()['data']
workspace_id = next((w['id'] for w in workspaces_data if w['name'] == workspace_name), None)

print(smart.Workspaces.get_workspace(workspace_id))



