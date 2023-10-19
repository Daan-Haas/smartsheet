# --- built ins ---
# --- Internals ---
# --- Externals ---
import smartsheet


def init_client():
    # Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
    smart = smartsheet.Smartsheet('47R8WwWUjjrnNwN6zGisOrHUh4Ed9b7CTKqya')
    # Make sure we don't miss any error
    smart.errors_as_exceptions(False)
    return smart

def init_sheet(workspace_name, sheet_name, client):
    workspaces_data = client.Workspaces.list_workspaces().to_dict()['data']
    workspace_id = next((w['id'] for w in workspaces_data if w['name'] == workspace_name), None)

    sheets_data = client.Workspaces.get_workspace(workspace_id).to_dict()['sheets']
    sheet_id = next((w['id'] for w in sheets_data if w['name'] == sheet_name), None)

    return client.Sheets.get_sheet(sheet_id), sheet_id

def makedropdown(origin_sheet_name, origin_column_name, target_sheet_name, target_column_name, client, workspace):

    # Exception handling
    if workspace not in [ws.name for ws in client.Workspaces.list_workspaces().data]:
        raise Exception(f"{workspace} not found as workspace, please check spelling and try again")

    origin_sheet, _ = init_sheet(workspace, origin_sheet_name, client)
    target_sheet, target_sheet_id = init_sheet(workspace, target_sheet_name, client)
    origin_column_id = origin_sheet.get_column_by_title(origin_column_name).id_

    # read the genders from the sheet and fill list with options
    options = []
    for row in origin_sheet.rows:
        for cell in row.cells:
            if cell.column_id == origin_column_id:
                options.append(cell.value)

    # find the column that needs dropdown menus
    target_column = next(c for c in target_sheet.columns if c.title == target_column_name)
    target_column_id = target_column.id

    # update the column with the dropdown list, and remove attributes to write to ss
    target_column.options = options
    target_column.id = None
    target_column.version = None

    client.Sheets.update_column(target_sheet_id, target_column_id, target_column)
