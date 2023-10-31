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

    if sheet_name not in [s.name for s in client.Workspaces.get_workspace(workspace_id).sheets.to_list()]:
        raise Exception(f'{sheet_name} not found in {workspace_name}, please check spelling and try again')

    sheets_data = client.Workspaces.get_workspace(workspace_id).to_dict()['sheets']
    sheet_id = next((w['id'] for w in sheets_data if w['name'] == sheet_name), None)

    return client.Sheets.get_sheet(sheet_id), sheet_id


def lookup_cellvalue(workspace_name, sheet_name, lookup_column, lookup_value, return_column, client):

    sheet_data, _ = init_sheet(workspace_name, sheet_name, client)

    return_column_id = sheet_data.get_column_by_title(return_column).id_
    lookup_column_id = sheet_data.get_column_by_title(lookup_column).id_

    lookup_row = None
    for row in sheet_data.rows:
        for cell in row.cells:
            if cell.column_id == lookup_column_id and cell.value == lookup_value:
                lookup_row = row
                break
        if lookup_row:
            break

    if lookup_row:
        for cell in lookup_row.cells:
            if cell.column_id == return_column_id:
                return_value = cell.value
                print(
                    f"The value in the '{return_column}' column for the row with '{lookup_column}' equal to "
                    f"'{lookup_value}' is: '{return_value}'")
    else:
        print(f"No row found where '{lookup_column}' is equal to '{lookup_value}'.")

    return lookup_column, lookup_value


def write_cellvalue(workspace_name, sheet_name, lookup_column, lookup_value, write_column, write_value, client):

    sheet_data, sheet_id = init_sheet(workspace_name, sheet_name, client)

    write_column_id = sheet_data.get_column_by_title(write_column).id_
    lookup_column_id = sheet_data.get_column_by_title(lookup_column).id_

    updated_row = None
    for row in sheet_data.rows:
        for cell in row.cells:
            if cell.column_id == lookup_column_id and cell.value == lookup_value:
                updated_row = row
                break
        if updated_row:
            break

    print(updated_row)
    for cell in updated_row.cells:
        if cell.column_id == write_column_id:
            cell.value = write_value
            break

    updated_row.created_at = None
    updated_row.modified_at = None
    updated_row.row_number = None
    updated_row.below = True

    print(updated_row)

    client.Sheets.update_rows(sheet_id, [updated_row])


def make_dropdown(origin_sheet_name, origin_column_name, target_sheet_name, target_column_name, client, workspace):
    # 1 Exception handling

    if workspace not in [ws.name for ws in client.Workspaces.list_workspaces().data]:
        raise Exception(f"{workspace} not found as workspace, please check spelling and try again")

    origin_sheet, _ = init_sheet(workspace, origin_sheet_name, client)
    target_sheet, target_sheet_id = init_sheet(workspace, target_sheet_name, client)
    origin_column_id = origin_sheet.get_column_by_title(origin_column_name).id_

    # 2 read the values from the sheet and fill list

    list_values = []
    for row in origin_sheet.rows:
        for cell in row.cells:
            if cell.column_id == origin_column_id:
                list_values.append(cell.value)

    # 3 find the target column

    target_column = next(c for c in target_sheet.columns if c.title == target_column_name)
    target_column_id = target_column.id

    # 4 update the column with the dropdown list, and remove attributes to write to ss

    target_column.options = list_values
    target_column.id = None
    target_column.version = None

    client.Sheets.update_column(target_sheet_id, target_column_id, target_column)


def write_in_column(workspace_name, sheet_name, lookup_column, lookup_value, write_column, write_value, client):
    if workspace_name not in [ws.name for ws in client.Workspaces.list_workspaces().data]:
        raise Exception(f"{workspace_name} not found as workspace, please check spelling and try again")

    origin_sheet, _ = init_sheet(workspace_name, sheet_name, client)
    target_sheet, target_sheet_id = init_sheet(workspace_name, sheet_name, client)
    origin_column_id = origin_sheet.get_column_by_title(lookup_column).id_

    print(origin_sheet.get_column_by_title(lookup_column))


def multi_dropdown(origin_sheet_name, origin_column_name, target_sheet_name, target_column_name, client, workspace):
    # 1 Exception handling
    pass
