# --- built ins ---
import sys
# --- Internals ---
# --- Externals ---
from smartsheet import Smartsheet, sheets, workspaces
from smartsheet.models import Column, Sheet, Cell, MultiPicklistObjectValue, Row


def init_client():

    """Initialize smartsheet.smartsheet object, unlocked with API key"""

    # Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
    smart = Smartsheet('47R8WwWUjjrnNwN6zGisOrHUh4Ed9b7CTKqya')
    # Make sure we don't miss any error
    smart.errors_as_exceptions(False)
    return smart


def init_sheet(workspace_name, sheet_name, client):

    """Read in smartsheet to python

    Args:
        workspace_name (str): Name of workspace
        sheet_name (str): Sheet name
        client (Client): Smartsheet client

    Returns:
        Void
    """

    try:
        # Find workspace ID so sheets van be extracted
        workspaces_data = client.Workspaces.list_workspaces().to_dict()['data']
        workspace_id = next((w['id'] for w in workspaces_data if w['name'] == workspace_name), None)

    # Error handling in case of invalid key
    except KeyError:
        print('Invalid Key')
        sys.exit(1)

    # Error handling if workspace or sheet not found
    if not workspace_id:
        raise Exception(f"'{workspace_name}' not found in user's workspaces")
    if sheet_name not in [s.name for s in client.Workspaces.get_workspace(workspace_id).sheets.to_list()]:
        raise Exception(f"no sheet: '{sheet_name}' found in workspace: '{workspace_name}'"
                        ", please check spelling and try again")

    # Extract sheet data and sheet ID
    sheets_data = client.Workspaces.get_workspace(workspace_id).to_dict()['sheets']
    sheet_id = next((w['id'] for w in sheets_data if w['name'] == sheet_name), None)

    return client.Sheets.get_sheet(sheet_id)


def make_dropdown(workspace_name, origin_sheet_name, origin_column_name,
                  target_sheet_name, target_column_name, col_type):

    """Populates a dropdown list with variables from a sheet, and updates a column in a different sheet to this dropdown
    list

    Args:
        workspace_name (str): Workspace name
        origin_sheet_name (str): Name of sheet where you will retreive data
        origin_column_name (str): Name of column where you will retreive data
        target_sheet_name (str): Name of sheet where you will populate dropdown
        target_column_name (str): Name of column where to populate dropdown
        col_type (str): Type of column

    Returns:
        Void
    """

    # Initiate client
    client = init_client()

    # Create sheet objects and map clumn names to IDs
    origin_sheet = init_sheet(workspace_name, origin_sheet_name, client)
    target_sheet = init_sheet(workspace_name, target_sheet_name, client)

    # Error handling if incorrect column names
    if origin_column_name not in [c.title for c in origin_sheet.columns]:
        raise Exception(f"No column called: '{origin_column_name}' found in sheet: '{origin_sheet.name}'")
    if target_column_name not in [c.title for c in target_sheet.columns]:
        raise Exception(f"No column called: '{origin_column_name}' found in sheet: '{origin_sheet.name}'")

    if col_type not in ['MULTI_PICKLIST', 'PICKLIST']:
        raise Exception(f'Invalid column type entered: {col_type}')
    # Fill options list with entries of column
    origin_column_id = origin_sheet.get_column_by_title(origin_column_name).id
    options = []
    for row in origin_sheet.rows:
        cell = row.get_column(origin_column_id)
        options.append(cell.value)

    # Find target column for name and index
    target_column = target_sheet.get_column_by_title(target_column_name)

    # Make new column variable with multi picklist
    update_column = Column({
        'title': target_column.title,
        'type': col_type,
        'options': options,
    })

    # POST updated column to smartsheet
    client.Sheets.update_column(target_sheet.id, target_column.id, update_column)


def make_picklist_cell(column_id, values: list):

    """Generates a cell that can be PUT into a MULTI PICK dropdown column

    Args:
        column_id (int): ID of column which contains cell of interest
        values (lst): list of values to populate picklist cell

    Returns:
        new_cell (Cell): Smartsheet cell object
    """

    new_cell = Cell({
        'column_id': column_id,
        'Object_value': MultiPicklistObjectValue({'values': values})
    })
    return new_cell


def make_list_from_string(string):

    """Simple string handling, turns parts of a string separated by ', ' into a list of strings

    Args:
        string (str): String to sepparate on commas

    Returns:
        lst (list[str,]): list
    """

    if isinstance(string, str):
        lst = string.split(', ')
        return lst
    else:
        raise TypeError('Not a string')


def change_cell_in_row(sheet_id, row, lst, column_id, client):

    """writes new information to row object in python, then PUTs updated row over old row in smartsheet

    Args:
        sheet_id (int): Sheet ID
        row (Row): Row object to change
        lst (list): List of values to append to cell
        column_id (int): Column ID of cell to edit
        client (Client): Smartsheet client

    Returns:
        Void
    """
    # Rebuild row cell by Cell
    new_row = []
    for cell in row.cells:
        if cell.column_id == column_id:  # This is the Cell that needs to be changed
            new_row.append(make_picklist_cell(column_id, lst))  # So append updated version

    # Rebuild row object
    row_object = client.models.Row({
        'id': row.id_,
        'cells': new_row
    })

    # POST row to smartsheet
    try:
        client.Sheets.update_rows(sheet_id, [row_object])
    except:
        print(f"couldn't write cell '{lst}' to row: {row.row_number}, column: '{column_id}' in sheet: "
              f"'{client.Sheets.get_sheet(sheet_id)}'")


def search(sheet, value):

    """Search a sheet for a specified value, returns the row that contains the value
    Args:
        sheet (Sheet): Sheet to search
        value (str): Search value
    Returns:
        row (Row): Row that contains search value
    """

    # Go through all rows and columns and return the row and column ID of cell you are looking for
    for row in sheet.rows:
        for cell in row.cells:
            if cell.value == value:
                return row
    raise ValueError(f"{value} not Found in sheet: '{sheet.name}'")


def build_dict(workspace_name, sheet_name, key_column, values_column, client):

    """Builds a dictionary of all the options from a multi picklist

    Args:
        workspace_name (str): Workspace name
        sheet_name (str): Sheet name
        key_column (str): Name of column containing dict keys
        values_column (str): Name of column containing dict values
        client (Client): Smartsheet client

    Returns:
        thisdict (dict(str, list(str,))): Dictionary containing keys and values
    """

    # Initiate sheet and column IDs
    sheet = init_sheet(workspace_name, sheet_name, client)

    # Error handling if incorrect column names
    if key_column not in [c.title for c in sheet.columns]:
        raise Exception(f"No column called: '{key_column}' found in sheet: '{sheet.name}'")
    if values_column not in [c.title for c in sheet.columns]:
        raise Exception(f"No column called: '{values_column}' found in sheet: '{sheet.name}'")

    keys_id = sheet.get_column_by_title(key_column).id
    vals_id = sheet.get_column_by_title(values_column).id

    # Go through sheet row by row to fill dict with keys from key column and values from value column
    thisdict = {}
    for row in sheet.rows:
        cell = row.get_column(keys_id)
        if cell.value in thisdict:  # If this key is already in the dict
            thisdict[cell.value].append(row.get_column(vals_id).value)  # Append the value to the value list
        elif row.get_column(vals_id).value:  # If the key is new, and there is data in the values column
            thisdict[cell.value] = make_list_from_string(row.get_column(vals_id).value)  # Add it to the dictionary
        else:  # Otherwise
            thisdict[cell.value] = []  # Add an empty entry

    return thisdict


def compare_dicts(workspace_name, this_sheet_name, this_key_column, this_data_column, that_sheet_name,
                  that_key_column, that_data_column):

    """Compares two multi dropdowns to see if the selections in one correlate to the rows of the other.
    I.E. if the parts present in an engine correlate to the engines which use that part.

    Args:
        workspace_name (str): Workspace name
        this_sheet_name (str): Name of one of the sheets to compare
        this_key_column (str): Name of column containing keys for first dict
        this_data_column (str): Name of column containing values for first dict
        that_sheet_name (str): Name of other sheet to compare
        that_key_column (str): Name of column containing keys for second dict
        that_data_column (str): Name of cloumn containing values for second dict

    Returns:
        Void
    """

    # initiate a client
    client = init_client()

    # First build both dicts to compare
    thisdict = build_dict(workspace_name, this_sheet_name, this_key_column, this_data_column, client)
    thatdict = build_dict(workspace_name, that_sheet_name, that_key_column, that_data_column, client)
    length1 = len(thisdict)
    length2 = len(thatdict)
    i = 0
    # Go through the first dict, to check that all the values in this dict
    for key in thisdict:
        for value in thisdict[key]:
            if value in thatdict:  # Are keys in the other dict
                if key in thatdict[value]:  # If so, is the key of the first dict also a value in this dict?
                    thatdict[value].remove(key)  # Remove it, so if they are matched, we are left with an empty dict
                else:  # If we find an entry which is unmatched, POST value to smartsheet cell
                    update_cell(workspace_name, that_sheet_name, that_data_column, key, value, client)
                    print(f'updated {i}th {length1} entries, from first dict')
                    i += 1
    # This dict should be empty
    for key in thatdict:
        if thatdict[key]:  # If not, POST value to smartsheet cell
            for value in thatdict[key]:
                update_cell(workspace_name, this_sheet_name, this_data_column, key, value, client)
                print(f'updated {i}th of {length2} entries, from second dict')
                i += 1


def update_cell(workspace_name, sheet, column_name, value, search_value, client):

    """Appends an option onto a multi picklist cell

    Args:
        workspace_name (str): Workspace name
        sheet (str): Sheet name
        column_name (str): Column name
        value (str): Value to append to cell
        search_value (str): Search value to find cell
        client (Client): Smartsheet client

    Returns:
        Void
    """

    # Initiate sheet and row object
    sheet = init_sheet(workspace_name, sheet, client)
    row = search(sheet, search_value)
    column = sheet.get_column_by_title(column_name)  # Turns column name into column ID

    # Find the old cell, and turn its value into a list
    old_cell = row.get_column(column.id)
    if old_cell.value:
        values = make_list_from_string(old_cell.value)
    else:
        values = []

    # Add new values to list and POST back to smartsheet
    values.append(value)
    change_cell_in_row(sheet.id, row, values, column.id, client)
