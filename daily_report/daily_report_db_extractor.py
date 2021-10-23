import os
import glob
import json
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

# Sheet index number
SHT_IDX = 0
# Main dictionary which all data are stored in then parsed to json
main_dict = []
# Get the directory path
package_dir = os.path.dirname(os.path.abspath(__file__))
# Get the list of all excel files inside the directory excel_files
excel_files = glob.glob(f"{package_dir}/excel_files/*.xlsx")

def cell_coordinate(string):
    """cell_coordinate, gets the number of coordinates of the excel cell like
    'A15' and 'Z39' etc.

    Parameters
    ----------
    string : str
        Excel cell position like 'A17' as string.

    Returns
    -------
    dict
        Dictionary of x and y coordinates as the row number x and column
        number y
    """
    xy = coordinate_from_string(string)
    x = xy[1]
    y = column_index_from_string(xy[0])
    return {'x': x, 'y': y}


def get_parameters_mod(first_cell, last_cell, sht_idx):
    """get_parameters_mod which accept the excel cell values in letters like
    'A19' and converts it to its cell values in numbers then iterates over
    the selected range of cells and copy the cell values and store it in a
    dictionary and offset with 5 cells to the right and store it as well,make
    sure to select the correct starting cell (first cell) and end cell
    correctly.

    Parameters
    ----------
    first_cell : str
        first_cell
    last_cell : str
        last_cell
    sht_idx : int
        sht_idx

    Returns
    -------
    dict
        Returns a dictionary of extracted values inside a variable called well
        parameters in the form of title and value
    """

    first_coor     = cell_coordinate(first_cell)
    last_coor      = cell_coordinate(last_cell)
    well_parameter = {}

    for row in range(first_coor['x'], last_coor['x']):
        for col in range(first_coor['y'], first_coor['y'] + 1):
            c_title = wb[wb.sheetnames[sht_idx]].cell(row, col)
            c_value = wb[wb.sheetnames[sht_idx]].cell(row, col+5)

            if c_value.value is None:
                continue
            well_parameter[c_title.value] = c_value.value

    return well_parameter


def get_activities_mod(first_cell, last_cell, sht_idx, is_last: bool):
    """get_activities_mod, This scrip is used to get the activities report
    and append it to a list than can be used later to be added to a final
    report.

    Parameters
    ----------
    first_cell : str
        first_cell
    last_cell : str
        last_cell
    sht_idx : int
        sht_idx
    is_last : bool
        The boolean value is used to differentiate between getting the data
        from the last activity or the next activity so its easier and
        distinguish between the two function calls

    Returns
    -------
    list
        Returns a list of extracted values values in the given cell reference
    """

    first_coor    = cell_coordinate(first_cell)
    last_coor     = cell_coordinate(last_cell)
    activity_list = []
    for row in range(first_coor['x'], last_coor['x']):
        for col in range(first_coor['y'], first_coor['y']+1):
            c_lastActivities = wb[wb.sheetnames[sht_idx]].cell(row, col)
            if c_lastActivities.value is None:
                continue
            if is_last:
                activity_list.append(c_lastActivities.value)
            else:
                activity_list.append(c_lastActivities.value)
    return activity_list


for excel_file in excel_files:
    wb         = load_workbook(filename = excel_file, read_only = True)
    date       = wb.worksheets[SHT_IDX]["I5"].value
    client     = wb.worksheets[SHT_IDX]["F5"].value
    well       = wb.worksheets[SHT_IDX]["F7"].value
    jobID      = wb.worksheets[SHT_IDX]["I7"].value
    parameters = get_parameters_mod('I18', 'I38', SHT_IDX)

    if client == 'DNO':
        last_activity = get_activities_mod('A49', 'A56', 0, True)
        next_activity = get_activities_mod('A57', 'A64', 0, False)
    elif client == 'HKN':
        last_activity = get_activities_mod('A68', 'A77', 0, True)
        next_activity = get_activities_mod('A79', 'A88', 0, False)
    elif client == 'TAQA':
        last_activity = get_activities_mod('A60', 'A66', 0, True)
        next_activity = get_activities_mod('A70', 'A77', 0, False)
    else:
        last_activity = get_activities_mod('A62', 'A71', 0, True)
        next_activity = get_activities_mod('A72', 'A78', 0, False)

    finalReport = {
        "client":             client,
        "well":               well,
        "jobID":              jobID,
        "date":               date,
        "Well Parameters":    parameters,
        "last 24 activities": last_activity,
        "next 24 activities": next_activity,
        "file name":          excel_file,
    }
    main_dict.append(finalReport)

# Open a json file and dump the content of the dictionary inside it.
with open('file_test.json', 'a') as f:
    json.dump(main_dict, f, default=str)


