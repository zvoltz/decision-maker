import openpyxl as sheet
import PySimpleGUI as sg
import sqlite3

import os

colors = {
    "FF0000": 0,
    "FFFF00": 1,
    "000000": 2,
    "FFFFFF": 2,
    "00FF00": 3
}
global rows
global names


# Show GUI to get and return the absolute path to the Excel file.
def get_file():
    layout = [[sg.Text('Select the Excel(.xlsx) file:'), sg.InputText(), sg.FileBrowse(file_types=((
                'ALL Files', '*.*xlsx'),))],
              [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Create Database From Excel', layout, keep_on_top=True)

    event, values = window.read()
    window.close()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        exit()
    return values[0]


# Show GUI to get and return the name of the new database file.
def get_file_name():
    layout = [[sg.InputText(), sg.FileSaveAs(file_types=(('ALL Files', '*.db'),), default_extension='*.db')],
              [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Save the database', layout)

    event, values = window.read()
    window.close()
    if event == sg.WIN_CLOSED or event == 'Cancel' or values[0] is None:
        exit()
    return values[0]


# Using the previously set excel_file variable, set global names and rows appropriately.
# rows becomes a list of items and preferences with no names row.
# names becomes a list of names, removing the first blank cell.
def get_data():
    global rows
    global names
    ws = sheet.load_workbook(filename=excel_file).active
    rows = list(ws.rows)
    names = list(rows[0])
    for i in range(len(names)):
        if not names[i].value and i != 0:
            names = names[0:i]
            break
    names.pop(0)
    rows.pop(0)


def make_database():
    remove_copy()
    connection = sqlite3.connect(file_name)
    cursor = connection.cursor()
    cursor.execute("CREATE TABLE preferences (object_name TEXT)")

    # create a column for each person's preferences
    for cell in names:
        if not cell.value:
            break
        cursor.execute(f"ALTER TABLE preferences ADD {cell.value} SMALLINT(1);")

    # add each item along with everyone's opinion towards that item
    for row in rows:
        add_row(cursor, row)
    connection.commit()
    connection.close()


# Check if a database file exists in the given path. If one does, delete it.
def remove_copy():
    if os.path.exists(file_name):
        os.remove(file_name)


def add_row(cursor, row):
    data = ["'" + str(row[0].value) + "'"]
    numQuestions = "?, " * (len(names) + 1)
    numQuestions = numQuestions[0:-2]
    for i in range(len(row[1:])):
        if i == len(names):
            break
        data.append(str(get_color_value(row[1:][i])))
    cursor.execute(f"INSERT INTO preferences VALUES ({numQuestions});", data)


# Returns an integer representation of a single person's opinion based on the color they put.
# Checks if the color is valid (see colors dictionary or overview), defaults to 2 if invalid.
def get_color_value(cell):
    color = cell.fill.start_color.index[2:]
    valid = colors.get(color, 2)
    return valid


if __name__ == '__main__':
    excel_file = get_file()
    while excel_file == "":
        excel_file = get_file()
    get_data()
    file_name = get_file_name()
    while file_name == "":
        file_name = get_file_name()
    make_database()
