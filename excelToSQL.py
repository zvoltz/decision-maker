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


def getFile():
    layout = [[sg.Text('Select the excel file:', size=(15, 1)), sg.InputText(), sg.FileBrowse(file_types=((
                'ALL Files', '*.*xlsx'),))],
              [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Create Database From Excel', layout, keep_on_top=True)

    event, values = window.read()
    window.close()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        exit()
    return values[0]


def getFileName():
    layout = [[sg.InputText(), sg.FileSaveAs(file_types=(('ALL Files', '*.db'),), default_extension='*.db')],
              [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Save the database', layout)

    event, values = window.read()
    window.close()
    if event == sg.WIN_CLOSED or event == 'Cancel' or values[0] is None:
        exit()
    return values[0]


def getData(file):
    global rows
    global names
    ws = sheet.load_workbook(filename=file).active
    rows = list(ws.rows)
    names = list(rows[0])
    names.pop(0)
    rows.pop(0)
    return names,


def makeDatabase(file_name):
    removeCopy(file_name)
    connection = sqlite3.connect(file_name)
    cursor = connection.cursor()
    cursor.execute("CREATE TABLE preferences (object_name TEXT)")

    # create a column for each person's preferences
    for cell in names:
        cursor.execute(f"ALTER TABLE preferences ADD {cell.value} TEXT;")

    # add each item along with everyone's opinion towards that item
    for row in rows:
        addRow(cursor, row)
    connection.commit()
    connection.close()


def removeCopy(file_name):
    if os.path.exists(file_name):
        os.remove(file_name)


def addRow(cursor, row):
    data = ["'" + str(row[0].value) + "'"]
    numQuestions = "?, " * (len(row))
    numQuestions = numQuestions[0:-2]
    for cell in row[1:]:
        data.append(str(get_color_value(cell)))
    cursor.execute(f"INSERT INTO preferences VALUES ({numQuestions});", data)


# Returns an integer representation of a single person's opinion based on the color they put.
# Checks if the color is valid (see colors dictionary or overview), defaults to 2 if invalid.
def get_color_value(cell):
    color = cell.fill.start_color.index[2:]
    valid = colors.get(color, 2)
    return valid


if __name__ == '__main__':
    excel_file = getFile()
    while excel_file == "":
        excel_file = getFile()
    getData(excel_file)
    file_name = getFileName()
    while file_name == "":
        file_name = getFileName()
    makeDatabase(file_name)


