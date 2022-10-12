"""
Given a file of preferences and a list of names, display the best items according to their preferences

The file is currently an Excel (.xlsx) file with A1 being blank. The rest of the A row is a list of names of people.
The first column is a list of items. Their intersections are colored cells: red (FF0000) for disliking the item,
yellow (FFFF00) for neutral, white/black (FFFFFF/000000) for no opinion, and green (00FF00) for liking it.
Any other colors will cause errors.

TODO:
1. Add GUI.
2. Remove reference to games (generalize to work for all preferences like restaurants).
3. Implement support for comments about each person's preference.
4. Add support for android.
5. Update README.
6. Add way to make a new database from scratch.
"""
import PySimpleGUI as sg


# Show GUI to get and return the absolute path to the Excel file.
def get_file():
    layout = [[sg.Text('Select the Excel(.xlsx) or database(.db) file:'), sg.InputText(), sg.FileBrowse(file_types=((
                'ALL Files', '*.*xlsx'), ('ALL Files', '*.*db')))],
              [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Decision Maker', layout, keep_on_top=True)

    event, values = window.read()
    window.close()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        exit()
    return values[0]


def run_decider():
    file = get_file()
    if file[-3:] == '.db':
        call_SQL(file)
    elif file[-5:] == '.xlsx':
        call_excel(file)
    else:
        error()


def call_SQL(file):
    import decide_SQL
    decide_SQL.start(file)


def call_excel(file):
    import decide_excel
    decide_excel.start(file)


def error():
    sg.popup_error("File not recognized")
    run_decider()


if __name__ == '__main__':
    run_decider()
