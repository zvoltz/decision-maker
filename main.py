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

layout = [[sg.Text("Do you have an Excel(.xlsx) or database (.db) to decide from?")],
          [sg.Button("Excel"), sg.Button("Database")]]
window = sg.Window("Select File Type", layout)
event, values = window.read()
window.close()
if event == sg.WIN_CLOSED or event == 'Cancel':
    exit()
elif event == "Excel":
    import decide_excel
    decide_excel.start()
elif event == "Database":
    import decide_SQL
    decide_SQL.start()
