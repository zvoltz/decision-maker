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
