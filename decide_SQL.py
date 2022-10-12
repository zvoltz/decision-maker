import sqlite3
import PySimpleGUI as sg

global cursor


# Set the global cursor to the connected database.
def setup_SQL():
    global cursor
    connection = sqlite3.connect(getFile())
    cursor = connection.cursor()


# Instead of doing setup_SQL and getting the file here, used to set the global cursor before calling main_SQL(). Takes
# a sqlite3 cursor as a parameter.
def set_cursor(cursor_param):
    global cursor
    cursor = cursor_param


# Get the database to use.
def getFile():
    layout = [[sg.Text('Select the database(.db) file:'), sg.InputText(), sg.FileBrowse(file_types=((
                'ALL Files', '*.*db'),))],
              [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Pick item from database', layout, keep_on_top=True)
    event, values = window.read()
    window.close()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        exit()
    return values[0]


# Return a list of the names of the columns after the 0th column
def get_column_names():
    cursor.execute("SELECT * FROM preferences")
    all_arr = list(cursor.description)
    # remove the 'object name' column
    all_arr.pop(0)
    column_names = []
    # cursor.description returns "returns a 7-tuple for each column where the last six items of each tuple are None."
    # Remove all the Nones:
    for tuple7 in all_arr:
        column_names += tuple7
    column_names = [x for x in column_names if x]

    return column_names


# Given the names of the columns, present the user with a GUI to check which people to consider when making the
# decision. Returns a list of boolean values that relate to whether that column should be considered where the
# boolean value at i relates to column i+1.
def select_people(column_names):
    layout = [[sg.Text("Check everyone to consider while deciding:")]]
    for x in column_names:
        layout.append([sg.Checkbox(x)])
    layout.append([sg.Button("Submit")])
    window = sg.Window("Decision Maker", layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            exit()
        if event == 'Submit':
            break
    window.close()

    return list(values.values())


# Returns a string of the selected people's names in the form:
# person1, person2, person3
# for the purpose of being used in SELECT statements
def get_people_SQL():
    column_names = get_column_names()
    values = select_people(column_names)
    wanted_names = "object_name"
    for i in range(len(column_names)):
        if values[i]:
            wanted_names += ", " + column_names[i]
    return wanted_names


def main_SQL():
    people = get_people_SQL()
    # possible SQL injection
    command = "SELECT " + people + " FROM preferences"
    res = cursor.execute(command)
    max_rating = people.count(',') * 3
    items = prep_list(max_rating)
    for row in res:
        index = max_rating - sum(row[1:])
        items[index].append(row[0])
    show_results(items)


# Returns a list of size max_rating + 1, each initialized as an empty list
# max_rating + 1 because range is 0 to max_rating
def prep_list(max_rating):
    preference_list = []
    for rank in range(max_rating + 1):
        preference_list.append([])
    return preference_list


# Given the items in a 2D array, where index 0 is a list of the highest rated items, display them in order, starting
# with the highest score and going down. The GUI shows all items with equal rank in the same window, and each new window
# indicates a single decrease in rank and a single increase in index unless an index has no items to display.
def show_results(items):
    index = 0
    while index < len(items):
        if not items[index]:
            index += 1
            continue
        index += show_window(", ".join(items[index]))


# Given the text to show as string of items, display them as a multiline GUI. Returns either -1 or 1 based on what the
# user selects. This integer represents whether the integer wants to see the previous index (-1) or next (+1).
def show_window(text):
    layout = [[sg.Multiline(text, s=(150, 2))], [sg.Button("Back"), sg.Button("Next"), sg.Button("Cancel")]]
    window = sg.Window("Decision Maker", layout)
    event, values = window.read()
    window.close()
    match event:
        case sg.WIN_CLOSED | 'Cancel':
            exit()
        case 'Back':
            return -1
        case 'Next':
            return 1


def start():
    setup_SQL()
    main_SQL()


if __name__ == '__main__':
    start()
