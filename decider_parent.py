import PySimpleGUI as sg


class Decider:

    # Show GUI to get and return the absolute path to the Excel file.
    def get_file(self):
        layout = [[sg.Text('Select the Excel(.xlsx) file:'), sg.InputText(), sg.FileBrowse(file_types=((
                    'ALL Files', '*.*xlsx'),))],
                  [sg.Submit(), sg.Cancel()]]

        window = sg.Window('Decision Maker', layout, keep_on_top=True)

        event, values = window.read()
        window.close()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            exit()
        return values[0]

    # Given the names of the columns, present the user with a GUI to check which people to consider when making the
    # decision. Returns a list of boolean values that relate to whether that column should be considered where the
    # boolean value at i relates to column i+1.
    def select_people(self, names):
        layout = [[sg.Text("Check everyone to consider while deciding:")]]
        for x in names:
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

    # Returns a list of size max_rating + 1, each initialized as an empty list
    # max_rating + 1 because range is 0 to max_rating
    def prep_list(self, max_rating):
        preference_list = []
        for rank in range(max_rating + 1):
            preference_list.append([])
        return preference_list

    # Given the items in a 2D array, where index 0 is a list of the highest rated items, display them in order, starting
    # with the highest score and going down. The GUI shows all items with equal rank in the same window, and each new
    # window indicates a single decrease in rank and a single increase in index unless an index has no items to display.
    def show_results(self, items):
        index = 0
        while index < len(items):
            if not items[index]:
                index += 1
                continue
            index += self.show_window(", ".join(items[index]))

    # Given the text to show as string of items, display them as a multiline GUI. Returns either -1 or 1 based on what
    # the user selects. This integer represents whether the integer wants to see the previous index (-1) or next (+1).
    def show_window(self, text):
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

    # Given an error message, pop up an error window and exit.
    def show_error(self, error=None):
        if error is not None:
            sg.popup_error(error)
        else:
            sg.popup_error("Unspecified error")
        exit()

    def __init__(self):
        pass


if __name__ == '__main__':
    exit()