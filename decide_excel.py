import openpyxl as sheet
import PySimpleGUI as sg


class ExcelDecider:
    rows = []
    colors = {
        "FF0000": 0,
        "FFFF00": 1,
        "FFFFFF": 2,
        "000000": 2,
        "00FF00": 3
    }

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

    # Reads all rows of the given file and puts them in the 2D array rows
    def setup(self, file=None):
        if file is None:
            file = self.get_file()
        ws = sheet.load_workbook(filename=file).active
        self.rows = list(ws.rows)

    # Using the first row of the Excel file, present a GUI checkbox to check who to consider while making the decision.
    # Returns an array of boolean values that relate to the people to consider where array[i] relates to Excel cell
    # A i+2.
    def select_people(self):
        layout = [[sg.Text("Check everyone to consider while deciding:")]]
        for x in self.rows[0][1:]:
            layout.append([sg.Checkbox(x.value)])
        layout.append([sg.Button("Submit")])
        window = sg.Window("Decision Maker", layout)
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                exit()
            if event == 'Submit':
                break
        window.close()
        self.rows.pop(0)
        return list(values.values())

    def main(self, file=None):
        if not self.rows:
            self.setup(file)
        people = self.select_people()
        max_rating = max(self.colors.values()) * sum(people)
        items = self.prep_list(max_rating)
        for item in self.rows:
            index = max_rating - self.get_rating(item[1:], people)
            items[index].append(str(item[0].value))
        self.show_results(items)

    # Returns a list of size max_rating + 1, each initialized as an empty list
    # max_rating + 1 because range is 0 to max_rating
    def prep_list(self, max_rating):
        preference_list = []
        for rank in range(max_rating + 1):
            preference_list.append([])
        return preference_list

    # Returns an integer representation of everyone's opinion on that item.
    # Ratings are calculated by multiplying their color coded opinion by their name weight.
    def get_rating(self, item, weights):
        score = 0
        for index in range(len(item)):
            color = item[index].fill.start_color.index[2:]
            score += self.get_color_value(color) * weights[index]
        return score

    # Returns an integer representation of a single person's opinion based on the color they put.
    # Checks if the color is valid (see colors dictionary or overview), exits if it is not.
    def get_color_value(self, color):
        valid = self.colors.get(color, None)
        if valid is None:
            sg.popup_error("Invalid color in sheet")
            exit()
        return valid

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

    def __init__(self, file=None):
        self.main(file)


if __name__ == '__main__':
    ExcelDecider()
