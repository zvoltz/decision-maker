import openpyxl as sheet
import decider_parent


class ExcelDecider(decider_parent.Decider):
    rows = []
    colors = {
        "FF0000": 0,
        "FFFF00": 1,
        "FFFFFF": 2,
        "000000": 2,
        "00FF00": 3
    }

    # Reads all rows of the given file and puts them in the 2D array rows
    def setup(self, file=None):
        if file is None:
            file = decider_parent.get_file()
        ws = sheet.load_workbook(filename=file).active
        self.rows = list(ws.rows)

    # Return the list of names from the instance variable rows, then pop the first row.
    # Returns the names as a list of strings. Stops looking after a cell is none.
    def get_names(self):
        # First cell is blank, ignore it
        names = []
        for cell in self.rows[0][1:]:
            if cell.value is None:
                break
            names.append(cell.value)
        self.rows.pop(0)
        return names

    def main(self, file=None):
        if not self.rows:
            self.setup(file)
        people = self.select_people(self.get_names())
        max_rating = max(self.colors.values()) * sum(people)
        items = self.prep_list(max_rating)
        for item in self.rows:
            object_name = str(item[0].value).strip()
            if object_name is not None and not "":
                index = max_rating - self.get_rating(item[1:], people)
                items[index].append(object_name)
        self.show_results(items)

    # Returns an integer representation of everyone's opinion on that item.
    # Ratings are calculated by multiplying their color coded opinion by their name weight.
    def get_rating(self, item, weights):
        score = 0
        for index in range(len(weights)):
            color = item[index].fill.start_color.index[2:]
            score += self.get_color_value(color) * weights[index]
        return score

    # Returns an integer representation of a single person's opinion based on the color they put.
    # Checks if the color is valid (see colors dictionary or overview), defaults to 2.
    def get_color_value(self, color):
        return self.colors.get(color, 2)

    def __init__(self, file=None):
        super().__init__()
        self.main(file)


if __name__ == '__main__':
    ExcelDecider()
