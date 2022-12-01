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
        if self.rows:
            # Already setup
            return
        if file is None:
            file = decider_parent.get_file()
        ws = sheet.load_workbook(filename=file).active
        self.rows = list(ws.rows)

    # Return the list of all names from the instance variable rows.
    # Returns the names as a list of strings. Stops looking after a cell is none.
    def get_names(self):
        # First cell is blank, ignore it
        names = []
        for cell in self.rows[0][1:]:
            if cell.value is None:
                break
            names.append(cell.value)
        return names

    def parse_data(self, people):
        max_rating = max(self.colors.values()) * len(people)
        items = self.prep_list(max_rating)
        for item in self.rows[1:]:
            object_name = str(item[0].value).strip()
            if object_name is not None and not "":
                index = max_rating - self.get_rating(item[1:], people)
                items[index].append(object_name)
        return items

    def main(self, file=None):
        self.setup(file)
        people = self.get_names()
        people = self.select_people(people)
        items = self.parse_data(people)
        self.show_results(items)

    # Returns an integer representation of everyone's opinion on that item.
    # Ratings are calculated by adding the correlated color value of all selected people.
    def get_rating(self, item, selected_people):
        score = 0
        for index in range(len(self.rows[0])):
            if self.rows[0][index].value in selected_people:
                color = item[index].fill.start_color.index[2:]
                score += self.get_color_value(color)
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
