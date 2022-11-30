import sqlite3
import decider_parent


class SQLDecider(decider_parent.Decider):
    cursor = None

    # Set the cursor to the connected database.
    def setup(self, file=None):
        if self.cursor:
            # Already setup
            return
        if file is None:
            decider_parent.get_file()
        connection = sqlite3.connect(file)
        self.cursor = connection.cursor()

    # Return a list of the names of the columns after the 0th column
    def get_names(self):
        self.cursor.execute("SELECT * FROM preferences")
        all_arr = list(self.cursor.description)
        # remove the 'object name' column
        all_arr.pop(0)
        column_names = []
        # cursor.description "returns a 7-tuple for each column where the last six items of each tuple are None."
        # Remove all the Nones:
        for tuple7 in all_arr:
            column_names += tuple7
        column_names = [x for x in column_names if x]

        return column_names

    # Given all the people to be considered, parse the data and return a list of items sorted such that the highest
    # rated items are at index 0, next highest are index 1, and so on.
    def parse_data(self, selected_people):
        all_items = self.select_SQL(selected_people)
        max_rating = len(selected_people) * 3
        sorted_items = self.prep_list(max_rating)
        for row in all_items:
            index = max_rating - sum(row[1:])
            sorted_items[index].append(row[0])
        return sorted_items

    # Given a list of names to consider, select them from the connected SQL database, and return the result.
    def select_SQL(self, selected_people):
        wanted_names = "object_name"
        for i in range(len(selected_people)):
            wanted_names += ", " + selected_people[i]
        # TODO: fix SQL injection
        command = "SELECT " + wanted_names + " FROM preferences"
        return self.cursor.execute(command)

    def main(self, file=None):
        self.setup(file)
        people = self.get_names()
        people = self.select_people(people)
        items = self.parse_data(people)
        self.show_results(items)

    def __init__(self, file=None):
        super().__init__()
        self.main(file)


if __name__ == '__main__':
    SQLDecider()
