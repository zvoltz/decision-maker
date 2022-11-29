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

    # Returns a string of the selected people's names in the form:
    # person1, person2, person3
    # for the purpose of being used in SELECT statements
    def get_people(self):
        column_names = self.get_names()
        values = self.select_people(column_names)
        wanted_names = "object_name"
        for i in range(len(column_names)):
            if values[i]:
                wanted_names += ", " + column_names[i]
        return wanted_names

    def main(self, file=None):
        self.setup(file)
        people = self.get_people()
        # possible SQL injection
        command = "SELECT " + people + " FROM preferences"
        res = self.cursor.execute(command)
        max_rating = people.count(',') * 3
        items = self.prep_list(max_rating)
        for row in res:
            index = max_rating - sum(row[1:])
            items[index].append(row[0])
        self.show_results(items)

    def __init__(self, file=None):
        super().__init__()
        self.main(file)


if __name__ == '__main__':
    SQLDecider()
