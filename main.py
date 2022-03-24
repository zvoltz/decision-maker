"""
Given a file of preferences and a list of names, display the best items according to their preferences

The file is currently an Excel (.xlsx) file with A1 being blank. The rest of the A row is a list of names of people.
The first column is a list of items. Their intersections are colored cells: red (FF0000) for disliking the item,
yellow (FFFF00) for neutral, white/black (FFFFFF/000000) for no opinion, and green (00FF00) for liking it.
Any other colors will cause errors.

TODO:
1. Add GUI
2. Add SQL support
3. Remove reference to games (generalize to work for all preferences like restaurants)
4. Implement comment support
"""
import openpyxl as sheet


global rows
colors = {
    "FF0000": 0,
    "FFFF00": 1,
    "FFFFFF": 2,
    "000000": 2,
    "00FF00": 3
}


# Reads all rows of the given file and puts them in the 2D array rows
def setup():
    global rows
    ws = sheet.load_workbook(filename="example_file.xlsx").active
    rows = list(ws.rows)


# Ask the user for all names
# allows the user to type the same person's name multiple times to increase their influence on the listed items
def get_people():
    people = []
    person = input("Type the first person's name: ")
    while person != "":
        people.append(person)
        person = input("Type the next person's name or press ENTER to stop: ")
    return people


def main():
    weight_per_name = get_weight_per_name()
    max_rating = max(colors.values()) * sum(weight_per_name)
    items = prep_list(max_rating)
    for item in rows:
        index = max_rating - get_rating(item[1:], weight_per_name)
        items[index].append(item[0].value)
    show_results(items, max_rating)


# Returns a list of size max_rating + 1, each initialized as an empty list
# max_rating + 1 because range is 0 to max_rating
def prep_list(max_rating):
    preference_list = []
    for rank in range(max_rating + 1):
        preference_list.append([])
    return preference_list


# name weights is how much each person should affect the outcome. It'll be multiplied by their
# opinion of the item when creating the item ratings.
# Removes the first row of rows because the names of the people are no longer necessary.
def get_weight_per_name():
    people = get_people()
    weight_per_name = []
    for name in rows[0]:
        weight_per_name.append(people.count(name.value))
    # A1 is blank so the 1st "Name" is blank, pop to remove
    weight_per_name.pop(0)
    # First row is the list of the names, remove so the rows are just items
    rows.pop(0)
    return weight_per_name


# Returns an integer representation of everyone's opinion on that item.
# Ratings are calculated by multiplying their color coded opinion by their name weight.
def get_rating(item, weights):
    score = 0
    for index in range(len(item)):
        color = item[index].fill.start_color.index[2:]
        score += get_color_value(color) * weights[index]
    return score


# Returns an integer representation of a single person's opinion based on the color they put.
# Checks if the color is valid (see colors dictionary or overview), exits if it is not.
def get_color_value(color):
    valid = colors.get(color, None)
    if valid is None:
        print("Invalid color in sheet")
        exit()
    return valid


def show_results(items, max_rating):
    cont = ""
    index = 0
    while cont != "exit":
        if index >= len(items):
            return
        if not items[index]:
            index += 1
            continue
        print(f"With a score of {max_rating - index}:")
        print(items[index])
        cont = input("Press ENTER to see the next best list or type exit to exit: ")
        index += 1


if __name__ == '__main__':
    setup()
    main()
