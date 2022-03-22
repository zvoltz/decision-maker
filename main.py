"""
Given a file of preferences and a list of names, display the best items according to their preferences

The file is currently an Excel (.xlsx) file with A1 being blank. The rest of the A row is a list of names.
The first column is a list of items. Their intersections are colored cells: red (FF0000) for disliking the item,
orange (FF8000) for somewhat disliking it, yellow (FFFF00) for neutral, white/black (FFFFFF/000000) for no opinion,
lime (08FF00) for somewhat liking it, and green (00FF00) for liking it. Any other colors will cause errors.

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
    ws = sheet.load_workbook(filename="Games.xlsx").active
    rows = list(ws.rows)


# Ask the user for all names
# allows the user to type the same person's name multiple times to increase their influence on the listed items
def get_players():
    players = []
    person = input("Type a player's name: ")
    while person != "":
        players.append(person)
        person = input("Type the next player's name or press ENTER to stop: ")
    return players


def main():
    weight_per_name = get_weight_per_name()
    num_people = sum(weight_per_name)
    games = prep_list(num_people)
    for game in rows:
        index = 3 * num_people - get_game_score(game[1:], weight_per_name)
        games[index].append(game[0].value)
    show_results(games, num_people)


# Returns a list of size 3 * num_people + 1, each initialized as an empty list
# 3 * num_people + 1 because max score is 3x the number of people and min is 0.
def prep_list(num_people):
    preference_list = []
    for rank in range(3 * num_people + 1):
        preference_list.append([])
    return preference_list


# name weights is how much each person should affect the outcome. It'll be multiplied by their
# opinion of the game when creating the game scores.
# Removes the first row of rows because the names of the people are no longer necessary.
def get_weight_per_name():
    players = get_players()
    weight_per_name = []
    for name in rows[0]:
        weight_per_name.append(players.count(name.value))
    # 1st "Name" is the blank spot for the game names, pop the 1st name
    weight_per_name.pop(0)
    # names are not part of the games, remove them after accessing them
    rows.pop(0)
    return weight_per_name


# Returns an integer representation of everyone's opinion on that item.
# Scores are calculated by multiplying their color coded opinion by their name weight.
def get_game_score(game, weights):
    score = 0
    for index in range(len(game)):
        color = game[index].fill.start_color.rgb[2:]
        score += colors[color] * weights[index]
    return score


def show_results(games, num_people):
    cont = ""
    index = 0
    while cont != "exit":
        if index >= len(games):
            return
        if not games[index]:
            index += 1
            continue
        print(f"With a score of {2 * num_people - index}:")
        print(games[index])
        cont = input("Press ENTER to see the next best list or type exit to exit: ")
        index += 1


if __name__ == '__main__':
    setup()
    main()
