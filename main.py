"""
Given a file of preferences and a list of names, display the best items according to their preferences.

The program can currently use both Excel (.xlsx) and database (.db) files.

An Excel (.xlsx) file should start with A1 being blank. The rest of the A row is a list of names. The first column is a
list of items. Their intersections are colored cells: red (FF0000) for disliking the item, yellow (FFFF00) for neutral,
white/black (FFFFFF/000000) for no opinion, and green (00FF00) for liking it. Any other colors will be counted as no
opinion.

A database file should be made using the excel_to_SQL file, but if you want to make your own, the first column should be
named "object_name" and type text. Each subsequent column should be the name of each person and type smallint(1). Each
row consists of the name of an item and each person's preference of the item in the form of a number.

TODO:
1. Implement support for comments about each person's preference.
2. Add support for android.
3. Add way to make a new database from scratch.
4. Add web support.
"""


def run_decider():
    import decider_parent
    file = decider_parent.get_file()
    if file[-3:] == '.db':
        from decide_SQL import SQLDecider
        SQLDecider(file)
    elif file[-5:] == '.xlsx':
        from decide_excel import ExcelDecider
        ExcelDecider(file)
    else:
        decider_parent.show_error("File not recognized")


if __name__ == '__main__':
    run_decider()
