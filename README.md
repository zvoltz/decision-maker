# decision-maker
## Overview
Given a file of preferences and a list of names, display the best items according to their preferences.

## File Format
The file is currently an Excel (.xlsx) file with A1 being blank. The rest of the A row is a list of names. The first column is a list of items. Their intersections are colored cells: red (FF0000) for disliking the item, yellow (FFFF00) for neutral, white/black (FFFFFF/000000) for no opinion, and green (00FF00) for liking it. Any other colors will cause errors.
The file should look like this:

![Example format](example_format.png)

## TODO
- Add GUI.
- Change file format to SQL.
- Remove reference to games (generalize to work for all preferences like restaurants).
- Implement support for comments about the each person's preference.

## License
Licensed under GPL V3. See [LICENSE](LICENSE) for more details.
