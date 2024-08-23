import pygetwindow as gw

# Get a list of all window titles
windows = gw.getAllTitles()

# Filter out empty titles (which often represent background windows)
windows = [title for title in windows if title]

# Print all window titles
for title in windows:
    print(title)