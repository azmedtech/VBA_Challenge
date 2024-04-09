# Modules
import os
import csv

# Prompt user for title lookup

# Set path for file
csvpath = os.path.join("..", "Resources", "comic_books.csv")

# Set variable to check if we found the comic book

# Open the CSV using the UTF-8 encoding
with open(csvpath, encoding='UTF-8') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=",")

    # Loop through looking for the comic book

            # Set variable to confirm we have found the comic book

    # If the comic book is never found, alert the user
