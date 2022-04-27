#! /usr/bin/env python3
"""
File script for finding each file and folder and add it
to the excel sheet
"""
import os
import xlwings as xw

# from rich import print


# Get the directory path and file path
package_dir = os.path.dirname(os.path.abspath(__file__))
excel_file = os.path.join(package_dir, "report.xlsx")

list_of_files = []
list_of_dirs = []

"""
Loop through each file inside each folder and get
the list of files names and directory names
"""
for root, dirs, files in os.walk(package_dir):
    for dir_name in dirs:
        list_of_dirs.append(dir_name)

    for file in files:
        # Confirm to catch only the right image type
        if file.endswith(".jpg"):
            list_of_files.append(os.path.join(root, file))

# Sort the list of directories and files
list_of_dirs.sort()
list_of_files.sort()

# Open the excel sheet and make it ready for input
op_book = xw.Book(excel_file)
op_book.app.visible = False

"""
The below loop will be used to open the excel sheet and create
sheets for each folder and then add the corresponding images
to it
"""
for dir_n in list_of_dirs:
    # Create the sheet and add it by its name
    wo = op_book.sheets.add(dir_n)
    """
    Find the correct image for the folder and add a / at the end
    so the search will be matching the directory folder
    """
    matching = [match for match in list_of_files if str(dir_n + "/") in match]
    # Looping on each match and add it to the excel sheet
    for image in matching:
        wo.pictures.add(image, scale=0.4)
