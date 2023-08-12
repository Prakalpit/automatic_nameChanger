import os
from openpyxl import load_workbook
from natsort import natsorted

# Load the Excel workbook
workbook = load_workbook('') # location of the excel file

# Choose the specific sheet you want to work with
sheet = workbook['Sheet1']  # Replace 'Sheet1' with the actual sheet name

# Directory containing the files
directory = '' # location of the directory

# Get the list of files in the directory (excluding hidden files)
files_in_directory = [f for f in os.listdir(directory) if not f.startswith('.')]
files_in_directory = natsorted(files_in_directory)  # Natural sort the filenames

# Get the list of new filenames from the Excel sheet
new_filenames = [cell.value for cell in sheet['A'][1:] if cell.value]  # Filter out empty values
new_filenames = natsorted(new_filenames)  # Natural sort the names

# Ensure the number of files and names match
if len(files_in_directory) != len(new_filenames):
    print("Number of files and names don't match.")
else:
    # Rename files in directory order and rename according to Excel sheet order
    for new_filename, old_filename in zip(new_filenames, files_in_directory):
        extension = os.path.splitext(old_filename)[1]
        new_path = os.path.join(directory, new_filename + extension)
        old_path = os.path.join(directory, old_filename)

        os.rename(old_path, new_path)
        print(f"Renamed {old_filename} to {new_filename + extension}")

    print("Files renamed and arranged successfully!")
