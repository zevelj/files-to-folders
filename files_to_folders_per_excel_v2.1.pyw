import os
import shutil
import pandas as pd
from tkinter import Tk, messagebox, filedialog

# Create a Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window

# Display a message dialog to select the files folder
messagebox.showinfo("Select Files Folder", "In the next window, please select the folder containing files to be organized.")
files_directory = filedialog.askdirectory()

# Display a message dialog to select the Excel file
messagebox.showinfo("Select Excel File", "In the next window, please select the Excel file with the organization list.")
excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])

# Read the Excel file
data_frame = pd.read_excel(excel_file_path)

# Convert the header values to lowercase
data_frame.columns = data_frame.columns.str.lower()

# Check if the headers exist with case-insensitive and plural forms
file_header = next((col for col in data_frame.columns if col.lower() in ['file', 'files']), None)
folder_header = next((col for col in data_frame.columns if col.lower() in ['folder', 'folders']), None)

if file_header is None or folder_header is None:
    messagebox.showerror("Header Not Found", "Required headers 'File' and 'Folder' were not found in the Excel file.")
else:
    # Check if the "_ORGANIZED_FOLDER" already exists in the files folder
    organized_folder_path = os.path.join(files_directory, "_ORGANIZED_FOLDER")
    if os.path.exists(organized_folder_path):
        # Prompt the user to merge or create a new folder
        answer = messagebox.askyesno("Folder Already Exists",
                                     "It seems an organized folder already exists. Do you want to merge the current files with the existing organized folder?\n\nSelecting 'No' will create an additional folder and leave the existing folder intact.\n\nSelecting 'Yes' will merge the current folder structure with the existing one.",
                                     default='no')

        if answer:
            # Merge with existing folder
            pass  # Add the logic for merging the folder structure here
        else:
            # Find the available folder name
            new_folder_name = "_ORGANIZED_FOLDER"
            counter = 1
            while os.path.exists(os.path.join(files_directory, new_folder_name)):
                new_folder_name = f"_ORGANIZED_FOLDER_{counter}"
                counter += 1

            # Create a new folder
            organized_folder_path = os.path.join(files_directory, new_folder_name)
            os.makedirs(organized_folder_path)

    else:
        # Create the "_ORGANIZED_FOLDER" within the selected files folder
        os.makedirs(organized_folder_path)

    # Create folders and copy files
    for index, row in data_frame.iterrows():
        file_name = row[file_header]
        target_folder_path = row[folder_header]

        # Create the target folder path within "_ORGANIZED_FOLDER"
        target_folder_path = os.path.join(organized_folder_path, target_folder_path)

        # Create the parent folders if they don't exist
        os.makedirs(target_folder_path, exist_ok=True)

        # Copy the file to the target folder
        source_file_path = os.path.join(files_directory, file_name)
        target_file_path = os.path.join(target_folder_path, file_name)
        shutil.copy(source_file_path, target_file_path)

    # Show a success message
    messagebox.showinfo("Success", "Files organized successfully.")

# Destroy the root window
root.destroy()
