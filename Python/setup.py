import sys
import os


# Add the paths to the files you want to include
files_to_include = [
    'gui.py',
    'docparser.py',
    'JsonToExcel.py',
    'updateAccountHistory.py',
    'accountHistory.json',  # Include the JSON file
    # Add more files here if needed
]

# Specify the options for PyInstaller
options = [
    '--onefile',  # Create a one-file bundled executable
    # Do not show a console window when running the executable (optional)
    '--windowed',
]

# Create the PyInstaller command
command = ['pyinstaller'] + options + files_to_include

# Run PyInstaller
os.system(' '.join(command))

# Move the data file to the "dist" directory
dist_dir = os.path.join('dist', 'accountHistory.json')
os.rename('accountHistory.json', dist_dir)
