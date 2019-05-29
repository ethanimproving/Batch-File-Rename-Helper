# Batch File Renamer
# Ethan Miller
# 5/29/2019

import subprocess
import xlsxwriter

# Get directory from user.
path=input('Enter directory to write file in: ')

# Create _list.txt file with list of files in directory.
textfile = '_list.txt'
subprocess.call(f'cd {path} && dir /b > {textfile}', shell=True)

# Read contents of file
with open(textfile, 'r') as infile:
  # Create a list lines from the textfile.
  lines = infile.readlines()

# Display each line for testing.
for line in lines:
  print(line.strip())