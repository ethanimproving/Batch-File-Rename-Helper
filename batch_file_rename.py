# Batch File Renamer
# Ethan Miller
# 5/29/2019

import subprocess
import xlsxwriter

# Get directory from user.
path=input('Enter directory to write file in: ')

# Manage file names.
filename = '_list'
textfile = filename+'.txt'
spreadsheet = filename+'.xlsx'

# Create _list.txt file with list of files in directory.
subprocess.call(f'cd {path} && dir /b > {textfile}', shell=True)

# Read contents of file.
with open(textfile, 'r') as infile:
  # Create a list lines from the textfile.
  lines = infile.readlines()

# Display each line for testing.
for line in lines:
  print(line.strip())

# Create an Excel Spreadsheet.
workbook = xlsxwriter.Workbook(spreadsheet)
worksheet = workbook.add_worksheet()

# Write each line to a new row in column A.
for line in lines:
  # Create headers.

  line_num = lines.index(line)+2
  worksheet.write(f'A{line_num}', line.strip())
  # Rename column A with column C.
  worksheet.write(f'D{line_num}', f'=CONCATENATE("ren """,A{line_num},""""," ","""",C{line_num},"""")')

# Close the workbook.
workbook.close()