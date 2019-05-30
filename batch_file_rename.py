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
workbook = xlsxwriter.Workbook({spreadsheet})
worksheet = workbook.add_worksheet()

# Write each line to a new row in column A.
for line in lines:
  # Create headers.

  line_num = lines.index(line)+2
  worksheet.write(f'A{line_num}', line.strip())
  # Rename column A with column C.
  worksheet.write(f'D{line_num}', f'=CONCATENATE("ren """,A{line_num},""""," ","""",C{line_num},"""")')

# Create a new spreadsheet to hold example string fuctions.
example = workbook.add_worksheet('string functions')

# Widen the first column to make the text clearer.
example.set_column('A:A', 16)
example.set_column('B:B', 48)
example.set_column('C:C', 18)

# Insert text into first column.
example_text = [
  [
    'Function',
    'String',
    'Output'
  ],
  [
    'Trailing Zeroes: ',
    '58 - Manliness.mp4',
    '=TEXT(LEFT(B2,3)+1,"000")&RIGHT(B2,LEN(B2)-2)'
  ],
  [
    'Number from #: ',
    '#43 "Marching to Zion" (Soul-stirring Songs & Hymns)',
    '=MID(B3,FIND("#",B3)+1,FIND("""",B3)-FIND("#",B3)-2)+0'
  ],
  [
    'String from "": ',
    '#43 "Marching to Zion" (Soul-stirring Songs & Hymns)',
    '=MID(B4,FIND("""",B4)+1,FIND("""",B4,FIND("""",B4)+1)-FIND("""",B4)-1)'
  ]
]

# Insert the example_text table into Excel.
for row in example_text:
  line_num = example_text.index(row)+1
  example.write(f'A{line_num}', row[0])
  example.write(f'B{line_num}', row[1])
  example.write(f'C{line_num}', row[2])

# Close the workbook.
workbook.close()