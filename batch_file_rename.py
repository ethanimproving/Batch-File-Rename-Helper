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
with open(f'{path}/{textfile}', 'r') as infile:
  # Create a list lines from the textfile.
  lines = infile.readlines()

# Display each line for testing.
for line in lines:
  print(line.strip())

# Create an Excel Spreadsheet.
workbook = xlsxwriter.Workbook(f'{path}/{spreadsheet}')
worksheet = workbook.add_worksheet()

# Get length of longest line.
longest_line = len(max(lines, key=len))
line_length = longest_line/1.025  # Offset for varrying character widths
cmd_length = line_length + 7.43   # ren "" "" has a cell width of 7.43

# Widen columns to fit lines.
worksheet.set_column('A:A', line_length)
worksheet.set_column('C:C', 10.14)
worksheet.set_column('D:D', cmd_length)

# Write each line to a new row in column A.
for line in lines:
  # Create headers.
  worksheet.write('A1', 'Old Name')
  worksheet.write('C1', 'New Name')
  worksheet.write('D1', 'Command Line')

  # Insert our list of lines to column A.
  line_num = lines.index(line)+2
  worksheet.write(f'A{line_num}', line.strip())

  # Rename column A with column C.
  worksheet.write(f'D{line_num}', f'=CONCATENATE("ren """,A{line_num},""""," ","""",C{line_num},"""")')

# Create a new spreadsheet to hold example string fuctions.
example = workbook.add_worksheet('string functions')

# Widen columns to make the text clearer.
example.set_column('A:A', 16)
example.set_column('B:B', 48)
example.set_column('C:C', 18)

# Insert text into first column.
example_text = [
  [
    'Function',
    'String',
    'Output',
    'Formula'
  ],
  [
    'Trailing Zeroes: ',
    '58 - Manliness.mp4',
    '=TEXT(LEFT(B2,2)+1,"000")&RIGHT(B2,LEN(B2)-2)',
    '\'=TEXT(LEFT(text,num_char)+index_offset,"000")&RIGHT(text,LEN(text)num_char)'
  ],
  [
    'Number from #: ',
    '#43 "Marching to Zion" (Soul-stirring Songs & Hymns)',
    '=MID(B3,FIND("#",B3)+1,FIND("""",B3)-FIND("#",B3)-2)+0',
    '\'=MID(text,FIND(front_character,text)+1,FIND(back_character,text)-FIND(front_character,text)-2)+0'
  ],
  [
    'String from "": ',
    '#43 "Marching to Zion" (Soul-stirring Songs & Hymns)',
    '=MID(B4,FIND("""",B4)+1,FIND("""",B4,FIND("""",B4)+1)-FIND("""",B4)-1)',
    '\' =MID(text,FIND(first_occurence,text)+1,FIND(second_occurence,text,FIND(first_occurence,text)+1)-FIND(first_occurence,text)-1)'
  ],
  [
    'Split at character:',
    '011016_Assessment.xlsx',
    '=LEFT(B5,FIND("_",B5)-1)',
    '\'=LEFT(text,FIND(character,text)-1)'
  ],
  [
    'Split at character:',
    '011016_Assessment.xlsx',
    '=RIGHT(B6,LEN(B6)-FIND("_",B6))',
    '\'=RIGHT(text,LEN(text)-FIND(character,text))'
  ],
  [
    'String from ():',
    'Jones (60)',
    '=MID(B7,SEARCH("(",B7)+1,SEARCH(")",B7)-SEARCH("(",B7)-1)+0',
    '\'=MID(text,SEARCH(first_character,text)+1,SEARCH(second_character,text)-SEARCH(first_character,text)-1)'
  ],
  [
    'nth word:',
    'This is the fourth word.',
    '=TRIM(MID(SUBSTITUTE(B8," ",REPT(" ",LEN(B8))), (4-1)*LEN(B8)+1, LEN(B8)))',
    '\'=TRIM(MID(SUBSTITUTE(text," ",REPT(" ",LEN(text))), (N-1)*LEN(text)+1, LEN(text)))'
  ]
]

# Insert the example_text table into Excel.
for row in example_text:
  line_num = example_text.index(row)+1
  example.write(f'A{line_num}', row[0])
  example.write(f'B{line_num}', row[1])
  example.write(f'C{line_num}', row[2])
  example.write(f'E{line_num}', row[3])

# Close the workbook.
workbook.close()