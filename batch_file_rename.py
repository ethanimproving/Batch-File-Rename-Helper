# Batch File Renamer
# Ethan Miller
# 5/29/2019

import xlsxwriter
import os

class List:
  '''
  A list of files in the directory.
  '''
  def __init__(self):
    # Get directory from user.
    pathin=input('Enter directory to write file in: ')
    # Replace escaape characters with forward slashes.
    self.path = pathin.replace('\\\\', '/')
    # Create a list to hold file names.
    self.lines=[f for f in os.listdir(self.path) if os.path.isfile(os.path.join(self.path, f))]
    # Display each line for testing.
    for line in self.lines:
      print(line.strip())

  def generateWorkbook(self):
    lines = self.lines
    path = self.path
    # Create an Excel Spreadsheet.
    spreadsheet = '_list.xlsx'
    workbook = xlsxwriter.Workbook(f'{path}/{spreadsheet}')
    worksheet = workbook.add_worksheet()

    # Get length of longest line.
    longest_line = len(max(lines, key=len))
    line_length = longest_line/1.025  # Offset for varrying character widths
    cmd_length = line_length + 7.43   # ren "" "" has a cell width of 7.43

    # Widen columns to fit lines.
    worksheet.set_column('A:A', line_length)
    worksheet.set_column('C:C', 12.43)
    worksheet.set_column('D:D', cmd_length)

    # Format as table
    count = len(lines)+1
    worksheet.add_table(f'A1:D{count}', {'style': 'Table Style Medium 4',
                                'columns': [{'header': 'Old Name'},
                                            {'header': 'String'},
                                            {'header': 'New Name'},
                                            {'header': 'Command Line'},
                                            ]})

    # FROMAT THE WORKSHEET
    # Hide unused rows.
    worksheet.set_default_row(hide_unused_rows=True)
    # Show 10 unused rows.
    for row in range(count, count+10):
      worksheet.set_row(row, None, None, {'hidden': False})
    # Hide all columns past G.
    worksheet.set_column('G:XFD', None, None, {'hidden': True})

    # Write each line to a new row in column A.
    for line in lines:
      # Insert our list of lines to column A.
      line_num = lines.index(line)+2
      worksheet.write(f'A{line_num}', line.strip())

      # Rename column A with column C.
      worksheet.write(f'D{line_num}', f'=CONCATENATE("ren """,A{line_num},""""," ","""",C{line_num},"""")')

    # Create a new spreadsheet to hold example string fuctions.
    example = workbook.add_worksheet('string functions')

    # FROMAT THE WORKSHEET
    # Hide gridlines.
    example.hide_gridlines(option=2)
    # Hide unused rows.
    example.set_default_row(hide_unused_rows=True)
    # Show 7 unused rows and a spacer.
    example.set_row(1, None, None, {'hidden': False})
    for row in range(11, 18):
      example.set_row(row, None, None, {'hidden': False})
    # Hide all columns past G.
    example.set_column('G:XFD', None, None, {'hidden': True})


    # Insert text into first column.
    example_text = [
      [
        'Trailing Zeroes: ',
        '58 - Manliness.mp4',
        '=TEXT(LEFT(C4,2)+1,"000")&RIGHT(C4,LEN(C4)-2)',
        '\'=TEXT(LEFT(text,num_char)+index_offset,"000")&RIGHT(text,LEN(text)num_char)'
      ],
      [
        'Number from #: ',
        '#43 "Marching to Zion" (Soul-stirring Songs & Hymns)',
        '=MID(C5,FIND("#",C5)+1,FIND("""",C5)-FIND("#",C5)-2)+0',
        '\'=MID(text,FIND(front_character,text)+1,FIND(back_character,text)-FIND(front_character,text)-2)+0'
      ],
      [
        'String from "": ',
        '#43 "Marching to Zion" (Soul-stirring Songs & Hymns)',
        '=MID(C6,FIND("""",C6)+1,FIND("""",C6,FIND("""",C6)+1)-FIND("""",C6)-1)',
        '\' =MID(text,FIND(first_occurence,text)+1,FIND(second_occurence,text,FIND(first_occurence,text)+1)-FIND(first_occurence,text)-1)'
      ],
      [
        'Split at character:',
        '011016_Assessment.xlsx',
        '=LEFT(C7,FIND("_",C7)-1)',
        '\'=LEFT(text,FIND(character,text)-1)'
      ],
      [
        'Split at character:',
        '011016_Assessment.xlsx',
        '=RIGHT(C8,LEN(C8)-FIND("_",C8))',
        '\'=RIGHT(text,LEN(text)-FIND(character,text))'
      ],
      [
        'String from ():',
        'Jones (60)',
        '=MID(C9,SEARCH("(",C9)+1,SEARCH(")",C9)-SEARCH("(",C9)-1)+0',
        '\'=MID(text,SEARCH(first_character,text)+1,SEARCH(second_character,text)-SEARCH(first_character,text)-1)'
      ],
      [
        'nth word:',
        'This is the fourth word.',
        '=TRIM(MID(SUBSTITUTE(C10," ",REPT(" ",LEN(C10))), (4-1)*LEN(C10)+1, LEN(C10)))',
        '\'=TRIM(MID(SUBSTITUTE(text," ",REPT(" ",LEN(text))), (N-1)*LEN(text)+1, LEN(text)))'
      ]
    ]

    # Add table caption.
    caption = 'Example String Functions.'
    cell_format = workbook.add_format({'bold': True, 'font_color': '#76933C', 'font_size': 18})
    example.write('B1', caption, cell_format)

    # Set the columns widths.
    example.set_column('A:A', 2.29)
    example.set_column('B:B', 17.43)
    example.set_column('C:C', 48.53)
    example.set_column('D:D', 19.29)
    example.set_column('E:E', 119.00)

    # Insert the example_text table into Excel.
    example.add_table('B3:E15', {'data': example_text,
                                'style': 'Table Style Medium 11',
                                'columns': [{'header': 'Function'},
                                            {'header': 'String'},
                                            {'header': 'Output'},
                                            {'header': 'Formula'},
                                            ]})

    # Close the workbook.
    workbook.close()

# Generate a list of files in the directory.
d = List()

# Generate an excel spreadsheet containing the list of files.
d.generateWorkbook()