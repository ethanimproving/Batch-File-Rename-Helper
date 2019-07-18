# Batch File Rename Helper
**Version 1.0.0**

A simple python application to rename batch files.

This application was created more for a handy place to store useful cmd commands than using Python to handle it for you. Check out the tutorials to see what this app is doing under the hood.

## Download Batch File Handler

The easiest way to use Batch File Rename Helper is to download the exe from the [releases](https://github.com/ethanmiller1/Batch-File-Rename-Helper/releases) page.

Note: you may have to unblock it from Windows SmartScreen prevention:

![](https://github.com/ethanmiller1/Batch-File-Rename-Helper/raw/master/images/unblock.png)

## Get started with Batch File Handler inside Python

Ensure you have [Python](https://www.python.org/downloads/windows/ "Python Releases for Windows") installed, then:

``` bash
# Activate command prompt if using powershell
$ cmd
# Create a virtual environment
$ python -m venv venv
# Activate your virtual environment
$ cd venv/Scripts && activate && cd ../..
# Install dependencies to venv
$ pip install -r requirements.txt
# Run the program
$ python batch_file_rename.py
```

### Entry Point

`batch_file_rename.py` is our entry point. Use `python batch_file_rename.py` to launch the app.

### Turn the project into an EXE

Make sure your Python Scripts folder is a PATH in your system's Environment Variables:

``` bash
Environement Variables > Path > Edit > New ..\Python37\Scripts
```

Install pyinstaller from the command line:

``` bash
pip install pyinstaller
pyinstaller
```

Run a command to compile the py file into an executable, and choose an icon:

``` bash
pyinstaller -F -i "..\images\favicon.ico" batch_file_rename.py
```

## Usage

### Generate Your File List

Simply run the program and copy the directory that contains your files into the command console. The program will create a txt file called "_list" that contains a list of all the files in that directory, copy the contents into an excel spreadsheet, and delete the txt file.

![](https://github.com/king-melchizedek/Batch-File-Rename-Helper/raw/master/demos/GenerateFileList.gif)

### Edit In A Spreadsheet

Add whichever excel formulas you need to accomplish your new naming convention.

![](https://github.com/king-melchizedek/Batch-File-Rename-Helper/raw/master/demos/EditWithExcel.gif)

### Use string functions to parse strings

There are a number of example string functions you can alter to suit your needs in the sheet entitles "string functions."

![](https://github.com/king-melchizedek/Batch-File-Rename-Helper/raw/master/demos/UseExampleFunctions.gif)

### Rename with CMD

The syntax for renaming files with the command prompt are as so:

``` bash
ren [original filename] [new filename]
```

So in our spreadsheet we are going to create this. In column D, our spreadsheet has this function preloaded into it:

``` bash
=CONCATENATE("ren """,A2,""""," ","""",C2,"""")
```

This takes the old name, wraps it in quotes, adds a space, and then wraps the new name in quotes. When we copy the entire column we can simply paste it directly into our command shell.

![](https://github.com/king-melchizedek/Batch-File-Rename-Helper/raw/master/demos/RenameWithCMD.gif)

Shift+Right-Click your directory and select `Open PowerShell window here`, then type `cmd` into the console to access depreciated commands. 

## Rename batch files manually

#### 1. Create a list of the files in your directory

Open the desired directory in a command shell, then:

``` bash
dir /b > _list.txt
```

(`cmd /r dir /b > _list.txt` if using Powershell)

To get DATE and TIME:

``` bash
cmd /r dir /T:W > _list.txt
```

#### 2. Copy the contents of _list.txt to a spreadsheet named _list.xlsx
#### 3. Use an excel formula to rename the old files

Let column A represent the old names, and column C represent the new names. Copy this formula into D1, and copy it down for as many rows as needed.

``` vb
=CONCATENATE("ren """,A1,""""," ","""",C1,"""")
```

## Contributors

---

- Ethan Miller <ethan.romans5.8@gmail.com>

---

## License & copyright

© 2019 Ethan Miller