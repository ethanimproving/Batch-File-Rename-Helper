# Batch File Handler
**Version 1.0.0**

A simple python application to rename batch files.

This application was created more for a handy place to store useful cmd commands than using Python to handle it for you. Check out the tutorials to see what this app is doing under the hood.

## Get started with Batch File Handler

Ensure you have [Python](https://www.python.org/downloads/windows/ "Python Releases for Windows") installed, then:

``` bash
python batch_file_rename.py
```

### Entry Point

`batch_file_rename.py` is our entry point. Use `python batch_file_rename.py` to launch the app.

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