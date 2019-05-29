# Batch File Renamer
# Ethan Miller
# 5/29/2019

import subprocess

# Get directory from user.
path=input('Enter directory to write file in: ')

# Create _list.txt file with list of files in directory.
subprocess.call(f'cd {path} && dir /b > _list.txt', shell=True)