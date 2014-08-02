ViewMsg
=======

Python program to open Outlook *.msg files without needing an Outlook install

I wrote this script to open Outlook format files on my Ubuntu Linux computer.  
It utilizes a few other project that were written to read Microsoft OLE2 files, 
and one that extracts the Outlook-specific text in particular. 

--- INSTALLATION ---

Extract the python files into a directory

Install python-tk
ie: sudo apt-get install python-tk

--- RUNNING ---

From the command-line run:
>python ViewMsg.py

Script optionally takes in a command-line parameter of the file name
>python ViewMsg.py filename.msg

--- FUTURES ---

The script and user-interface uses the easygui module, which is basic GUI examples.  
Work will be done to convert this to a more modern-looking text window.
