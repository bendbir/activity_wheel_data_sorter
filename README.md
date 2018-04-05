# activity_wheel_data_sorter
Benjamin Doyle 4/4/2018

Bioseb mouse activity wheels output each recording session as a single excel file with wheels on separate sheets.
Copying and pasting data from sheets was taking too much time, hindering efficient data analysis.
I wrote this Python script with a basic GUI for sorting excel files output from activity wheels
to make the data analysis process more efficient.

The script requires "input" and "output" file directories in the same directory as the python script.  The titles
of the folders must be "input" and "output", all lowercase.  The excel files you wish to combine are placed together
in an "input" file directory.  Specify which wheels you want data from, then press the combine button.  Data will be 
combined into a single .csv file that will be generated in the "output" file directory.

An explanation of how to use the script can also be found in the instruction.txt file.
