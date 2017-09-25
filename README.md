# activity_wheel_data_sorter
Mouse activity wheels output each recording session as a single excel file with wheels on separate sheets.
Copying and pasting data from sheets was taking too much time, hindering efficient data analysis.
I wrote this Python script with a basic GUI for sorting excel files output from activity wheels
to make the data analysis process more efficient.

The script requires "input" and "output" file directories in the same directory as the python script.
The excel files you wish to combine are placed together in an "input" file directory.
The excel Files being combined must use the same activity wheels otherwise an error will occur.
Data will be combined into a single .csv file that will be generated in the "output" file directory.

An explaination of how to use the script can also be found in the instruction.txt file.
