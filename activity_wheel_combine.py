import xlrd
import csv
import os
import sys
from tkinter import *


####
# Define GUI function to combine files
def combine_files():
    print ("COMBINED")

    # Extract Wheel Num Info from Input Window
    raw_wheel_input = wheel_nums_entry.get()
    wheel_input = raw_wheel_input.split(',')
    wheel_list = [x.strip() for x in wheel_input]
    print ("Wheel List: ", wheel_list)

    # Extract output file name from input window
    output_csv = output_name_entry.get()

    # Get cwd of python script.
    os.chdir(os.path.dirname(sys.argv[0]))


    # Define input and output directory paths.
    input_dir = os.getcwd() + '/input'
    output_dir = os.getcwd() + '/output'
    # Change to input directory
    os.chdir(input_dir)
    # Check that directory has changed
    print (os.getcwd())

    # List files form input directory
    cwd_files = os.listdir(input_dir)

    # Seperate out xls files, just in case
    xls_files = [x for x in cwd_files if x.endswith('.xls')]

    # Print all files in input directory.
    print ("Xls files in input directory: ", xls_files)

    # Will store information in master_dict using this format:
    # {wkbook:{sheet:[speed, accel, etc],...},...}
    master_dict = {}

    # Cycle through all xls_files
    for excel_file in xls_files:

        # Add blank dictionary to master_dict
        master_dict[excel_file] = {}

        # Open excel file with xlrd.open_workbook
        workbook = xlrd.open_workbook(excel_file)

        # Cycle through wheel_nums to cycle through sheets.
        for wheel_num in wheel_list:
            sheet_name = "Wheel " + wheel_num
            worksheet = workbook.sheet_by_name(sheet_name)
            print (worksheet)



            # Test worksheet.nrows
            print ("worksheet.nrows:", worksheet.nrows)

            # Parse through rows, look for mean row.
            # NOTE, if mouse did not run there won't be a mean row.
            # Setup boolean for no_mean_row, assume there isn't one unless found.
            no_mean_row = True
            for row_idx in range(0, worksheet.nrows):
                test_row = worksheet.row(row_idx)
                if str(test_row[0]) == "text:'Mean:'":
                    print ("found mean row!")
                    no_mean_row = False
                    mean_row = row_idx
                    print ("Mean Row:", mean_row)

                # Also find first row of data recorded.
                # This will help sum for total_distance and total_time_running
                if str(test_row[0]) == "text:'Period Start'":
                    period_start = row_idx
                    print ("Period Start Row:", period_start)

            if no_mean_row == False:
                # column 3 is mean Av. Speed (rev/min)
                # column 4 is mean Max Speed (rev/min)
                # column 6 is mean Av. Acceleration (cm/s-2)
                # column 7 is mean Max Acceleration (cm/s-2)
                # column 8 is mean Distance Traveled (cm)
                num_of_bouts = worksheet.row(mean_row - 2)[1].value
                av_speed = worksheet.row(mean_row)[3].value
                max_speed = worksheet.row(mean_row)[4].value
                av_accel = worksheet.row(mean_row)[6].value
                max_accel = worksheet.row(mean_row)[7].value
                dist_traveled = worksheet.row(mean_row)[8].value
                num_of_bouts = worksheet.row(mean_row - 2)[1].value

                # Range of data is ((period_start + 1), (mean row - 4))
                # Use this to Determine total time Running
                # and total distance traveled.
                total_time_running = 0
                total_distance = 0
                for x in range((period_start + 1), (mean_row - 4)):
                    # NOTE ON BOTTOM OF RANGE USED
                    # 37 for 2.5 hr session, 5min intervals
                    # 25 for 1.5 hr session, 5min intervals
                    # 31 for 2hr - 5mins (shortened session)
                    total_time_running += worksheet.row(x)[1].value
                    total_distance += worksheet.row(x)[8].value

            # If there is NO mean row, no_mean_row still equals True
            if no_mean_row == True:

                # Set all to "NA", since there is no or negligable running activity
                num_of_bouts = "NA"
                av_speed = "NA"
                max_speed = "NA"
                av_accel = "NA"
                max_accel = "NA"
                dist_traveled = "NA"
                num_of_bouts = "NA"

                total_time_running = "NA"
                total_distance = "NA"


            # Print to test Data
            print ("Mean Data From Excel File...")
            print ("Number of Running Bouts:", num_of_bouts)
            print ("Average Speed:", av_speed)
            print ("Max Speed:", max_speed)
            print ("Average Acceleration:", av_accel)
            print ("Max Acceleration:", max_accel)
            print ("Distance Traveled:", dist_traveled)
            print ("Total Running Time:", total_time_running)
            print ("Total Distance Traveled:", total_distance)

            # Add new values to master dictionary for files in input folder
            sheet_list = [num_of_bouts, av_speed, max_speed, av_accel, max_accel, dist_traveled, total_time_running, total_distance]
            master_dict[excel_file][sheet_name] = sheet_list

    # Test master_dict
    print (master_dict)

    #Change Directory to Output
    os.chdir(output_dir)

    # Write master_dict to csv file
    with open(output_csv, "a") as csv_file:
        csv_app = csv.writer(csv_file)
        # Add Labels at top of CSV file
        csv_app.writerow(["File & Wheel", "# of Running Bouts","Average Speed","Max Speed","Average Acceleration","Max Acceleration","Distance Traveled:","Total Running Time","Total Distance Traveled"])

        for xl in xls_files:
            csv_app.writerow([str(xl)])

            for wheel_num in wheel_list:
                # Need all items in one list to be added as a row
                row_list = []
                sheet_name = "Wheel " + wheel_num
                # Add sheet_name
                row_list.append(sheet_name)
                # Add data
                row_list.extend(master_dict[xl][sheet_name])
                # Write to csv file
                csv_app.writerow(row_list)

    print ("Done! Activity wheel excel data has been combined into a single csv file, ", output_csv)
######

######
# Create GUI window
master = Tk(className=" Combine Activity Wheel Data")
Label(master, text="Hello!  Welcome to activity wheel data parser.  \n\
Please make sure you have read the instructions .txt file and all excel files \n\
are in the input file directory.  Output file (.csv) will be generated \n\
in the output directory.").grid(columnspan=2)
Label(master, text="Wheel numbers (seperate with commas): ").grid(row=1, column=0)
Label(master, text="Desired output file name (end with .csv): ").grid(row=2, column=0)

wheel_nums_entry = Entry(master)
output_name_entry = Entry(master)

wheel_nums_entry.grid(row=1, column=1)
output_name_entry.grid(row=2, column=1)

Button(master, text='Quit', command=master.quit).grid(row=3, column=0, sticky=W+E+N+S, pady=4)
Button(master, text='Combine Excel Files', command=combine_files).grid(row=3, column=1, sticky=W+E+N+S, pady=4)

mainloop( )
