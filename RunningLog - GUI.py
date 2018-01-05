#! python2
#-------------------------------------------------------------------------------
# Name: Shoe Log
# Purpose:     Update the running log
#
# Author:      Jamie Smart
#
# Created:     14/08/2017
# Copyright:   (c) Jamie 2017
# Notes:       Edit in PyScripter only - not Notepad++
#-------------------------------------------------------------------------------


import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from operator import itemgetter

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Running Log").sheet1

upper_bound = sheet.row_count+1
dateFormat = re.compile('\d{2}/\d{2}/\d{4}') ## CREATE AND COMPUTE GLOBAL VARIABLE TO TEST FOR VALID DATE FORMATS



### DEFINE FUNCTIONS
# GREETING STARTS THE PROGRAM
class RunningLog:
    def __init__(self):
        self.edit_window_exists = 0
        self.main_window = tk.Tk()
        self.main_window.title('Shoe Log')


        self.menubar = tk.Menu(self.main_window)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)

        self.filemenu.add_command(label='Log run', command=self.logrun)
        self.filemenu.add_command(label='View shoe data', command=self.shoe_miles)
        self.filemenu.add_command(label='Calculate miles run', command=self.workouts)
        self.filemenu.add_command(label='View/Edit workout data', command=self.edit_workout_master)
        self.filemenu.add_command(label='Sort log', command=self.sortLog)
        self.filemenu.add_separator()
        self.filemenu.add_command(label='Exit', command=self.main_window.destroy)
        self.menubar.add_cascade(label='Menu', menu=self.filemenu)

        self.main_window.config(menu=self.menubar)

        self.greeting = tk.Label(self.main_window, text="Welcome to Shoe Log! Today's date is %s.\n" %datetime.datetime.now().strftime("%m/%d/%Y"))

        self.greeting.pack()
        self.main_window.mainloop()


    def refresh_screen(self, window, fields):
        window.destroy()
        for field in fields:
            field.delete(0, 'end')

    # Displays a warning if the shoe passed in has more than 250 miles recorded.
    def warning(self, shoe):
        if sheet.row_count < 2:
            pass
        else:
            tmlist = sheet.col_values(4)[1:]
            slist = sheet.col_values(5)[1:]

            shoeWarn = 0.0
            for i in range(0,len(slist)):
                if slist[i] == shoe:
                    shoeWarn += float(tmlist[i])

            if shoeWarn > 250:
                warn = tk.Toplevel()

                if shoeWarn > 500:
                    warning = "WARNING: These shoes (%s) have over 500 miles. It is time to replace them." %(shoe)
                    warning = tk.Label(warn, text=warning)

                elif shoeWarn > 450:
                    warning = "WARNING: These shoes (%s) have over 450 miles. It is time to replace them." %(shoe)
                    warning = tk.Label(warn, text=warning)

                elif shoeWarn > 400:
                    warning = "WARNING: These shoes (%s) have over 400 miles. It is time to replace them." %(shoe)
                    warning = tk.Label(warn, text=warning)

                elif shoeWarn > 350:
                    warning = "WARNING: These shoes (%s) have over 350 miles. It is time to replace them." %(shoe)
                    warning = tk.Label(warn, text=warning)

                elif shoeWarn > 300:
                    warning = "WARNING: These shoes (%s) have over 300 miles. It is time to replace them." %(shoe)
                    warning = tk.Label(warn, text=warning)

                elif shoeWarn > 250:
                    warning = "WARNING: These shoes (%s) have over 250 miles. It is time to replace them." %(shoe)
                    warning = tk.Label(warn, text=warning)

                warning.pack()
                warn_ok = tk.Button(warn, text='OK', command=self.warn.destroy)
                warn_ok.pack()
                warn.mainloop()

    # LOGS A RUN
    def logrun(self):
        def validate_save_run():
            run_date = tk.StringVar()  # create variable to hold run date
            run_date.set(self.run_date_box.get())
            test = re.match(dateFormat, self.run_date_box.get())  # see if date is formatted correctly
            while test is None: ## PROMPT USER TO ENTER A CORRECTLY FORMATTED DATE IF FORMAT IS INCORRECT
                if self.error == 0:  # error is not showing
                    date_error_label = tk.Label(self.warning_frame, text='The date you entered is not in the correct format. Please format your date correctly (mm/dd/yyyy)')
                    date_error_label.pack()
                    self.error = 1
                run_date.set(self.run_date_box.get)
                test = re.match(dateFormat, self.run_date_box.get())
                self.ok_button.wait_variable(run_date)



            # REPLACE MISSING ENTRIES FOR MILES RUN/WALKED WITH 0
            if self.run_box.get() == '':
                run_miles = 0
            else:
                run_miles = float(self.run_box.get())

            if self.walk_box.get() == '':
                walk_miles = 0
            else:
                walk_miles = float(self.walk_box.get())

            total_miles = str(float(walk_miles) + float(run_miles))

            # REPLACE OTHER BLANKS WITH QUESTION MARKS
            if self.shoe_box.get() == '':
                shoe = '?'
            else:
                shoe = self.shoe_box.get().title()

            if self.type_box.get() == '':
                run_type = '?'
            else:
                run_type = self.type_box.get().title()

            record = run_date.get(),walk_miles,run_miles,total_miles,shoe,run_type

            # Create header row if one doesn't exist, and add run to spreadsheet
            if sheet.cell(1,1).value == "":
                header = "Date","Miles walked","Miles run","Total miles","Shoes","Type of run"
                sheet.insert_row(header,1)
                sheet.insert_row(record,2)
            else:
                sheet.insert_row(record,sheet.row_count+1)

            self.warning(shoe) ## DISPLAYS A WARNING IF THE SHOE HAS >250 MILES
            boxes = [self.run_date_box, self.shoe_box, self.walk_box, self.run_box, self.type_box]


            self.confirmation = tk.Toplevel()
            self.record_confirmation = tk.Label(self.confirmation, text='Your run has been recorded.')
            self.record_confirmation.pack()
            self.confirm_ok_button = tk.Button(self.confirmation, text="OK", command=lambda: self.refresh_screen(self.confirmation, boxes))
            self.confirm_ok_button.pack()
            self.confirmation.mainloop()

        self.log_run_screen = tk.Toplevel()

        self.frame1 = tk.Frame(self.log_run_screen)
        self.frame2 = tk.Frame(self.log_run_screen)
        self.frame3 = tk.Frame(self.log_run_screen)
        self.frame4 = tk.Frame(self.log_run_screen)
        self.frame5 = tk.Frame(self.log_run_screen)
        self.frame6 = tk.Frame(self.log_run_screen)
        self.warning_frame = tk.Frame(self.log_run_screen)

        self.run_date_label = tk.Label(self.frame1, text='Date of your run (mm/dd/yyyy)')
        self.shoe_label = tk.Label(self.frame2, text='What shoes did you wear')
        self.walk_label = tk.Label(self.frame3, text='How many miles did you walk in these shoes')
        self.run_label = tk.Label(self.frame4, text='How many miles did you run in these shoes')
        self.type_label = tk.Label(self.frame5, text='What type of run was this')

        self.run_date_box = tk.Entry(self.frame1, width=50)
        self.shoe_box = tk.Entry(self.frame2, width=50)
        self.walk_box = tk.Entry(self.frame3, width=50)
        self.run_box = tk.Entry(self.frame4, width=50)
        self.type_box = tk.Entry(self.frame5, width=50)

        self.ok_button = tk.Button(self.frame6, text='OK', command=validate_save_run)
        self.cancel_button = tk.Button(self.frame6, text='Exit', command=self.log_run_screen.destroy)


        self.run_date_label.pack(side='left')
        self.run_date_box.pack(side='left')
        self.shoe_label.pack(side='left')
        self.shoe_box.pack(side='left')
        self.walk_label.pack(side='left')
        self.walk_box.pack(side='left')
        self.run_label.pack(side='left')
        self.run_box.pack(side='left')
        self.type_label.pack(side='left')
        self.type_box.pack(side='left')
        self.ok_button.pack(side='left')
        self.cancel_button.pack(side='left')
        self.frame1.pack()
        self.warning_frame.pack()
        self.frame2.pack()
        self.frame3.pack()
        self.frame4.pack()
        self.frame5.pack()
        self.frame6.pack()

        self.error = 0

        self.log_run_screen.mainloop()


    # CALCULATES NUMBER OF MILES ON SHOES
    def shoe_miles(self):
        def calculate_miles():
            shoe_warn = tk.Label(self.shoe_warning_frame, text='Please select a pair of shoes.')
            while self.shoe.get() == '':
                if self.shoe_error == 0:
                    shoe_warn.pack()
                    self.shoe_error = 1
                self.ok_button.wait_variable(self.shoe)

            # Create variables for the number of miles walked, run, and total, and add to them if the shoe in the row matches the shoe entered
            if self.shoe_error == 1:
                shoe_warn.destroy()
            walkReturn = 0.0
            runReturn = 0.0
            totalReturn = 0.0
            for i in range(0, len(slist)):
                if slist[i] == self.shoe.get():
                    walkReturn += float(wmlist[i])
                    runReturn += float(rmlist[i])
                    totalReturn += float(tmlist[i])
            shoe_data = 'Miles walked:  %.2f\nMiles run:  %.2f\nTotal Miles:  %.2f' %(walkReturn,runReturn,totalReturn)
            message = tk.Label(self.middle_frame, text=shoe_data)
            if self.show_count == 0:
                message.pack()
                self.show_count += 1
            self.ok_button.wait_variable(self.shoe)
            self.show_count = 0
            message.destroy()

        self.shoe_miles_window = tk.Toplevel()
        if sheet.row_count < 2: ## CHECKS TO SEE THAT THERE IS DATA IN THE SPREADSHEET
            self.shoe_warning = tk.Label(self.shoe_miles_window, text='You have not entered any data.')
            self.shoe_warn_confirm = tk.Button(self.shoe_miles_window, text='OK', command=self.shoe_miles_window.destroy)
            self.shoe_warning.pack()
            self.shoe_warn_confirm.pack()
        else: ## CREATE LISTS OF THE DIFFERENT DATA COLUMNS
            #rdlist = sheet.col_values(1)[1:]
            wmlist = sheet.col_values(2)[1:]
            rmlist = sheet.col_values(3)[1:]
            tmlist = sheet.col_values(4)[1:]
            slist = sheet.col_values(5)[1:]
            #rtlist = sheet.col_values(6)[1:]


            # Create and display a list of shoes in the spreadsheet
            list_of_shoes = list(set(slist))
            if '?' in list_of_shoes:
                list_of_shoes.remove('?')
            list_of_shoes.sort()

            self.shoe_error = 0

            self.shoe = tk.StringVar(self.shoe_miles_window)
            self.shoe.set('')
            self.top_frame = tk.Frame(self.shoe_miles_window)
            self.shoe_warning_frame = tk.Frame(self.shoe_miles_window)
            self.middle_frame = tk.Frame(self.shoe_miles_window)
            self.bottom_frame = tk.Frame(self.shoe_miles_window)
            self.shoe_lookup = tk.Label(self.top_frame, text='Which shoes would you like to view?')
            self.shoe_lookup_box = tk.OptionMenu(self.top_frame, self.shoe, *list(list_of_shoes))
            self.ok_button = tk.Button(self.bottom_frame, text='OK', command=calculate_miles)
            self.exit_button = tk.Button(self.bottom_frame, text='Exit', command=self.shoe_miles_window.destroy)

            self.shoe_lookup.pack(side='left')
            self.shoe_lookup_box.pack(side='left')
            self.ok_button.pack(side='left')
            self.exit_button.pack(side='left')
            self.top_frame.pack()
            self.shoe_warning_frame.pack()
            self.middle_frame.pack()
            self.bottom_frame.pack()

            self.show_count = 0

            self.shoe_miles_window.mainloop()

    def workouts(self):
        def validate_and_save():
            dates = tk.StringVar()
            dates.set(self.begin_box.get() + self.end_box.get())
            while re.match(dateFormat,self.begin_box.get()) is None or re.match(dateFormat,self.end_box.get()) is None:
                self.warning = tk.Label(self.warning_frame, text='Please format your dates correctly (mm/dd/yyyy).')
                if self.error == 0:
                    self.warning.pack()
                    self.error = 1
                self.ok_button.wait_variable(dates)

            self.warning_frame.destroy()
            if self.show_count > 0:
                self.run_data.destroy()
            ## CONVERT DATES FROM STRING TO DATE FORMAT SO DATES CAN BE COMPARED
            periodStart = self.begin_box.get().split('/')
            periodEnd = self.end_box.get().split('/')
            start = datetime.date(int(periodStart[2]),int(periodStart[0]),int(periodStart[1]))
            end = datetime.date(int(periodEnd[2]),int(periodEnd[0]),int(periodEnd[1]))

            runReturn = 0.0

            for i in range(2, sheet.row_count+1): ## COMPARE ROW DATE TO RANGE ENTERED AND ADD MILES RUN ON MATCHING DATES TO RUNNING TOTAL
                rowDate = sheet.cell(i,1).value.split('/')
                rowDate = datetime.date(int(rowDate[2]),int(rowDate[0]),int(rowDate[1]))
                if (rowDate <= end and rowDate >= start):
                    runReturn += float(sheet.cell(i,3).value)


            sentence = "You ran %.2f miles between %s and %s." %(runReturn, self.begin_box.get(), self.end_box.get())
            self.run_data = tk.Label(self.results, text=sentence)

            self.run_data.pack()
            self.show_count += 1

        self.workout_window = tk.Toplevel()

        self.box_frame = tk.Frame(self.workout_window)
        self.warning_frame = tk.Frame(self.workout_window)
        self.button_frame = tk.Frame(self.workout_window)

        self.date_label = tk.Label(self.workout_window, text='Please enter the date range that you would like to view (mm/dd/yyyy).')
        self.begin_box = tk.Entry(self.box_frame)
        self.dash = tk.Label(self.box_frame, text='-')
        self.end_box = tk.Entry(self.box_frame)
        self.results = tk.Frame(self.workout_window)
        self.ok_button = tk.Button(self.button_frame, text='OK', command=validate_and_save)
        self.exit_button = tk.Button(self.button_frame, text='Exit', command=self.workout_window.destroy)

        self.error = 0

        if sheet.row_count < 2:
            self.warning = tk.Label(self.warning_frame, text='You do not have any data entered.')
            self.no_data_ok = tk.Button(self.workout_window, text='OK', command=self.workout_window.destroy)
            self.warning.pack()
            self.warning_frame.pack()
            self.no_data_ok.pack()
        else:
            self.date_label.pack()
            self.begin_box.pack(side='left')
            self.dash.pack(side='left')
            self.end_box.pack(side='left')
            self.ok_button.pack(side='left')
            self.exit_button.pack(side='left')
            self.box_frame.pack()
            self.warning_frame.pack()
            self.results.pack()
            self.button_frame.pack()

        self.show_count = 0
        self.workout_window.mainloop()


    # View or edit data
    def edit_workout_master(self, upperBound = upper_bound):
        def update_row(row_num, lower_bound, delete):
            row = lower_bound + row_num
            record = str(row_num+1)

            if record == '1':
                nonlocal date1, walk1, run1, total1, shoes1, type1, delete_box_1, update1
                date_entered = date1.get("1.0","end")
            if record == '2':
                nonlocal date2, walk2, run2, total2, shoes2, type2, delete_box_2, update2
                date_entered = date2.get("1.0","end")
            if record == '3':
                nonlocal date3, walk3, run3, total3, shoes3, type3, delete_box_3, update3
                date_entered = date3.get("1.0","end")
            if record == '4':
                nonlocal date4, walk4, run4, total4, shoes4, type4, delete_box_4, update4
                date_entered = date4.get("1.0","end")
            if record == '5':
                nonlocal date5, walk5, run5, total5, shoes5, type5, delete_box_5, update5
                date_entered = date5.get("1.0","end")
            if record == '6':
                nonlocal date6, walk6, run6, total6, shoes6, type6, delete_box_6, update6
                date_entered = date6.get("1.0","end")
            if record == '7':
                nonlocal date7, walk7, run7, total7, shoes7, type7, delete_box_7, update7
                date_entered = date7.get("1.0","end")
            if record == '8':
                nonlocal date8, walk8, run8, total8, shoes8, type8, delete_box_8, update8
                date_entered = date8.get("1.0","end")
            if record == '9':
                nonlocal date9, walk9, run9, total9, shoes9, type9, delete_box_9, update9
                date_entered = date9.get("1.0","end")
            if record == '10':
                nonlocal date10, walk10, run10, total10, shoes10, type10, delete_box_10, update10
                date_entered = date10.get("1.0","end")


            if delete.get() > 0 and delete.get() == int(record):
                exec('date'+record+'.grid_forget()')
                exec('walk'+record+'.grid_forget()')
                exec('run'+record+'.grid_forget()')
                exec('total'+record+'.grid_forget()')
                exec('shoes'+record+'.grid_forget()')
                exec('type'+record+'.grid_forget()')
                exec('delete_box_'+record+'.grid_forget()')
                exec('update'+record+'.grid_forget()')
                sheet.delete_row(row)
                self.edit_window.update()
            else:
                run_date = tk.StringVar()  # create variable to hold run date
                exec('run_date.set(date'+record+'.get("1.0","end"))')
                test = re.match(dateFormat, date_entered)  # see if date is formatted correctly
                exec('date'+record+'.config(foreground="black")')
                while test is None: ## PROMPT USER TO ENTER A CORRECTLY FORMATTED DATE IF FORMAT IS INCORRECT
                    exec('date'+record+'.config(foreground="red")')
                    exec('run_date.set(date'+record+'.get("1.0","end"))')
                    exec('test = re.match(dateFormat, date'+record+'.get("1.0","end"))')
                    exec('update'+record+'.wait_variable(run_date)')


                exec('sheet.update_cell(row, 1, date'+record+'.get("1.0","end"))')
                exec('sheet.update_cell(row, 2, walk'+record+'.get("1.0","end"))')
                exec('sheet.update_cell(row, 3, run'+record+'.get("1.0","end"))')
                exec('sheet.update_cell(row, 4, total'+record+'.get("1.0","end"))')
                exec('sheet.update_cell(row, 5, shoes'+record+'.get("1.0","end"))')
                exec('sheet.update_cell(row, 6, type'+record+'.get("1.0","end"))')

        def prev_runs():
            if lowerBound == 2:
                self.edit_workout_master(upperBound)
            else:
                self.edit_workout_master(lowerBound)

        def next_runs(upperBound = upperBound):
            upperBound += 10
            if upperBound + 10 > sheet.row_count:
                upperBound = sheet.row_count
            else:
                pass
            self.edit_workout_master(upperBound+1)

        if not self.edit_window_exists:
            self.edit_window = tk.Toplevel()
        lowerBound = upperBound-10

        if lowerBound <= 1:
            lowerBound = 2

        rowNum = 1
        rowList = []
        for i in range(lowerBound,upperBound):
            rowContents = sheet.row_values(i)[0:6]
            rowContents.insert(0, rowNum)
            rowList.append(rowContents)
            rowNum += 1


        ## CREATE AND DISPLAY TABLE OF X RUNS
        date_header = tk.Label(self.edit_window, text='Date').grid(row=0, column=0)
        mi_walked_header = tk.Label(self.edit_window, text='Mi. walked').grid(row=0, column=1)
        mi_run_header = tk.Label(self.edit_window, text='Mi. run').grid(row=0, column=2)
        total_mi_header = tk.Label(self.edit_window, text='Total mi.').grid(row=0, column=3)
        shoes_header = tk.Label(self.edit_window, text='Shoes').grid(row=0, column=4)
        type_header = tk.Label(self.edit_window, text='Run type').grid(row=0, column=5)
        delete_header = tk.Label(self.edit_window, text='Delete?').grid(row=0, column=6)

        date1 = tk.Text(self.edit_window, width=10, height=1)
        date2 = tk.Text(self.edit_window, width=10, height=1)
        date3 = tk.Text(self.edit_window, width=10, height=1)
        date4 = tk.Text(self.edit_window, width=10, height=1)
        date5 = tk.Text(self.edit_window, width=10, height=1)
        date6 = tk.Text(self.edit_window, width=10, height=1)
        date7 = tk.Text(self.edit_window, width=10, height=1)
        date8 = tk.Text(self.edit_window, width=10, height=1)
        date9 = tk.Text(self.edit_window, width=10, height=1)
        date10 = tk.Text(self.edit_window, width=10, height=1)

        walk1 = tk.Text(self.edit_window, width=5, height=1)
        walk2 = tk.Text(self.edit_window, width=5, height=1)
        walk3 = tk.Text(self.edit_window, width=5, height=1)
        walk4 = tk.Text(self.edit_window, width=5, height=1)
        walk5 = tk.Text(self.edit_window, width=5, height=1)
        walk6 = tk.Text(self.edit_window, width=5, height=1)
        walk7 = tk.Text(self.edit_window, width=5, height=1)
        walk8 = tk.Text(self.edit_window, width=5, height=1)
        walk9 = tk.Text(self.edit_window, width=5, height=1)
        walk10 = tk.Text(self.edit_window, width=5, height=1)

        run1 = tk.Text(self.edit_window, width=5, height=1)
        run2 = tk.Text(self.edit_window, width=5, height=1)
        run3 = tk.Text(self.edit_window, width=5, height=1)
        run4 = tk.Text(self.edit_window, width=5, height=1)
        run5 = tk.Text(self.edit_window, width=5, height=1)
        run6 = tk.Text(self.edit_window, width=5, height=1)
        run7 = tk.Text(self.edit_window, width=5, height=1)
        run8 = tk.Text(self.edit_window, width=5, height=1)
        run9 = tk.Text(self.edit_window, width=5, height=1)
        run10 = tk.Text(self.edit_window, width=5, height=1)

        total1 = tk.Text(self.edit_window, width=5, height=1)
        total2 = tk.Text(self.edit_window, width=5, height=1)
        total3 = tk.Text(self.edit_window, width=5, height=1)
        total4 = tk.Text(self.edit_window, width=5, height=1)
        total5 = tk.Text(self.edit_window, width=5, height=1)
        total6 = tk.Text(self.edit_window, width=5, height=1)
        total7 = tk.Text(self.edit_window, width=5, height=1)
        total8 = tk.Text(self.edit_window, width=5, height=1)
        total9 = tk.Text(self.edit_window, width=5, height=1)
        total10 = tk.Text(self.edit_window, width=5, height=1)

        shoes1 = tk.Text(self.edit_window, width=20, height=1)
        shoes2 = tk.Text(self.edit_window, width=20, height=1)
        shoes3 = tk.Text(self.edit_window, width=20, height=1)
        shoes4 = tk.Text(self.edit_window, width=20, height=1)
        shoes5 = tk.Text(self.edit_window, width=20, height=1)
        shoes6 = tk.Text(self.edit_window, width=20, height=1)
        shoes7 = tk.Text(self.edit_window, width=20, height=1)
        shoes8 = tk.Text(self.edit_window, width=20, height=1)
        shoes9 = tk.Text(self.edit_window, width=20, height=1)
        shoes10 = tk.Text(self.edit_window, width=20, height=1)

        type1 = tk.Text(self.edit_window, width=10, height=1)
        type2 = tk.Text(self.edit_window, width=10, height=1)
        type3 = tk.Text(self.edit_window, width=10, height=1)
        type4 = tk.Text(self.edit_window, width=10, height=1)
        type5 = tk.Text(self.edit_window, width=10, height=1)
        type6 = tk.Text(self.edit_window, width=10, height=1)
        type7 = tk.Text(self.edit_window, width=10, height=1)
        type8 = tk.Text(self.edit_window, width=10, height=1)
        type9 = tk.Text(self.edit_window, width=10, height=1)
        type10 = tk.Text(self.edit_window, width=10, height=1)

        delete = tk.IntVar()

        delete_box_1 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=1)
        delete_box_2 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=2)
        delete_box_3 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=3)
        delete_box_4 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=4)
        delete_box_5 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=5)
        delete_box_6 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=6)
        delete_box_7 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=7)
        delete_box_8 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=8)
        delete_box_9 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=9)
        delete_box_10 = tk.Checkbutton(self.edit_window, variable=delete, offvalue=0, onvalue=10)

        date1.grid(row=1, column=0)
        date2.grid(row=2, column=0)
        date3.grid(row=3, column=0)
        date4.grid(row=4, column=0)
        date5.grid(row=5, column=0)
        date6.grid(row=6, column=0)
        date7.grid(row=7, column=0)
        date8.grid(row=8, column=0)
        date9.grid(row=9, column=0)
        date10.grid(row=10, column=0)

        walk1.grid(padx=5, row=1, column=1)
        walk2.grid(row=2, column=1)
        walk3.grid(row=3, column=1)
        walk4.grid(row=4, column=1)
        walk5.grid(row=5, column=1)
        walk6.grid(row=6, column=1)
        walk7.grid(row=7, column=1)
        walk8.grid(row=8, column=1)
        walk9.grid(row=9, column=1)
        walk10.grid(row=10, column=1)

        run1.grid(row=1, column=2)
        run2.grid(row=2, column=2)
        run3.grid(row=3, column=2)
        run4.grid(row=4, column=2)
        run5.grid(row=5, column=2)
        run6.grid(row=6, column=2)
        run7.grid(row=7, column=2)
        run8.grid(row=8, column=2)
        run9.grid(row=9, column=2)
        run10.grid(row=10, column=2)

        total1.grid(row=1, column=3)
        total2.grid(row=2, column=3)
        total3.grid(row=3, column=3)
        total4.grid(row=4, column=3)
        total5.grid(row=5, column=3)
        total6.grid(row=6, column=3)
        total7.grid(row=7, column=3)
        total8.grid(row=8, column=3)
        total9.grid(row=9, column=3)
        total10.grid(row=10, column=3)

        shoes1.grid(row=1, column=4)
        shoes2.grid(row=2, column=4)
        shoes3.grid(row=3, column=4)
        shoes4.grid(row=4, column=4)
        shoes5.grid(row=5, column=4)
        shoes6.grid(row=6, column=4)
        shoes7.grid(row=7, column=4)
        shoes8.grid(row=8, column=4)
        shoes9.grid(row=9, column=4)
        shoes10.grid(row=10, column=4)

        type1.grid(row=1, column=5)
        type2.grid(row=2, column=5)
        type3.grid(row=3, column=5)
        type4.grid(row=4, column=5)
        type5.grid(row=5, column=5)
        type6.grid(row=6, column=5)
        type7.grid(row=7, column=5)
        type8.grid(row=8, column=5)
        type9.grid(row=9, column=5)
        type10.grid(row=10, column=5)

        delete_box_1.grid(row=1, column=6)
        delete_box_2.grid(row=2, column=6)
        delete_box_3.grid(row=3, column=6)
        delete_box_4.grid(row=4, column=6)
        delete_box_5.grid(row=5, column=6)
        delete_box_6.grid(row=6, column=6)
        delete_box_7.grid(row=7, column=6)
        delete_box_8.grid(row=8, column=6)
        delete_box_9.grid(row=9, column=6)
        delete_box_10.grid(row=10, column=6)

        date1.insert('1.0', chars=rowList[0][1])
        date2.insert('1.0', chars=rowList[1][1])
        date3.insert('1.0', chars=rowList[2][1])
        date4.insert('1.0', chars=rowList[3][1])
        date5.insert('1.0', chars=rowList[4][1])
        date6.insert('1.0', chars=rowList[5][1])
        date7.insert('1.0', chars=rowList[6][1])
        date8.insert('1.0', chars=rowList[7][1])
        date9.insert('1.0', chars=rowList[8][1])
        date10.insert('1.0', chars=rowList[9][1])

        walk1.insert('1.0', chars=rowList[0][2])
        walk2.insert('1.0', chars=rowList[1][2])
        walk3.insert('1.0', chars=rowList[2][2])
        walk4.insert('1.0', chars=rowList[3][2])
        walk5.insert('1.0', chars=rowList[4][2])
        walk6.insert('1.0', chars=rowList[5][2])
        walk7.insert('1.0', chars=rowList[6][2])
        walk8.insert('1.0', chars=rowList[7][2])
        walk9.insert('1.0', chars=rowList[8][2])
        walk10.insert('1.0', chars=rowList[9][2])

        run1.insert('1.0', chars=rowList[0][3])
        run2.insert('1.0', chars=rowList[1][3])
        run3.insert('1.0', chars=rowList[2][3])
        run4.insert('1.0', chars=rowList[3][3])
        run5.insert('1.0', chars=rowList[4][3])
        run6.insert('1.0', chars=rowList[5][3])
        run7.insert('1.0', chars=rowList[6][3])
        run8.insert('1.0', chars=rowList[7][3])
        run9.insert('1.0', chars=rowList[8][3])
        run10.insert('1.0', chars=rowList[9][3])

        total1.insert('1.0', chars=rowList[0][4])
        total2.insert('1.0', chars=rowList[1][4])
        total3.insert('1.0', chars=rowList[2][4])
        total4.insert('1.0', chars=rowList[3][4])
        total5.insert('1.0', chars=rowList[4][4])
        total6.insert('1.0', chars=rowList[5][4])
        total7.insert('1.0', chars=rowList[6][4])
        total8.insert('1.0', chars=rowList[7][4])
        total9.insert('1.0', chars=rowList[8][4])
        total10.insert('1.0', chars=rowList[9][4])

        shoes1.insert('1.0', chars=rowList[0][5])
        shoes2.insert('1.0', chars=rowList[1][5])
        shoes3.insert('1.0', chars=rowList[2][5])
        shoes4.insert('1.0', chars=rowList[3][5])
        shoes5.insert('1.0', chars=rowList[4][5])
        shoes6.insert('1.0', chars=rowList[5][5])
        shoes7.insert('1.0', chars=rowList[6][5])
        shoes8.insert('1.0', chars=rowList[7][5])
        shoes9.insert('1.0', chars=rowList[8][5])
        shoes10.insert('1.0', chars=rowList[9][5])

        type1.insert('1.0', chars=rowList[0][6])
        type2.insert('1.0', chars=rowList[1][6])
        type3.insert('1.0', chars=rowList[2][6])
        type4.insert('1.0', chars=rowList[3][6])
        type5.insert('1.0', chars=rowList[4][6])
        type6.insert('1.0', chars=rowList[5][6])
        type7.insert('1.0', chars=rowList[6][6])
        type8.insert('1.0', chars=rowList[7][6])
        type9.insert('1.0', chars=rowList[8][6])
        type10.insert('1.0', chars=rowList[9][6])

        update1 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(0, lowerBound, delete))
        update2 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(1, lowerBound, delete))
        update3 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(2, lowerBound, delete))
        update4 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(3, lowerBound, delete))
        update5 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(4, lowerBound, delete))
        update6 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(5, lowerBound, delete))
        update7 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(6, lowerBound, delete))
        update8 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(7, lowerBound, delete))
        update9 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(8, lowerBound, delete))
        update10 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(9, lowerBound, delete))

        update1.grid(row=1, column=7, padx=5)
        update2.grid(row=2, column=7)
        update3.grid(row=3, column=7)
        update4.grid(row=4, column=7)
        update5.grid(row=5, column=7)
        update6.grid(row=6, column=7)
        update7.grid(row=7, column=7)
        update8.grid(row=8, column=7)
        update9.grid(row=9, column=7)
        update10.grid(row=10, column=7)

        prev_button = tk.Button(self.edit_window, text='<<', command=prev_runs).grid(row=11, column=0, pady=20)
        next_button = tk.Button(self.edit_window, text='>>', command=next_runs).grid(row=11, column=7)
        exit_button = tk.Button(self.edit_window, text='Exit', command=self.edit_window.destroy).grid(row=11, column=4)


        if not self.edit_window.winfo_ismapped():
            self.edit_window_exists = 1
            self.edit_window.mainloop()
        else:
            self.edit_window.update()



    # FUNCTION TO SORT WORKBOOK - CURRENTLY RUNS VERY SLOWLY
    def sortLog(self):
        all_values = sheet.get_all_values() # GET A LIST OF ALL OF THE ROWS IN THE WORKSHEET, EACH ROW IS A LIST
        all_values = all_values[1:] # REMOVE HEADER ROW FROM LIST OF VALUES
        all_values_sorted = sorted(all_values, key=itemgetter(0)) # SORT THE LIST BY DATE

        r=2 # INITIALIZE VAR FOR ROW TO START UPDATING FROM

        for line in all_values_sorted: # LOOP THROUGH ROWS
            c=1 # INITIALIZE VAR FOR COL TO START UPDATING FROM - REFRESH FOR EACH ITERATION
            for cell in line: # LOOP THROUGH CELLS IN ROW
                sheet.update_cell(r,c,cell) # UPDATE CELL WITH SORTED DATA
                c += 1 # INCREASE COL COUNTER
            r += 1 # INCREASE ROW COUNTER

        print('Sorting complete.')


################################


RunningLog()

