#! python3
#-------------------------------------------------------------------------------
# Name: Shoe Log
# Purpose:     Update the running log
#
# Author:      Jamie Smart
#
# Created:     14/08/2017
# Copyright:   (c) Jamie 2017
# Notes:       Edit in PyCharm only
#-------------------------------------------------------------------------------


import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
import tkinter.messagebox
from tkinter import ttk
from operator import itemgetter

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Running Log").sheet1

dateFormat = re.compile('\d{2}/\d{2}/\d{4}')  # CREATE AND COMPUTE GLOBAL REGEX PATTERN TO TEST FOR VALID DATE FORMATS


class RunningLog:
    def __init__(self):
        self.upper_bound = sheet.row_count+1
        self.edit_window_exists = 0
        self.main_window = tk.Tk()
        self.main_window.title('Shoe Log')
        self.main_window.minsize(width=300, height=300)

        # CREATE MENU
        self.menubar = tk.Menu(self.main_window)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)

        self.filemenu.add_command(label='Log run', command=self.logrun)
        self.filemenu.add_command(label='View shoe data', command=self.shoe_miles)
        self.filemenu.add_command(label='Calculate miles run', command=self.workouts)
        self.filemenu.add_command(label='View/Edit workout data', command=lambda *args: self.edit_workout_master(self.upper_bound))
        self.filemenu.add_command(label='Sort log', command=self.sort_log)
        self.filemenu.add_separator()
        self.filemenu.add_command(label='Exit', command=self.main_window.destroy)
        self.menubar.add_cascade(label='Menu', menu=self.filemenu)

        self.main_window.config(menu=self.menubar)

        self.greeting = tk.Label(self.main_window, text="Welcome to Shoe Log! Today's date is %s.\n" %datetime.datetime.now().strftime("%m/%d/%Y"))

        self.greeting.grid(padx=250, pady=250)
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
            for i in range(0, len(slist)):
                if slist[i] == shoe:
                    shoeWarn += float(tmlist[i])

            if shoeWarn > 250:
                if shoeWarn > 500:
                    warning = "WARNING: These shoes (%s) have over 500 miles. It is time to replace them." %(shoe.rstrip())

                elif shoeWarn > 450:
                    warning = "WARNING: These shoes (%s) have over 450 miles. It is time to replace them." %(shoe.rstrip())

                elif shoeWarn > 400:
                    warning = "WARNING: These shoes (%s) have over 400 miles. It is time to replace them." %(shoe.rstrip())

                elif shoeWarn > 350:
                    warning = "WARNING: These shoes (%s) have over 350 miles. It is time to replace them." %(shoe.rstrip())

                elif shoeWarn > 300:
                    warning = "WARNING: These shoes (%s) have over 300 miles. It is time to replace them." %(shoe.rstrip())

                else:
                    warning = "WARNING: These shoes (%s) have over 250 miles. It is time to replace them." %(shoe.rstrip())

                tk.messagebox.showwarning("Warning", warning)

    def shoe_list(self):
        shoe_list = list(set(sheet.col_values(5)[1:]))
        for i in range(len(shoe_list)):
            shoe_list[i] = shoe_list[i].rstrip()
        while shoe_list.count('?') > 0:
            shoe_list.remove('?')
        shoe_list.sort()
        return(tuple(shoe_list))

    # LOGS A RUN
    def logrun(self):
        def validate_save_run():
            run_date = tk.StringVar()  # create variable to hold run date
            run_date.set(self.run_date_box.get())
            test = re.match(dateFormat, self.run_date_box.get())  # see if date is formatted correctly
            while test is None:  # PROMPT USER TO ENTER A CORRECTLY FORMATTED DATE IF FORMAT IS INCORRECT
                if self.error == 0:  # error is not showing
                    self.date_error_label = tk.Label(self.log_run_screen, text='The date you entered is not in the correct format. Please format your date correctly (mm/dd/yyyy)', fg="red")
                    self.date_error_label.grid(row=3, column=1, columnspan=2)
                    self.error = 1
                run_date.set(self.run_date_box.get)
                test = re.match(dateFormat, self.run_date_box.get())
                self.ok_button.wait_variable(run_date)

            if self.error == 1:
                self.date_error_label.destroy()

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
            if sheet.cell(1, 1).value == "":
                header = "Date","Miles walked","Miles run","Total miles","Shoes","Type of run"
                sheet.insert_row(header,1)
                sheet.insert_row(record,2)
            else:
                sheet.insert_row(record, sheet.row_count+1)

            self.warning(shoe) ## DISPLAYS A WARNING IF THE SHOE HAS >250 MILES
            boxes = [self.run_date_box, self.shoe_box, self.walk_box, self.run_box, self.type_box]

            self.confirmation = tk.Toplevel()
            self.record_confirmation = tk.Label(self.confirmation, text='Your run has been recorded.')
            self.record_confirmation.pack()
            self.confirm_ok_button = tk.Button(self.confirmation, text="OK", command=lambda: self.refresh_screen(self.confirmation, boxes))
            self.confirm_ok_button.pack()

            self.upper_bound += 1

            self.confirmation.mainloop()

        self.log_run_screen = tk.Toplevel()

        self.spacer = tk.Label(self.log_run_screen, text=" ").grid(row=1, column=1)
        self.run_date_label = tk.Label(self.log_run_screen, text='Date of your run (mm/dd/yyyy):').grid(row=2, column=1, sticky="E")
        self.shoe_label = tk.Label(self.log_run_screen, text='What shoes did you wear?').grid(row=4, column=1, sticky="E")
        self.walk_label = tk.Label(self.log_run_screen, text='How many miles did you walk in these shoes?').grid(row=5, column=1, sticky="E")
        self.run_label = tk.Label(self.log_run_screen, text='How many miles did you run in these shoes?').grid(row=6, column=1, sticky="E")
        self.type_label = tk.Label(self.log_run_screen, text='What type of run was this?').grid(row=7, column=1, sticky="E")

        self.run_date_box = tk.Entry(self.log_run_screen, width=50)
        self.shoe_box = ttk.Combobox(self.log_run_screen, values=self.shoe_list())
        self.walk_box = tk.Entry(self.log_run_screen, width=50)
        self.run_box = tk.Entry(self.log_run_screen, width=50)
        self.type_box = tk.Entry(self.log_run_screen, width=50)

        self.run_date_box.grid(row=2, column=2, padx=10, pady=2)
        self.shoe_box.grid(row=4, column=2, padx=10, pady=2, sticky="W")
        self.walk_box.grid(row=5, column=2, padx=10, pady=2)
        self.run_box.grid(row=6, column=2, padx=10, pady=2)
        self.type_box.grid(row=7, column=2, padx=10, pady=2)

        self.ok_button = tk.Button(self.log_run_screen, text='OK', command=validate_save_run)
        self.ok_button.grid(row=8, column=2, sticky="W", padx=5, pady=20)
        self.cancel_button = tk.Button(self.log_run_screen, text='Exit', command=self.log_run_screen.destroy).grid(row=8, column=2, sticky="W", padx=35, pady=20)

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
            wmlist = sheet.col_values(2)[1:]
            rmlist = sheet.col_values(3)[1:]
            tmlist = sheet.col_values(4)[1:]
            slist = sheet.col_values(5)[1:]

            self.shoe_error = 0

            self.shoe = tk.StringVar(self.shoe_miles_window)
            self.shoe.set('')
            self.top_frame = tk.Frame(self.shoe_miles_window)
            self.shoe_warning_frame = tk.Frame(self.shoe_miles_window)
            self.middle_frame = tk.Frame(self.shoe_miles_window)
            self.bottom_frame = tk.Frame(self.shoe_miles_window)
            self.shoe_lookup = tk.Label(self.top_frame, text='Which shoes would you like to view?')
            self.shoe_lookup_box = ttk.Combobox(self.top_frame, values=self.shoe_list(), state='readonly', textvar=self.shoe)
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
            self.no_data_warning = tk.Label(self.warning_frame, text='You do not have any data entered.')
            self.no_data_ok = tk.Button(self.workout_window, text='OK', command=self.workout_window.destroy)
            self.no_data_warning.pack()
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
    def edit_workout_master(self, upperBound):
        def reset():
            self.edit_window_exists = 0
            self.edit_window.destroy()

        def update_row(row_num, lower_bound, delete, date, walk, run, total, shoes, type, delete_box, update, total_var):
            row = lower_bound + row_num
            record = str(row_num+1)

            if delete.get() > 0 and delete.get() == int(record):  # DELETE RECORD IF BOX ON UPDATE ROW IS CHECKED
                date.grid_forget()
                walk.grid_forget()
                run.grid_forget()
                total.grid_forget()
                shoes.grid_forget()
                type.grid_forget()
                delete_box.grid_forget()
                update.grid_forget()
                sheet.delete_row(row)  # DELETE FROM GOOGLE SHEET
                self.edit_window.update()  # REFRESH EDIT WINDOW
            else:
                run_date = tk.StringVar()  # create variable to hold run date
                run_date.set(date.get("1.0","end"))
                test = re.match(dateFormat, date.get("1.0", "end"))  # see if date is formatted correctly
                date.config(foreground="black")
                while test is None:  # PROMPT USER TO ENTER A CORRECTLY FORMATTED DATE IF FORMAT IS INCORRECT
                    date.config(foreground="red")
                    run_date.set(date.get("1.0", "end"))
                    test = re.match(dateFormat, date.get("1.0", "end"))
                    update.wait_variable(run_date)

                sheet.update_cell(row, 1, date.get("1.0", "end"))
                sheet.update_cell(row, 2, walk.get("1.0", "end"))
                sheet.update_cell(row, 3, run.get("1.0", "end"))
                sheet.update_cell(row, 4, int(sheet.cell(row, 2).value) + int(sheet.cell(row, 3).value))
                total_var.set(int(walk.get("1.0", "end"))+int(run.get("1.0", "end")))
                sheet.update_cell(row, 5, shoes.get())
                sheet.update_cell(row, 6, type.get("1.0","end"))
                self.warning(shoes.get())

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
            self.edit_window.protocol("WM_DELETE_WINDOW", reset)
        lowerBound = upperBound-10

        if lowerBound <= 1:
            lowerBound = 2

        # CREATE LIST OF DATA IN ROWS (LIST OF LISTS)
        list_of_rows = []
        for i in range(lowerBound, upperBound):
            row_contents = sheet.row_values(i)[0:6]
            list_of_rows.append(row_contents)

        # CREATE AND DISPLAY TABLE OF 10 RUNS
        # CREATE TABLE HEADER
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
        dates = [date1, date2, date3, date4, date5, date6, date7, date8, date9, date10]

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
        walks = [walk1, walk2, walk3, walk4, walk5, walk6, walk7, walk8, walk9, walk10]

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
        runs = [run1, run2, run3, run4, run5, run6, run7, run8, run9, run10]

        total_var_1 = tk.IntVar()
        total_var_2 = tk.IntVar()
        total_var_3 = tk.IntVar()
        total_var_4 = tk.IntVar()
        total_var_5 = tk.IntVar()
        total_var_6 = tk.IntVar()
        total_var_7 = tk.IntVar()
        total_var_8 = tk.IntVar()
        total_var_9 = tk.IntVar()
        total_var_10 = tk.IntVar()

        r = 0
        for total in [total_var_1, total_var_2, total_var_3, total_var_4, total_var_5, total_var_6, total_var_7, total_var_8, total_var_9, total_var_10]:
            total.set(list_of_rows[r][3])
            r += 1

        total1 = tk.Label(self.edit_window, textvariable=total_var_1, width=5, height=1)
        total2 = tk.Label(self.edit_window, textvariable=total_var_2, width=5, height=1)
        total3 = tk.Label(self.edit_window, textvariable=total_var_3, width=5, height=1)
        total4 = tk.Label(self.edit_window, textvariable=total_var_4, width=5, height=1)
        total5 = tk.Label(self.edit_window, textvariable=total_var_5, width=5, height=1)
        total6 = tk.Label(self.edit_window, textvariable=total_var_6, width=5, height=1)
        total7 = tk.Label(self.edit_window, textvariable=total_var_7, width=5, height=1)
        total8 = tk.Label(self.edit_window, textvariable=total_var_8, width=5, height=1)
        total9 = tk.Label(self.edit_window, textvariable=total_var_9, width=5, height=1)
        total10 = tk.Label(self.edit_window, textvariable=total_var_10, width=5, height=1)

        totals = [total1, total2, total3, total4, total5, total6, total7, total8, total9, total10]

        shoe_var1 = tk.StringVar()
        shoe_var2 = tk.StringVar()
        shoe_var3 = tk.StringVar()
        shoe_var4 = tk.StringVar()
        shoe_var5 = tk.StringVar()
        shoe_var6 = tk.StringVar()
        shoe_var7 = tk.StringVar()
        shoe_var8 = tk.StringVar()
        shoe_var9 = tk.StringVar()
        shoe_var10 = tk.StringVar()

        r = 0
        for shoe in [shoe_var1, shoe_var2, shoe_var3, shoe_var4, shoe_var5, shoe_var6, shoe_var7, shoe_var8, shoe_var9, shoe_var10]:
            shoe.set(list_of_rows[r][4])
            r += 1

        shoes1 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var1)
        shoes2 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var2)
        shoes3 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var3)
        shoes4 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var4)
        shoes5 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var5)
        shoes6 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var6)
        shoes7 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var7)
        shoes8 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var8)
        shoes9 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var9)
        shoes10 = ttk.Combobox(self.edit_window, values=self.shoe_list(), textvar=shoe_var10)

        print(list_of_rows[9][4])
        print(shoe_var10.get())
        print(shoes10.get())

        shoeses = [shoes1, shoes2, shoes3, shoes4, shoes5, shoes6, shoes7, shoes8, shoes9, shoes10]

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
        types = [type1, type2, type3, type4, type5, type6, type7, type8, type9, type10]

        delete = tk.IntVar()  # CREATE INTVAR TO TRACK WHICH BOXES ARE CHECKED
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
        deletes = [delete_box_1, delete_box_2, delete_box_3, delete_box_4, delete_box_5, delete_box_6, delete_box_7, delete_box_8, delete_box_9, delete_box_10]

        update1 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(0, lowerBound, delete, date1, walk1, run1, total1, shoes1, type1, delete_box_1, update1, total_var_1))
        update2 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(1, lowerBound, delete, date2, walk2, run2, total2, shoes2, type2, delete_box_2, update2, total_var_2))
        update3 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(2, lowerBound, delete, date3, walk3, run3, total3, shoes3, type3, delete_box_3, update3, total_var_3))
        update4 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(3, lowerBound, delete, date4, walk4, run4, total4, shoes4, type4, delete_box_4, update4, total_var_4))
        update5 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(4, lowerBound, delete, date5, walk5, run5, total5, shoes5, type5, delete_box_5, update5, total_var_5))
        update6 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(5, lowerBound, delete, date6, walk6, run6, total6, shoes6, type6, delete_box_6, update6, total_var_6))
        update7 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(6, lowerBound, delete, date7, walk7, run7, total7, shoes7, type7, delete_box_7, update7, total_var_7))
        update8 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(7, lowerBound, delete, date8, walk8, run8, total8, shoes8, type8, delete_box_8, update8, total_var_8))
        update9 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(8, lowerBound, delete, date9, walk9, run9, total9, shoes9, type9, delete_box_9, update9, total_var_9))
        update10 = tk.Button(self.edit_window, text="Update", command=lambda *args: update_row(9, lowerBound, delete, date10, walk10, run10, total10, shoes10, type10, delete_box_10, update10, total_var_10))
        updates = [update1, update2, update3, update4, update5, update6, update7, update8, update9, update10]

        fields = [dates, walks, runs, totals, shoeses, types, deletes, updates]

        # CREATE AND POPULATE FIELDS IN EDIT WINDOW
        c = 0  # COLUMN ITERATOR
        for field in fields:  # LOOP THROUGH COLUMNS
            r = 1  # ROW ITERATOR
            for item in field:  # LOOP THROUGH ROWS
                item.grid(row=r, column=c, padx=5)  # PLACE BOX IN WINDOW
                if field in fields[0:3] or field in fields[5:6]:
                    item.insert('1.0', chars=list_of_rows[r-1][c])  # POPULATE BOX FROM LIST OF ROWS FROM GOOGLE SHEET
                r += 1
            c += 1

        prev_button = tk.Button(self.edit_window, text='<<', command=prev_runs).grid(row=11, column=0, pady=20)
        next_button = tk.Button(self.edit_window, text='>>', command=next_runs).grid(row=11, column=7)
        exit_button = tk.Button(self.edit_window, text='Exit', command=self.edit_window.destroy).grid(row=11, column=4)

        if not self.edit_window.winfo_ismapped():  # CREATE WINDOW IF IT DOES NOT EXIST, REFRESH IF IT DOES
            self.edit_window_exists = 1
            self.edit_window.mainloop()
        else:
            shoes10.update()
            self.edit_window.update()
            shoes10.update()


    # FUNCTION TO SORT WORKBOOK - CURRENTLY RUNS VERY SLOWLY
    def sort_log(self):
        all_values = sheet.get_all_values()  # GET A LIST OF ALL OF THE ROWS IN THE WORKSHEET, EACH ROW IS A LIST
        all_values = all_values[1:]  # REMOVE HEADER ROW FROM LIST OF VALUES
        all_values_sorted = sorted(all_values, key=itemgetter(0))  # SORT THE LIST BY DATE

        r = 2  # INITIALIZE VAR FOR ROW TO START UPDATING FROM

        for line in all_values_sorted: # LOOP THROUGH ROWS
            c = 1  # INITIALIZE VAR FOR COL TO START UPDATING FROM - REFRESH FOR EACH ITERATION
            for cell in line: # LOOP THROUGH CELLS IN ROW
                sheet.update_cell(r, c, cell)  # UPDATE CELL WITH SORTED DATA
                c += 1  # INCREASE COL COUNTER
            r += 1  # INCREASE ROW COUNTER

        tk.messagebox.showinfo("Log Sorting", "Sorting complete!")


################################


RunningLog()  # CREATE INSTANCE OF PROGRAM
