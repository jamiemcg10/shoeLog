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
import time
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from beautifultable import BeautifulTable
from operator import itemgetter


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Running Log").sheet1


dateFormat = re.compile('\d{2}/\d{2}/\d{4}') ## CREATE AND COMPUTE GLOBAL VARIABLE TO TEST FOR VALID DATE FORMATS


### DEFINE FUNCTIONS
# GREETING STARTS THE PROGRAM
def main():
    today = datetime.datetime.now()
    os.system('cls')
    print "\nWelcome to Shoe Log! Today's date is %s.\n" %today.strftime("%m/%d/%Y")
    time.sleep(2)

    branch()

# DISPLAYS A MENU OF PROGRAM CHOICES:
def branch():

    os.system('cls')
    print "Type 'Ctrl+C' at any time to quit the program.\n\n"
    option = input("Please choose an option from the menu below:\n" +
                    "1: Log a run\n" +
                    "2: View shoe data\n" +
                    "3: Calculate miles run\n" +
                    "4: View/Edit workout data\n" +
                    "5: Exit program\n\n")

    while option not in [1,2,3,4,5]: ## PROMPT FOR A VALID SELECTION IF ONE IS NOT ENTERED

        option = input("The selection that you entered is invalid. Please make a selection from the list:\n" +
                    "1: Log a run\n" +
                    "2: View shoe data\n" +
                    "3: Calculate miles run\n" +
                    "4: View/Edit workout data\n" +
                    "5: Exit program\n\n")
        os.system('cls')

    if option == 1:
        logrun()

    if option == 2:
        shoemiles()

    if option == 3:
        workouts()

    if option == 4:
        upperBound = sheet.row_count+1
        editWorkout(upperBound)

    if option == 5:
        os.system('cls')
        print "\nThank you for using Shoe Log!"
        time.sleep(2)
        os.system('cls')


# LOGS A RUN
def logrun():
    rundate = raw_input("\nDate of your run (format mm/dd/yyyy): ")
    test = re.match(dateFormat,rundate)
    while test is None: ## USER IS PROMTED TO ENTER A CORRECTLY FORMATTED DATE IF FORMAT IS INCORRECT
        rundate = raw_input("The date you entered is not in the correct format. Please format your date correctly (mm/dd/yyyy): ")
        test = re.match(dateFormat,rundate)
    shoe = raw_input("What shoes did you wear? ").title()
    walk_miles = raw_input("How many miles did you walk in these shoes? ")
    run_miles = raw_input("How many miles did you run in these shoes? ")

    # REPLACE MISSING ENTRIES FOR MILES RUN/WALKED WITH 0
    if walk_miles == '':
        walk_miles = '0'
    if run_miles == '':
        run_miles = '0'

    total_miles = str(float(walk_miles) + float(run_miles))
    run_type = raw_input("What kind of run was this? ").title()

    # REPLACE OTHER BLANKS WITH QUESTION MARKS
    if shoe == '':
        shoe = '?'
    if run_type == '':
        run_type = '?'


    print "\nYou ran " + run_miles + " miles and walked " + walk_miles + " miles for a total of " + total_miles + " miles in your " + shoe + " on " + rundate + "."

    record = rundate,walk_miles,run_miles,total_miles,shoe,run_type


    # Create header row if one doesn't exist, and add run to spreadsheet
    if sheet.cell(1,1).value == "":
        header = "Date","Miles walked","Miles run","Total miles","Shoes","Type of run"
        sheet.insert_row(header,1)
        sheet.insert_row(record,2)
    else:
        sheet.insert_row(record,sheet.row_count+1)

    print "Your run has been recorded.\n"

    warning(shoe) ## DISPLAYS A WARNING IF THE SHOE HAS >250 MILES

    time.sleep(2.5)
    branch() ## GO BACK TO THE MENU


# CALCULATES NUMBER OF MILES ON SHOES
def shoemiles():
    if sheet.row_count < 2: ## CHECKS TO SEE THAT THERE IS DATA IN THE SPREADSHEET
        print "You do not have any data entered.\n"
    else: ## CREATE LISTS OF THE DIFFERENT DATA COLUMNS
        rdlist = sheet.col_values(1)[1:]
        wmlist = sheet.col_values(2)[1:]
        rmlist = sheet.col_values(3)[1:]
        tmlist = sheet.col_values(4)[1:]
        slist = sheet.col_values(5)[1:]
        rtlist = sheet.col_values(6)[1:]


        # Create and display a list of shoes in the spreadsheet
        list_of_shoes = set(slist)
        print "\nShoe list:"
        num = 1
        shoeDict = {}
        for s in list_of_shoes:
            shoeDict[num] = s
            print str(num) + ": " + s
            num += 1


        shoeLookup = raw_input("\nWhich shoes listed above do you want to view? ").title()


        try:  ## Convert the number to a shoe if a number is given
            shoeLookup = int(shoeLookup)
            shoeLookup = shoeDict[shoeLookup]
        except ValueError:
            pass


        while shoeLookup not in list_of_shoes:
            shoeLookup = raw_input("There is no data for the shoes you entered. Please enter valid shoes: ")

        # Create variables for the number of miles walked, run, and total, and add to them if the shoe in the row matches the shoe entered
        walkReturn = 0.0
        runReturn = 0.0
        totalReturn = 0.0
        for i in range(0,len(slist)):
            if slist[i] == shoeLookup:
                walkReturn += float(wmlist[i])
                runReturn += float(rmlist[i])
                totalReturn += float(tmlist[i])

        print "\nYou have walked %.2f miles and run %.2f miles for a total of %.2f miles in your %s.\n\n" %(walkReturn, runReturn, totalReturn, shoeLookup)


    time.sleep(2.5)
    branch() ## GO BACK TO MENU



def workouts():

    if sheet.row_count < 2:
        print "\nYou do not have any data entered.\n"
    else:
        periodStart = raw_input("\nWhat is the beginning of the date range you want to review (format mm/dd/yyyy)? ")
        while re.match(dateFormat,periodStart) is None:
            periodStart = raw_input("\nPlease format your date correctly. What is the beginning of the date range you want to review (format mm/dd/yyyy)? ")

        periodEnd = raw_input("What is the end of the date range you want to review (format mm/dd/yyyy)? ")
        while re.match(dateFormat,periodEnd) is None:
                    periodEnd = raw_input("Please format your date correctly. What is the end of the date range you want to review (format mm/dd/yyyy)? ")


        ## CONVERT DATES FROM STRING TO DATE FORMAT SO DATES CAN BE COMPARED
        periodStart = periodStart.split('/')
        periodEnd = periodEnd.split('/')
        start = datetime.date(int(periodStart[2]),int(periodStart[0]),int(periodStart[1]))
        end = datetime.date(int(periodEnd[2]),int(periodEnd[0]),int(periodEnd[1]))

        runReturn = 0.0

        for i in range(2,sheet.row_count+1): ## COMPARE ROW DATE TO RANGE ENTERED AND ADD MILES RUN ON MATCHING DATES TO RUNNING TOTAL
            rowDate = sheet.cell(i,1).value.split('/')
            rowDate = datetime.date(int(rowDate[2]),int(rowDate[0]),int(rowDate[1]))
            if (rowDate <= end and rowDate >= start):
                runReturn += float(sheet.cell(i,3).value)

        ## CONVERT DATES BACK TO STRING FORMAT FOR PRINTING
        periodStart = periodStart[0]+"/"+periodStart[1]+"/"+periodStart[2]
        periodEnd = periodEnd[0]+"/"+periodEnd[1]+"/"+periodEnd[2]

        print "\nYou ran %.2f miles between %s and %s.\n" %(runReturn,periodStart,periodEnd)

    time.sleep(2.5)
    branch() ## GO TO MENU


# Displays a warning if the shoe passed in has more than 250 miles recorded.
def warning(shoe):
    if sheet.row_count < 2:
        pass
    else:
        tmlist = sheet.col_values(4)[1:]
        slist = sheet.col_values(5)[1:]

        shoeWarn = 0.0
        for i in range(0,len(slist)):
            if slist[i] == shoe:
                shoeWarn += float(tmlist[i])


        if shoeWarn > 500:
            print "WARNING: \nThese shoes (%s) have over 500 miles. It is time to replace them.\n" %(shoe)
            time.sleep(1)
        elif shoeWarn > 450:
            print "\WARNING: These shoes (%s) have over 450 miles. It is time to consider replacing them.\n" %(shoe)
            time.sleep(1)
        elif shoeWarn > 400:
            print "\WARNING: These shoes (%s) have over 400 miles. It is time to consider replacing them.\n" %(shoe)
            time.sleep(1)
        elif shoeWarn > 350:
            print "\WARNING: These shoes (%s) have over 350 miles. It is time to consider replacing them.\n" %(shoe)
            time.sleep(1)
        elif shoeWarn > 300:
            print "\WARNING: These shoes (%s) have over 300 miles. You may need to replace them soon.\n" %(shoe)
            time.sleep(1)
        elif shoeWarn > 250:
            print "\WARNINGs: These shoes (%s) have over 250 miles. You may need to replace them soon.\n" %(shoe)
            time.sleep(1)
        else:
            pass


# View or edit data
def editWorkout(upperBound):
    os.system('cls')
    lowerBound = upperBound-10

    if lowerBound <= 1:
        lowerBound = 2

    ## CREATE AND DISPLAY TABLE OF X RUNS
    table = BeautifulTable()
    table.column_headers = ['','Date','Mi. walked','Mi. run','Total mi.','Shoes','Run type']

    rowNum = 1
    rowList = []
    for i in range(lowerBound,upperBound):
        rowContents = sheet.row_values(i)[0:6]
        rowContents.insert(0,rowNum)
        rowList.append(rowContents)
        table.append_row(rowContents)
        rowNum += 1

    table.row_seperator_char = ""
    table.intersection_char = ""
    table.column_seperator_char = ""

    print table
    if lowerBound == 2:
        print "BEGINNING"
    if upperBound == sheet.row_count+1:
        print "END"

    next = raw_input("\n\nEnter the number of the row you want to edit, enter 'P' to see earlier runs, 'N' to see later runs, 'S' to sort runs, or 'E' to exit the editor. ")

    while next.upper() not in ["P","E","N","S","1","2","3","4","5","6","7","8","9","10"]:
        next = raw_input("\nPlease enter a valid selection. Enter the number of the row you want to edit, enter 'P' to see earlier runs, 'N' to see later runs, 'S' to sort runs, or 'E' to exit the viewer. ")
        os.system('cls')

    if next.upper() == "P":
        if lowerBound == 2:
            editWorkout(upperBound)
        else:
            editWorkout(lowerBound)
    elif next.upper() == "E":
        os.system('cls')
        branch()
    elif next.upper() == "N":
        upperBound += 10
        if upperBound > sheet.row_count:
            upperBound = sheet.row_count
        else:
            pass
        editWorkout(upperBound+1)
    elif next.upper() == "S":
		sortLog()
		upperBound = sheet.row_count+1
		editWorkout(upperBound)
    else:
        next = int(next)
        editTable = BeautifulTable()
        editTable.column_headers = ['','Date','Mi. walked','Mi. run','Total mi.','Shoes','Run type']
        editRow = sheet.row_values(lowerBound-1+next)[0:6]
        editRow.insert(0,next)
        editTable.append_row(editRow)
        os.system('cls')

        print editTable

        edit = raw_input("\nWhat do you want to edit?\n" +
                        "1: Date\n" +
                        "2: Miles walked\n" +
                        "3: Miles run\n" +
                        "4: Shoes\n" +
                        "5: Run type\n" +
                        "6: Delete run\n" +
                        "\n7: Cancel\n\n")
        while edit not in ["1","2","3","4","5","6","7"]:
            os.system('cls')
            print editTable
            print "\nPlease make a valid selection."
            edit = raw_input("\nWhat do you want to edit?\n" +
                        "1: Date\n" +
                        "2: Miles walked\n" +
                        "3: Miles run\n" +
                        "4: Shoes\n" +
                        "5: Run type\n" +
                        "6: Delete run\n" +
                        "\n7: Cancel\n\n")

        if edit == "1":
            rundate = raw_input("\nDate of your run (format mm/dd/yyyy): ")
            test = re.match(dateFormat,rundate)
            while test is None:
                rundate = raw_input("\nThe date you entered is not in the correct format. Please format your date correctly (mm/dd/yyyy): ")
                test = re.match(dateFormat,rundate)
            sheet.update_cell(lowerBound-1+next,1,rundate)
            print "Your run has been updated."
            time.sleep(1)
            editWorkout(upperBound)

        if edit == "2":
            sheet.update_cell(lowerBound-1+next,2,input("\nHow many miles did you walk in these shoes? "))
            sheet.update_cell(lowerBound-1+next,4,(float(sheet.cell(lowerBound-1+next,2).value) + float(sheet.cell(lowerBound-1+next,3).value)))
            print "Your run has been updated."
            time.sleep(1)
            editWorkout(upperBound)

        if edit == "3":
            sheet.update_cell(lowerBound-1+next,3,input("\nHow many miles did you run in these shoes? "))
            sheet.update_cell(lowerBound-1+next,4,(float(sheet.cell(lowerBound-1+next,2).value) + float(sheet.cell(lowerBound-1+next,3).value)))
            print "Your run has been updated."
            time.sleep(1)
            editWorkout(upperBound)

        if edit == "4":
            sheet.update_cell(lowerBound-1+next,5,raw_input("\nWhich shoes did you wear? ").title())
            print "Your run has been updated."
            time.sleep(1)
            editWorkout(upperBound)

        if edit == "5":
            sheet.update_cell(lowerBound-1+next,6,raw_input("\nWhat type of run was this? ").title())
            print "Your run has been updated."
            time.sleep(1)
            editWorkout(upperBound)

        if edit == "6":
            confirm = raw_input("\nAre you sure you want to delete this run? (Y/N) ")
            if confirm.upper() == "Y" or confirm.upper() == "YES":
                sheet.delete_row(lowerBound-1+next)
                print "This run has been deleted."
                time.sleep(1)
                editWorkout(sheet.row_count+1)
            else:
                print "This run will not be deleted."
                editWorkout(upperBound)

        if edit == "7":
            branch()

# FUNCTION TO SORT WORKBOOK - CURRENTLY RUNS VERY SLOWLY
def sortLog():
    all_values = sheet.get_all_values() # GET A LIST OF ALL OF THE ROWS IN THE WORKSHEET, EACH ROW IS A LIST
    all_values = all_values[1:] # REMOVE HEADER ROW FROM LIST OF VALUES
    all_values_sorted = sorted(all_values, key=itemgetter(0)) # SORT THE LIST BY DATE

    os.system('cls')

	#print all_values_sorted # PRINT THE SORTED LIST (FOR TESTING)

    print '\nSorting takes time. Please be patient.\n'
    print 'Sorting runs...\n'

    r=2 # INITIALIZE VAR FOR ROW TO START UPDATING FROM

    for line in all_values_sorted: # LOOP THROUGH ROWS
        c=1 # INITIALIZE VAR FOR COL TO START UPDATING FROM - REFRESH FOR EACH ITERATION
        for cell in line: # LOOP THROUGH CELLS IN ROW
            #print r, c, cell # PRINT ROW NUM, COL NUM AND CELL CONTENTS (FOR TESTING)
            sheet.update_cell(r,c,cell) # UPDATE CELL WITH SORTED DATA
            c += 1 # INCREASE COL COUNTER
        r += 1 # INCREASE ROW COUNTER

    print 'Sorting complete.'


################################


try:
    main()
except KeyboardInterrupt:
    os.system('cls')
    print 'Thank you for using Shoe Log!'
    time.sleep(2)
    os.system('cls')