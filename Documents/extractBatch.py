# This script was made on Python 3.7.3 on 07/02/2019
# In order for you to run this you need to install Python and then some libraries
# These are the libraries you need to install: tkinter, xlrd, matplotlib
# You can install them opening command prompt and going to :C\\ by "cd .." and then
# typing "pip install _____" ex. "pip install tkinter"

################################################################################
# Author: Josias Ochoa @@@@ josiasochoalozano@gmail.com
################################################################################
# Importing filedialog commands.
from tkinter import filedialog
import os, re, csv

#Prompt User if they want to extract data from Batch Reports.
userisExtractingBatch = input("Are you trying to extract data from Batch Reports? Enter (Y/N): ")

#User says yes!
if userisExtractingBatch == 'Y':
    #Searches the current directory for all files ending in xlsx/xls format. Also setting up arrarys for results and exceptions which will be used later.
    filenames = list(filter(lambda x : re.search("xlsx|xls", x) and not re.search("~", x), os.listdir('.')))
    results = []
    exceptions = []

    #Required libraries to run the stuff below
    import xlrd
    from xlrd import XLRDError

    #This code loops through each excel file for all xlsx and xls files in the directory and parses the data iteratively.
    for filename in filenames:
        #filenames is an array composed of all my excel files.
        print("Parsing data from ", filename)
        notFinished = True

        #The try and except sections handle errors due to encrypted files. I am trying to open an excel workbook. Exception Handling: Get the name of the workbook and make notFinished "true"
        #so that the while loop does not execute.
        try:
            workbook = xlrd.open_workbook(filename, on_demand=True)
        except xlrd.biffh.XLRDError as e:
            #xlrd.biffh.XLRDError is the error caused due to an encrypted file!
            exceptions.append((filename, e))
            notFinished = False
            continue

        # while I have not finished parsing through every file in filenames
        while notFinished:
            try:
                #try to navigate to the charging workbook
                charging = workbook.sheet_by_name('charging')
                chargingExists = workbook.sheet_loaded('charging')
            except:
                #this exception is if charging workbook does not exist.
                print (filename + " failed")
                break

            #if charging workbook exists or is true the following will execute!
            if chargingExists:
                numberRows = charging.nrows
                ##################################################################################################################################################
                ## DEFINING PARAMETERS IS THE MOST IMPORTANT LINE!!! The script looks for strings in this array and gets the values to the right of the cell!!! ##
                ## parameters and parameterValues should be the same size!!!!!!!!!                                                                              ##
                ##################################################################################################################################################
                parameters = ["PRODUCT","CHARGED DATE", "FILLING DATE", "Lot #", "essel", "Batch size - KG", "ime", "Circulation flow rate", "passes", "Filter Lot", "Filter Quantity & Size-MAIN", "Filter Quantity & Size-FINAL", "# of batches"]
                lengthParameters = len(parameters)
                parameterValues = ["","","0.0" , 1.0 , 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, "11"]

                #parameterValue is where I will be storing my numerical data.
                parameterValue = []

                #loop_through_charging is the function that actually iterates through 50 Cells in Col A and gets the values next to appearances of insertedParameter
                def loop_through_charging(insertedParameter):
                    global parameterValue
                    global porValue
                    i = 0
                    valuenotChanged = True
                    #From row 1 to the number of total rows in my excel file
                    for i in range(numberRows):
                        #if the data type in the cell is a string
                        if isinstance(charging.cell_value(i, 0), str):
                            #if the string matches insertedParameter
                            if insertedParameter in charging.cell_value(i, 0):
                                #then get the value of the cell next to it to populate parameterValue which will hold an array of all my strings I was parsing for
                                parameterValue = charging.cell_value(i, 1)
                                if (insertedParameter == "# of batches"):
                                    porValue = charging.cell_value(i, 2)
                                    return parameterValue, porValue
                                i = numberRows
                                valuenotChanged = False
                        if i > 50 and valuenotChanged:
                            parameterValue = "NA"
                            i = numberRows
                    return parameterValue

                #Initializing insertedParameter to be the 0th string in parameters or the first string to look for
                insertedParameter = parameters[0]

                #for each parameter
                for j in range(lengthParameters):
                    jthParameter = parameters[j]
                    #insert each parameter into my function and return an array of my data to an array which will hold these arrays
                    parameterValues[j] = loop_through_charging(jthParameter)

                    if j == lengthParameters-1 :
                        #print(parameterValues[12][1])
                        #print(parameterValues[13])
                        if isinstance(parameterValues[12], tuple):
                            parameterValues[13] = parameterValues[12][1]
                            parameterValues[12] = parameterValues[12][0]

                        results.append(parameterValues)
                        notFinished = False

        #The line below is super important!! if you delete it, .xls files cannot be parsed!
        workbook.release_resources()

    # This line inserts the strings I was originally searching for to appear as the first array within my arrays!
    TitleforCols = ["Product","CHARGED DATE", "FILLING DATE", "Lot #", "Vessel", "Batch size - KG", "Agitation time", "Circulation Flow Rate", "Circulation Passes", "Filter Lot #", "Filter Quantity & Size-MAIN", "Filter Quantity & Size-FINAL", "Filter Usage (Times Used)", "Filter POR Value"]
    results.insert(0, TitleforCols)

    # Ask user for a name of output file and saves it to a CSV
    resultsCsv = input("Enter a name for the results CSV File. Ex. (\"batchreports.csv\"): ")
    with open(resultsCsv, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(results)
    # Ask user if they want to see errors and saves it to a CSV
    seeExceptions = input("Do you want to see exceptions? Enter (Y/N): ")
    if seeExceptions == "Y":
        exceptionsCsv = input("Enter a name for the exceptions CSV File. Ex. (\"exceptions.csv\"): ")
        with open(exceptionsCsv, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(exceptions)

else:
    print("Currently this script only extracts data from Batch Reports. Please use another script for other purposes.")
