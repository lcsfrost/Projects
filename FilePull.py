from ast import Try
import glob
import openpyxl
import os
import warnings

warnings.simplefilter("ignore")


Reports = []
Reports = glob.glob("*\*.xls*",) #Generates a python list of all .xls file types in the directory, including subfolders. Returns temp files, .xls, .xlsx, .xlsm, etc.
f = open("infodump.csv", "w") #Creates spreadsheet to dump info into.
for report in Reports: 
    try:
        if '~' not in report: # This is to filter out temporary excel instances from recovered files
            if report.endswith('.xlsm') or report.endswith('.xlsx'): # Filters out basic .xls files that can't be handled by OpenPyxl
                wb = openpyxl.load_workbook(report,data_only=True,read_only = True) #Loads up workbook specified 
                print ("\n\n\n" + os.path.basename(report)) #not needed just makes it look nice while it's running


                #Below code blocks pull out relevant information from the first page of each spreadsheet. 
                QuoteNumber = str(wb.worksheets[0]['K11'].value or '')
                PONumber = str(wb.worksheets[0]['K12'].value or '')
                WorkOrderNumber = str(wb.worksheets[0]['N2'].value or '') + " " + str(wb.worksheets[0]['P2'].value or '')
                OrderCreateDate = str(wb.worksheets[0]['K4'].value or '')
                ShipDate = str(wb.worksheets[0]['O4'].value or '')
                CustomerName = str(wb.worksheets[0]['K5'].value or '')
                Carrier = str(wb.worksheets[0]['O8'].value or '')
                ShortDescription = str(wb.worksheets[0]['O12'].value or '')
                RepID = str(wb.worksheets[0]['O11'].value or '')
                QuoteWeight = str(wb.worksheets[0]['P13'].value or '')
                GrandTotal = str(wb.worksheets[0]['G27'].value or '')
                QuoteHoursRange = wb.worksheets[0]['E6':'E12'] or ''
                ActualHoursRange = wb.worksheets[0]['F6':'F12'] or ''
                HoursTypeRange = wb.worksheets[0]['C6':'C12'] or ''
                QuoteHours = ""
                ActualHours = ""
                HoursType = ""
                for row in QuoteHoursRange:
                    for cell in row:     
                        QuoteHours = QuoteHours + str(cell.value or '') + ","
                for row in ActualHoursRange:
                    for cell in row:     
                        ActualHours = ActualHours + str(cell.value or '') + ","
                for row in HoursTypeRange:
                    for cell in row:
                        HoursType = HoursType + str(cell.value or '') + ","


                #FinalString creates a string to add to a csv file
                FinalString = report + "," + QuoteNumber + "," + PONumber + "," + WorkOrderNumber + "," + OrderCreateDate + "," + ShipDate + "," + CustomerName + "," + Carrier + "," + ShortDescription + "," + RepID + "," + QuoteWeight + "," + GrandTotal + "," + HoursType +QuoteHours + ActualHours + "\n"
                print ("       Work order " + WorkOrderNumber + " written to: infodump.csv") #Not needed and can be removed.
                f.write(FinalString) #Writes string to CSV file
    except:
        pass