#Python 2.7

# MAKE SURE THAT DRIVE LETTER Y: IS MAPPED TO SHAREPOINT's Premier Customer folder!

# NOTE: Navigating using Windows Explorer with SharePoint
# ...is not the fastest with SharePoint.
# ...Use the REST API.
# Nevermind... we are switching to Google drive / Box (see Box python module)

import time
from timeit import default_timer as timer


#Note that the following SH Libraries need to be in the same directory as the running code's file
#(or in the pyhton path)
from shLib import *
from slDeck import createSlideDeck
import msvcrt

'''
On windows, it's the CMD console that closes,
because the Python process exits at the end.
To prevent this, open the console first,
then use the command line to run your script.
Do this by navigating to the folder that contains the script,
--Windows 10: on the top menu, select File Open Command Prompt
--Windows 7: shift right clickon folder, open command window here
and typing in python scriptname.py in the console.
'''


# ----------------------------------------------------------------
# TO DO
# ------------------
# 1) archive non SH zip files to ARCHIVE so that we don't waste time on each loop
#(make sure first that there are no existing san health folders with zip files before activating feature!

# 2) archive other non-SH zip files in the san health repository.
#(will speed up scanning)

# 3) convert all the raise exceptions to continue monitoring.
# (maybe the next time around things would have settled!)

# -----------------------------------------------------------------

#delete any old files in collector folder
initFolders()

#make a start log entry
logEntry("Started")


loopCount = 0

while True:
    
    # clear the keyboard buffer
    while msvcrt.kbhit():
        msvcrt.getch()

    start = timer()
    loopCount += 1

    
    #get a list of customer folders
    try:
        customers = os.listdir(drive + startFolder)
    except:
        print "Could not retrieve the list of folders"
        #Wait 5 minutes in case there are temporary networt issues
        time.sleep(300)
        #And start a new loop
        continue

    for customer in customers:

        #download zip files (if any) to local collector folder
        #(collector folder should be empty)
        #WHEN DOES THIS COLLECTOR FOLDER GETS RECREATED? -------------XXXXXXXXXXXXXXXXX
        #RENAME THIS TO COLLECTSHZIPFILES
        if getZipFiles(customer) == None:
            continue

        #now we have all SH ZIP files in the collector folder
        #Look inside each ZIP file for another ZIP with the CSV files
        for shFile in get_shFiles():
            
            data = extract_csvFiles(shFile)
            if data == None:
                #no csv files in this sh-zip file
                continue

            if data == 'Bad':
                print 'Bad zip', shFile
                logEntry('Bad zip', shFile)
                #add here a check so that if the same bad file
                #is downloaded more than once, archive it to /TEMP
                #archiveBad(customer, shFile)
                #allow a second chance in case the download happened
                #before the upload from the user has finished. 

                #for the time being, just skip
                continue
            
            #current SH report variables 
            csvPath, shName, sanName, shYear = data

            #add the customer folder name to the report variables' tuple
            custData = (customer,) + data

            createSlideDeck(custData)

            archiveFiles(shFile, custData, 'no_remote')

            logEntry('Slides Created', customer, shName)
            
        #delete the csv directory to remove the used files
        initFolders()

    #used only for the logs
    end = timer()
    tTime = time.strftime("%X")
    tDate = time.strftime("%x")
    print tDate, tTime, 'Loop:', loopCount, '-', len(customers), 'Folders', '-', format(round(end-start, 1), '.1f'), 'Seconds'

    #logData = ['Loop:', loopCount, '-', len(customers), 'Folders']
    #logEntry(logData)

    if msvcrt.kbhit():
        #print 'KEY:', msvcrt.getch()
        if msvcrt.getch() == 'q':
            logEntry("Stopped")
            raise SystemExit(0)
    








            
