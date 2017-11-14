#Python 2.7

# MAKE SURE THAT DRIVE LETTER Y: IS MAPPED TO SHAREPOINT's Premier Customer folder!

# NOTE: Navigating using Windows Explorer with SharePoint
# ...is not the fastest with SharePoint.
# ...Use the REST API.
# Nevermind... we are switching to Google drive / Box (see Box python module)

import msvcrt
import time
from timeit import default_timer as timer


#Note that the following SH Libraries need to be in the same directory as the running code's file
#(or in the pyhton path)
from shLib import *
#from slDeck import createSlideDeck
#from slLib import createSlideDeck
from slDeck import loadDbTables, createSlideDeck
from slLib import Table_Options

#--------------------------------------------
#Stop script to test slide design without archiving zip reports
#This setting only applies to LAB environment 
slideDesign = True
#slideDesign = False
#--------------------------------------------



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
#(make sure first that there are no existing san health folders with zip files
#before activating feature!

# 2) archive other non-SH zip files in the san health repository.
#(will speed up scanning)

# 3) convert all the raise exceptions to continue monitoring.
# (maybe the next time around things would have settled!)

# -----------------------------------------------------------------
# -----------------------------------------------------------------
#initialize table's default options
options = Table_Options()

#make a start log entry
logEntry("Started", siteEnv)

#delete any old files in collector folder
initFolders()

loopCount = 0

###this is for shLib\archiveBad but it is not working.
##global blacklist
##blacklist = [] #keep track of bad files that can't be moved

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
        #print "Could not retrieve the list of folders"
        logEntry("Folder List Error", "Could not retrieve folders. Waiting 10 sec")
        #Wait in case there are temporary network issues
        time.sleep(10)
        #And start a new loop
        continue


#-------------------------------------------------------
#PUT THIS SECTION INTO A FUNCTION: GET LIST OF REPORTS WITH VALID CSVS
    for customer in customers:
        #clear the csvPath and san lists
        options.csvPathList = []
        options.sanList = []
        #options.shFiles = []

        #download all zip files (if any) to local collector folder
        #(collector folder should be empty)
        #WHEN DOES THIS COLLECTOR FOLDER GETS RECREATED? -------------XXXXXXXXXXXXXXXXX
        #RENAME THIS TO COLLECTSHZIPFILES
        if getZipFiles(customer) == None:
            continue

        #now we have all SH ZIP files in the collector folder
        #Look inside each ZIP file for another ZIP with the CSV files
        
        for shFile in get_shFiles():
            #print 'shFile', shFile
            
            options.custData = None
                                    
            data = extract_csvFiles(shFile)
            if data == None:
                #no csv files in this sh-zip file
                logEntry('No CSVs', shFile)
                archiveBad(customer, shFile)
                continue

            if data == 'Bad':
                #print 'Bad zip', shFile
                logEntry('Bad Zip', shFile)
                archiveBad(customer, shFile)
                continue
            
            #This SH Report has CSV files. 
            
            # #Create slide deck with current SAN / SH Report variables 
            csvPath, shName, sanName, shDate, shYear = data
            
            #add to a list all the sh names common to each report's csv files
            options.sanList.append( (shDate, shName, shFile, sanName, csvPath) )

            # 
            # #add the customer folder name to the report variables' tuple
            custData = (customer,) + data
            # #store customer variables tuple into options
            options.custData = custData
            
            #do not archive the remote report zip files for this customer
            #place this in a config file!
            options.archv_opt = 'no_remote'
            
        #All zipped sh reports have been collected and csv files extracted

        # CHECK IF THERE ARE ANY VALID SH REPORTS (WITH CSV FILES)
        # FOR THIS CUSTOMER. IF NOT, GO TO NEXT CUSTOMER
        # OTHERWISE, CREATE SLIDE DECKS FOR THIS CUSTOMER

        # ------------------------------------------
        # populate the database
        # from all the CSV files from all the SH reports in the folder
        # Remove this function from slDeck.py?
        #if options.custData == None:
        if len(options.sanList) == 0:
            #no reports with valid csv files for this customer.
            #go to next customer.
            continue

        #Open the database and load one table per csv file
        
        loadDbTables(options)
        #create Slide Deck(s) 
        createSlideDeck(options)
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        #To test slide design. 
        #Stop after slides creation but before deleting SH reports.
        if slideDesign and siteEnv == 'LAB':
            print ''
            print '------------------------'
            print 'Stopping to check slides'
            quit()
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!        
        
        #Archive only the local SH Zip files ('no_remote')
        #to avoid excesive storage usage.
        #archiveFiles(shFile, custData, 'no_remote')
        archiveFiles(options)
        #delete the csv directory to remove the used files
        #and delete all files from the collector folder
        initFolders()

    #used only for the console loop counter
    end = timer()
    tTime = time.strftime("%X")
    tDate = time.strftime("%x")
    print tDate, tTime, 'Loop:', loopCount, '-', len(customers), 'Folders', '-', format(round(end-start, 1), '.1f'), 'Seconds'


    if msvcrt.kbhit():
        #print 'KEY:', msvcrt.getch()
        if msvcrt.getch() == 'q':
            logEntry("Stopped")
            raise SystemExit(0)
    








            
