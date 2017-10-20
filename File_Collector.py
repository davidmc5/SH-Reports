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
#from slDeck import createSlideDeck
from slDeck import createSlideDeck
from slDeck import loadDbTables

from slLib import Table_Options
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
#(make sure first that there are no existing san health folders with zip files
#before activating feature!

# 2) archive other non-SH zip files in the san health repository.
#(will speed up scanning)

# 3) convert all the raise exceptions to continue monitoring.
# (maybe the next time around things would have settled!)

# -----------------------------------------------------------------
# -----------------------------------------------------------------
#determine server environment (lab or production)
if os.path.isdir(lab_path[3]):
    print lab_path[3]
    location = lab_path
    
if os.path.isdir(prod_path[3]):
    print prod_path[3]
    location = prod_path

#set paths based on server environment
drive,\
startFolder,\
shFolder,\
localWorkFolder,\
slides_template_path,\
csvLogFile,\
tmpLogFile = location
# -----------------------------------------------------------------
# -----------------------------------------------------------------


#make a start log entry
logEntry("Started")

#delete any old files in collector folder
initFolders()

loopCount = 0

#initialize table's default options
options = Table_Options()

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
            
            #This SH Report has CSV files. Create slide deck
            #current SAN / SH Report variables 
            csvPath, shName, sanName, shYear = data
            
            #save report's name and file name to archive.
            #options.shFiles.append( (shName, shFile) )

            #add to a list all the sh names common to each report's csv files
            #options.csvPathList.append(csvPath)
            options.sanList.append( (shName, shFile, sanName, csvPath) )

            #add the customer folder name to the report variables' tuple
            custData = (customer,) + data
            options.custData = custData
            #print custData
            
            #do not archive the remote report zip files for this customer
            #place this in a config file!
            options.archv_opt = 'no_remote'

        # Store the list of all the SH.ZIP file names processed for this customer
        #options.shFiles = shFileNames    

        # ------------------------------------------
        # populate the database
        # from all the CSV files from all the SH reports in the folder
        # Remove this function from slDeck.py?
        
        #Open the database and load one table per csv file
        loadDbTables(options)

        #quit()
        
        #create Slide Deck(s) 
        createSlideDeck(options)

        #Archive only the local SH Zip files ('no_remote')
        #to avoid excesive storage usage.
        #archiveFiles(shFile, custData, 'no_remote')
        archiveFiles(options)

        #logEntry('Slides Created', customer, shName)
        
        #delete the csv directory to remove the used files
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
    








            
