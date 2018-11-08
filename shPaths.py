

import os

# zip file inside the SH ZIP that contains the CSV files
csvZipFileName = '_CSVReports.zip'

# TODO:
# THE Download folder needs to be created manually or it will crash.
# Set a try/except to either create that folder or print an error.
#---------------------------------------------------
###<------------- PRODUCTION
#---------------------------------------------------
def qa_path():
    ###QA Testing
    siteEnv = 'QA'

    #REMOTE STORAGE
    ##'Z: --> F:\Users\David\'
    # If the repository is on the same server as the application
    #(instead of in a network share) it will result in a "permision denied"
    #...when trying to save the slide decks. Not sure yet why. The file command is expecting a remote share.

    ##drive = 'D:\Users\David\Box Sync' #manually mapped on Windows to Z: to the parent of 'startFolder'
    drive = 'Z:'

    startFolder = '/Repository/'
    shFolder = '/' # Subfolder of the customer's folder where raw SH reports are uploaded. If just '/' then use the Customer's folder
    #LOCAL SERVER
    localWorkFolder = 'D:/Users/David/Projects/SH Reports/'
    # Slide Deck Template
    slides_template_path = localWorkFolder + 'shTemplate.pptx'
    #log files
    csvLogFile = localWorkFolder + 'shLog.csv'
    tmpLogFile = localWorkFolder + 'shLogTemp.csv'
    #----common
    collectorFolder =  localWorkFolder + 'Downloads/'
    csvTempFolder = collectorFolder + 'csvTemp/'
    archiveFolder = collectorFolder + 'ARCHIVE/'

    qa_paths = (
        siteEnv,
        drive, startFolder, shFolder,
        localWorkFolder,
        slides_template_path,
        csvLogFile, tmpLogFile,
        collectorFolder, csvTempFolder, archiveFolder)

    return qa_paths
#---------------------------------------------------
###<------------- PRODUCTION
#---------------------------------------------------
def prod_path():
    ###PRODUCTION
    siteEnv = 'PROD'
    #REMOTE STORAGE
    ##'Y:/http://connect.brocade.com/cs/Customers/Customer%20Information/'
    drive = 'X:' #manually mapped on Windows to the parent of 'startFolder'
    startFolder = '/Premier Customers/'
    shFolder = '/SAN Health/' # Subfolder of the customer's folder where raw SH reports are uploaded. If just '/' then use the Customer's folder
    #LOCAL SERVER
    localWorkFolder = 'C:/users/dmartin/Desktop/SH-Project/'
    slides_template_path = localWorkFolder + 'CODE/' + 'shTemplate.pptx'
    csvLogFile = localWorkFolder + 'shLog.csv'
    tmpLogFile = localWorkFolder + 'shLogTemp.csv'
    #---common
    collectorFolder =  localWorkFolder + 'Downloads/'
    csvTempFolder = collectorFolder + 'csvTemp/'
    archiveFolder = collectorFolder + 'ARCHIVE/'

    prod_paths = (
        siteEnv,
        drive, startFolder, shFolder,
        localWorkFolder,
        slides_template_path,
        csvLogFile, tmpLogFile,
        collectorFolder, csvTempFolder, archiveFolder)

    return prod_paths


#---------------------------------------------------
#<------------- LAB
#---------------------------------------------------
def lab_path():
    ###LAB
    siteEnv = 'LAB'
    #REMOTE STORAGE
    drive = 'Y:' #manually mapped on Windows to the parent of 'startFolder'
    startFolder = '/SH-Repository/'
    shFolder = '/' # Subfolder of the customer's folder where raw SH reports are uploaded. If just '/' then use the Customer's folder
    #LOCAL SERVER
    localWorkFolder = 'E:/Users/David/PROJECTS/SH-Reports/'
    # Slide Deck Template
    slides_template_path = localWorkFolder + 'shTemplate.pptx'
    #log files
    csvLogFile = localWorkFolder + 'shLog.csv'
    tmpLogFile = localWorkFolder + 'shLogTemp.csv'
    #----common
    collectorFolder =  localWorkFolder + 'Downloads/'
    csvTempFolder = collectorFolder + 'csvTemp/'
    archiveFolder = collectorFolder + 'ARCHIVE/'

    lab_paths = (
        siteEnv,
        drive, startFolder, shFolder,
        localWorkFolder,
        slides_template_path,
        csvLogFile, tmpLogFile,
        collectorFolder, csvTempFolder, archiveFolder)

    return lab_paths

#---------------------------------------------------
#---------------------------------------------------

sitePaths = (qa_path(), lab_path(), prod_path())
#sitePaths = (qa_path(), prod_path())

#---------------------------------------------------
#---------------------------------------------------

def setPaths(sitePaths):
# #to find the script path:
# server_path = os.path.dirname(os.path.abspath(__file__))
# #normalizes the path to elimiante extra /
# print os.path.normpath(sitePaths[0][3])
# #use os.path.normcase to set to low case

#determine server environment (lab, production, test, qa, dev, etc)
    for site in sitePaths:
        if os.path.isdir(site[4]): #does localWorkFolder exists?
        #TODO - print an error if no paths exist
            return site
#---------------------------------------------------
#---------------------------------------------------

siteEnv,\
drive,\
startFolder,\
shFolder,\
localWorkFolder,\
slides_template_path,\
csvLogFile,\
tmpLogFile,\
collectorFolder, csvTempFolder, archiveFolder = setPaths(sitePaths)
