

import os

# zip file inside the SH ZIP that contains the CSV files
csvZipFileName = '_CSVReports.zip'

#---------------------------------------------------
###<------------- PRODUCTION
#---------------------------------------------------
def qa_path():
    ###QA Testing
    siteEnv = 'QA'
    #REMOTE STORAGE
    ##'Z: F:\Users\David\'
    # If the repository is on the same server as the application
    #(instead of in a network share) it will result in a "permision denied"
    #when trying to save the slide decks. Not sure yet why.
    drive = 'F:/Users/David' #manually mapped on Windows to the parent of 'startFolder'
    startFolder = '/Repository/'
    shFolder = '/' # SH reports are placed directly on the custmer's folder.
    #LOCAL SERVER
    localWorkFolder = 'F:/Users/David/Desktop/SH-COLLECTOR/'
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
    drive = 'Y:' #manually mapped on Windows to the parent of 'startFolder'
    startFolder = '/Premier Customers/'
    shFolder = '/SAN Health/'
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
    startFolder = '/Test File Repository/'
    shFolder = '/SAN Health/'
    #LOCAL SERVER
    localWorkFolder = 'F:/Users/David/Desktop/SH-COLLECTOR/'
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

#---------------------------------------------------
#---------------------------------------------------

def setPaths(sitePaths):
# #to find the script path:
# server_path = os.path.dirname(os.path.abspath(__file__))
# #normalizes the path to elimiante extra /
# print os.path.normpath(sitePaths[0][3])
# #use os.path.normcase to set to low case

#determine server environment (lab, production, test, quality, etc)
    for site in sitePaths:    
        if os.path.isdir(site[4]): #does localWorkFolder exists?
            #if so, set paths based on this server's environment
            # siteEnv,\
            # drive,\
            # startFolder,\
            # shFolder,\
            # localWorkFolder,\
            # slides_template_path,\
            # csvLogFile,\
            # tmpLogFile,\
            # collectorFolder, csvTempFolder, archiveFolder = site

            #store the paths on options
            #options.sitePaths = location
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

