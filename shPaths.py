

#---------------------------------------------------
###<------------- PRODUCTION
#---------------------------------------------------
###REMOTE STORAGE
###'Y:/http://connect.brocade.com/cs/Customers/Customer%20Information/'
drive = 'Y:' #manually mapped on Windows to the parent of the 'startFolder'
startFolder = '/Premier Customers/'
shFolder = '/SAN Health/'
#LOCAL SERVER
localWorkFolder = 'C:/users/dmartin/Desktop/SH-Project/'
slides_template_path = localWorkFolder + 'CODE/'
csvLogFile = localWorkFolder + 'shLog.csv'
tmpLogFile = localWorkFolder + 'shLogTemp.csv'
#---common
collectorFolder =  localWorkFolder + 'Downloads/'
csvTempFolder = collectorFolder + 'csvTemp/'
archiveFolder = collectorFolder + 'ARCHIVE/'

prod_path = (
    drive, startFolder, shFolder, 
    localWorkFolder, 
    slides_template_path,
    csvLogFile, tmpLogFile,
    collectorFolder, csvTempFolder, archiveFolder)


#---------------------------------------------------
#<------------- LAB
#---------------------------------------------------
#REMOTE STORAGE
drive = 'Y:' #manually mapped on Windows to the parent folder
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

lab_path = (
    drive, startFolder, shFolder, 
    localWorkFolder, 
    slides_template_path,
    csvLogFile, tmpLogFile,
    collectorFolder, csvTempFolder, archiveFolder)
#---------------------------------------------------
#<------------- COMMON
#---------------------------------------------------
# collectorFolder =  localWorkFolder + 'Downloads/'
# csvTempFolder = collectorFolder + 'csvTemp/'
# archiveFolder = collectorFolder + 'ARCHIVE/'

# zip file inside the SH ZIP that contains the CSV files
csvZipFileName = '_CSVReports.zip'

