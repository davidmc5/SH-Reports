

###<------------- BROCADE
###REMOTE
###'Y:/http://connect.brocade.com/cs/Customers/Customer%20Information/'
##drive = 'Y:' #manually mapped on Windows to the parent of the startFolder
##startFolder = '/Premier Customers/'
##shFolder = '/SAN Health/'
###LOCAL
##localWorkFolder = 'C:/users/dmartin/Desktop/SH-Project/'

#<------------- HOME
#REMOTE
drive = 'Y:' #manually mapped on Windows to the parent folder
startFolder = '/Test File Repository/'
shFolder = '/SAN Health/'
#LOCAL
localWorkFolder = 'F:/Users/David/Desktop/SH-COLLECTOR/'

#---------------------------------------------------

#<------------- COMMON
collectorFolder =  localWorkFolder + 'Downloads/'
csvTempFolder = collectorFolder + 'csvTemp/'
archiveFolder = collectorFolder + 'ARCHIVE/'

#log file
logFilePath = localWorkFolder
csvLogFile = logFilePath + '/shLog.csv'
tmpLogFile = logFilePath + '/shLogTemp.csv'

# Slide Deck Template
slides_template_path = localWorkFolder + 'shTemplate.pptx'

# zip file inside the SH ZIP that contains the CSV files
csvZipFileName = '_CSVReports.zip'
