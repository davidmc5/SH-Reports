#Module of functions to extract data from csv files
# and produce powerpoint slides 

# this is for this module to access the namespace of the main script.
#but it did not work for my case (createSlideDeck)
#from __main__ import *

#csv data extraction
import string
import csv
import shutil
import errno
import time
import re
import os
import zipfile
from os.path import join

import ast
import copy
import six

# for all folder names and locations see shPaths
from shPaths import *


#Slide layouts
#0 Title Slide
#1 Title and Content
#2 Section Header
#3 Two Content
#4 Comparison
#5 Title Only
#6 Blank <--------------------- LAYOUT #7
#7 Content with Caption
#8 Picture with Caption

#define the index of blank slide template in the pptx slide's master
blank_slide_layout_index = 7


def logEntry(*logData):
    ''' Used to create a timestamped log entry into a csv log file'''
    
    tTime = time.strftime("%X")
    tDate = time.strftime("%x")
    logFields = [tDate, tTime]
    for arg in logData:
        logFields.append(arg)
        
    #logFields.extend(logData.split(','))
    
    try:
        with open(csvLogFile, 'a') as logFile:
            logWriter = csv.writer(logFile)
            logWriter.writerow(logFields)
    except:
        #main log file was left open. Try the temp log file
        try:
            with open(tmpLogFile, 'a') as logFile:
                logWriter = csv.writer(logFile)
                logWriter.writerow(logFields)
        except:
            print "BOTH LOG FILES ARE OPEN! UNABLE TO SAVE LOGs"

    for field in logFields:
        print field,
    print ''

    


def createCsvTempFolder():
    #Create a temp directory for csv zip and csv files
    try:
        os.makedirs(csvTempFolder)
    except OSError as exception:
        #a race condition occurred. 
        if exception.errno != errno.EEXIST:
            #a race condition might have occurred.
            print 'Problem trying to create folder:', csvTempFolder
            raise


def initFolders():
    #Remove temp csv directory if it exists
    #there is a race condition...
    #...re-creating the directory right after deleting it.
    #INTERNALLY shutil.rmtree calls the windows API function DeleteFile
    #The DeleteFile function marks a file for deletion on close.
    #Therefore, the file deletion does not occur until the last handle to the file is closed.
    #Subsequent calls to CreateFile to open the file fail with ERROR_ACCESS_DENIED.
    shutil.rmtree(csvTempFolder, ignore_errors=True)

    #delete any old zip files from collector folder
    for item in os.listdir(collectorFolder):
        if item.endswith(".zip"):
            os.remove(join(collectorFolder, item))



def archiveFiles(shFile, custData, arch_opt=None):
    #Archive remote and local copies of sh zip files

    #retrieve SH report variables
    customer, csvPath, shName, sanName, shYear = custData

    #remote folder
    folder = drive + startFolder + customer + shFolder

    #if option 'no_remote' has been included, do not archive the remote file.
    

    if arch_opt != 'no_remote':
        #Verify it exists or create a directory to archive the remote SH Files
        #Use the year from the file name as the archive directory
        try:
            os.makedirs(folder + shYear)
        except OSError as exception:
            if exception.errno != errno.EEXIST:
                print 'problem trying to create folder', folder + shYear 
                raise
        #Move sh zip file to the archive (year) directory to avoid re-processing
        startFile = folder + shFile
        endFile = folder + shYear + '/' + shFile
        shutil.move(startFile, endFile)
    else:
        #delete remote SH file (Do not archive)
        startFile = folder + shFile
        try:
            os.remove(startFile)
        except:
            #it does not seem this is necessary.
            #The zip file gets deleted even if one of its sub-files is open
            print tDate, tTime, 'Error: Can\'t delete Remote SH File'


    #Local zip file
    #Verify it exists or create a directory to archive the local copy of SH Files
    try:
        os.makedirs(archiveFolder)
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            print 'problem trying to create local archive folder', archiveFolder
            raise
    #Move sh zip file to the local archive directory 
    src = collectorFolder + shFile
    dst = archiveFolder + shFile
    shutil.move(src, dst)


def archiveBad(customer, shFile):
    #Archive bad zip to 'SAN HEALTH/TEMP'
    #archive local copy to archive

#   #remote folder
    folder = drive + startFolder + customer + shFolder
    
    #Verify it exists or create a directory to archive the remote SH Files
    #Use the year from the file name as the archive directory
    try:
        os.makedirs(folder + 'TEMP')
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            print 'problem trying to create folder', folder + 'TEMP' 
            raise
        
    #Move bad sh zip to SAN HEALTH/TEMP 
    startFile = folder + shFile
    endFile = folder + 'TEMP/' + shFile
    shutil.move(startFile, endFile)


    #local zip file
    #Verify it exists or create a directory to archive the local copy of SH Files
    try:
        os.makedirs(archiveFolder)
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            print 'problem trying to create folder', archiveFolder
            raise
    #Move local sh zip file to the archive
    src = collectorFolder + shFile
    dst = archiveFolder + shFile
    shutil.move(src, dst)
    

def getZipFiles(customer):
    #checks the given customer folder for SH ZIP files
    #returns None if no files found or
    #returns the number of files downloaded to the local collector folder
    
    #Customer's SH folder
    folder = drive + startFolder + customer + shFolder

    if not os.path.isdir(folder):
        #shFolder does not exist
        folder = None
        return None
    
    try:
        files = os.listdir(folder)
        #listdir returns an empty list if there are no files or sub folders
        firstCount = len(files)
        if firstCount == 0:
            #no files (of any type, or folders) in this directory
            return None
        else:
            zipCount = 0
            for eachFile in files:
                extension = os.path.splitext(eachFile)[1]
                if (extension == '.zip'):
                    zipCount += 1
            if zipCount == 0:
                return None


        #Users can be uploading multiple files in a single customer folder.
        # Wait until the number of files in the folder is not incrementing,
        # to avoid downloading an incomplete number of files,
        # or retrieving corrupt files that were not fully uploaded

        while True:
            #this timer needs to be higher than max time it takes
            #to upload the largest file!
            time.sleep(10) 
            lastCount = len(os.listdir(folder))
            #print lastCount,
            if lastCount != firstCount:
                #the number of files is still incrementing
                #Keep waiting
                firstCount = lastCount
            else:
                #The number of files uploaded is stable. Retrieve them.
                break


        #there are some files...
        #download just the zip files (if any)
        files = os.listdir(folder)
        fileCount = 0 #counts how many zip files have been collected
        for eachFile in files:
            extension = os.path.splitext(eachFile)[1]
            if (extension == '.zip'):
                createCsvTempFolder()
                fileCount += 1
                src = folder + eachFile
                dst = collectorFolder + eachFile
                
                #this is a patch to prevent downloading a file
                #that is not yet ready, and get it corrupt.
                #find a better way of doing this.
                #time.sleep(1)
                
                shutil.copyfile(src, dst)
    except:
        print 'PROBLEM RETRIEVING ZIP FILES FROM!', folder
        #Probably the file is still being uploaded by user
        #skip until the next pass
        return None
        #raise

#---------------------------
    # FOR TESTING ONLY -- TO DELETE
    #print 'Total ZIP:', fileCount
#---------------------------

    if fileCount == 0:
        #no SH ZIP files from this customer
        #check next customer
        return None
    return fileCount




def get_csvFileItems(csvZipFile):

    #remove the 15 characters from the end string (_CSVReports.zip)...
    #...to get the San Health report name
    #example: John_Morrison_170726_1640_Maiden_Prod_CSVReports.zip
    shName = csvZipFile[:-15]

    #extract report Date, Year and and SAN Name from the csv-zip file name 
    #example: shName = 'John_Morrison_170726_1640_Maiden_Prod'
    items = re.match('.+_(\d{6})_\d{4}_(.+)', shName)
    shDate = items.group(1)
    shYear = '20' + shDate[:2] #used to create SH report archive directory
    sanName = items.group(2) #used for slides subtitles
    
    #get full path common to all csv files
    #(only the csv file name ending is different)
    csvPath = csvTempFolder + shName + '_'
    return (csvPath, shName, sanName, shYear)

def get_shFiles():
    #generator to get the names of ONLY files (not direstories)...
    #...in the collector folder
    files = ( shFile for shFile in os.listdir(collectorFolder) 
         if os.path.isfile(os.path.join(collectorFolder, shFile)) )
    return files
    


def extract_csvFiles(shFile):

    #this extracts from the given SH zip file in the local collector
    #the csv files compressed inside (if any)
    #and places them in the same directory

    if shFile.endswith('.zip'):

        #shFile is the file name (without the path)...
        #...of the current SH zip report in process

        #print 'shFile', shFile


        #open a handle for the SH ZIP file to view its contents
        #with zipfile.ZipFile(collectorFolder + shFile, 'r') as file_zip:

        try:
            with zipfile.ZipFile(collectorFolder + shFile, 'r') as file_zip:

                #check each file in the SH ZIP for the CSV folder
                for compFile in file_zip.namelist():
                    
                    #determine if it is CSV ZIP file ends with the string
                    #referenced by csvZipFileName
                    if compFile.endswith(csvZipFileName):
                        #print compFile
                      
                        #extract the CSV ZIP file from the SH ZIP
                        #...into the local csv collector folder
                        file_zip.extract(compFile, csvTempFolder)
                        
                        #Open CSV ZIP file and extract all the .csv files
                        #...into the same folder
                        with zipfile.ZipFile(csvTempFolder + compFile, 'r') as csvZipFile:
                            for csvFile in csvZipFile.namelist():
                                #Extract all files to temp folder
                                csvZipFile.extract(csvFile, csvTempFolder)
                            csvZipFile.close()
                        file_zip.close()
                        #now we have in the local collector folder all the individual csv files
                        return get_csvFileItems(compFile)
        except:
            #print 'zip file appears to be corrupt'
            #rename
            #shFile = 'BAD_' + shFile
            return 'Bad'
            #raise
    else:
        return None

    

def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num




def getCsvData(options):
    #getCsvData extracts the specified columns from the csv file
    #the columns are a list of strings with the csv column letter identifier:
    #example: columns = ['a', 'c', 'd', 'l', 'm']
    #it returns a list of row-lists (one list per row) in that column order

    csvFile = options.csvPath + options.csvFile + '.csv'
    columns = options.csvColumns
    shData = [] #initialise a list to store each row-list
    
    with open(csvFile, 'rb') as f:
        reader = csv.reader(f)

        for shRow in reader:
            if len(shRow) == 0: break
            row=[]
            for col in columns:
                row.append(shRow[col2num(col)-1])
            #print row
            shData.append(row)
    return shData


#================================================================
# For testing, exectuting just this file
if __name__ == "__main__":

    # TO RUN JUST THIS SCRIPT, PUT ZIP FILES ON THE REMOTE REPOSITORY
    # THIS WILL DOWNLOAD THEM INTO THE LOCAL DOWNLOAD REPOSITORY
    while True:
        customer = 'Customer B'
        getZipFiles(customer)
        #print '.',




