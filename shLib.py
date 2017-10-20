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

import inspect

# for all folder names and locations see shPaths
from shPaths import *
#import shPaths

from itertools import islice

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
    caller = inspect.stack()[1][1] # get the module that made this call
    line = inspect.stack()[1][2] # get the line that made this call
    logFields = [tDate, tTime, caller, line]
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
            print "HELLO!!! BOTH LOG FILES ARE OPEN! UNABLE TO SAVE LOGs!"

    # Print Log Message also to the Command Prompt Screen
    print tDate, tTime, 
    for field in logData:
        print field,
    print ''

    


def createCsvTempFolder():
    #Create a temp directory for csv zip and csv files

    while True:
        try:
            os.makedirs(csvTempFolder)
            return
        except OSError as exception:
            #a race condition occurred. 
            if exception.errno != errno.EEXIST:
                #a race condition might have occurred.
                #print 'Problem trying to create folder:', csvTempFolder
                logEntry('Folder Creation Error', csvTempFolder)
                time.sleep(10)
                continue
            else:
                return



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
    try:
        for item in os.listdir(collectorFolder):
            if item.endswith(".zip"):
                os.remove(join(collectorFolder, item))
    except:
        logEntry('File Move Error', collectorFolder, item, 'Waiting 30 sec')
        #wait for the OS to close the files        
        time.sleep(30)



#def archiveFiles(shFile, custData, archv_opt=None):
def archiveFiles(options):
    #Archive remote and local copies of sh zip files

    #retrieve SH report variables
    customer, csvPath, shName, sanName, shYear = options.custData
    
    #remote folder
    folder = drive + startFolder + customer + shFolder

    #if option 'no_remote' has been included, do not archive the remote file.
    archv_opt = options.archv_opt
    
    for san in options.sanList:
        shName, shFile, sanName, csvPath = san
        #print 'ARCHIVING', shFile

        if archv_opt != 'no_remote':
            #Verify if the remote archive directory already exists
            #or create a new one to archive the processed remote SH Files
            #Use the year from the file name as the archive directory
            try:
                os.makedirs(folder + shYear)
            except OSError as exception:
                if exception.errno != errno.EEXIST:
                    #print 'problem trying to create folder', folder + shYear
                    logEntry('Folder Creation Error', folder + shYear)
                    
            #Move sh zip file to the archive (year) directory to avoid re-processing
            startFile = folder + shFile
            endFile = folder + shYear + '/' + shFile
            shutil.move(startFile, endFile)
        else:
            #Option "no_remote" given:
            #delete remote SH file (Do not archive)
            startFile = folder + shFile
            try:
                os.remove(startFile)
            except:
                #This does not seem this is necessary.
                #The zip file gets deleted even if one of its sub-files is open
                #print tDate, tTime, 'Error: Can\'t delete Remote SH File'
                logEntry('Folder Deletion Error', 'Can\'t delete Remote SH File', folder + shFile)

        #Local zip file
        #Verify it exists or create a directory to archive the local copy of SH Files
        try:
            os.makedirs(archiveFolder)
        except OSError as exception:
            if exception.errno != errno.EEXIST:
                logEntry('Folder Creation Error', 'Can\'t create local archive folder', archiveFolder)
                time.sleep(60)
                #raise
        #Move sh zip file to the local archive directory 
        src = collectorFolder + shFile
        dst = archiveFolder + shFile
        shutil.move(src, dst)


def archiveBad(customer, shFile):
    #Archive bad zip to 'SAN HEALTH/TEMP'
    #archive local copy to archive

    #remote folder
    #drive and startFolder are defined in shPath.py
    folder = drive + startFolder + customer + shFolder
    
    #Verify it exists or create a directory to archive the remote bad SH Files
    try:
        os.makedirs(folder + 'TEMP')
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            logEntry('Can\'t Create TEMP folder in ', folder)
        
        
    #Move bad sh zip to SAN HEALTH/TEMP
        try:            
            startFile = folder + shFile
            endFile = folder + 'TEMP/' + shFile
            shutil.move(startFile, endFile)
        except:
            logEntry('File Move Error', startFile, 'Waiting 30 sec')
            time.sleep(30)
            #Problem moving remote bad file to archive'
##            if startFile not in blacklist:
##                blacklist.append(startFile)
##                logEntry('Can\'t Move File', startFile)
##                print blacklist

#--------------------------------------------
#--------------------------------------------
    #print 'BAD FILE:', startFile

#--------------------------------------------
#--------------------------------------------

    #local zip file
    #Verify it exists or create a directory to archive the local copy of SH Files
    try:
        os.makedirs(archiveFolder)
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            #print 'Problem trying to create folder', archiveFolder
            logEntry('Can\'t Create Archive folder', archiveFolder)
            
        
    #Move local sh zip file to the archive
    try:        
        src = collectorFolder + shFile
        dst = archiveFolder + shFile
        shutil.move(src, dst)
    except:
        #print 'Problem moving local bad file to archive', src
        logEntry('Can\'t Archive Bad Zip', src)
        
    

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
                logEntry('File Upload', 'Waiting for mutiple files being uploaded')                
            else:                
                #The number of files uploaded is stable. Retrieve them.
                #print lastCount
                break


        #there are some files...
        #download just the zip files (if any)
        files = os.listdir(folder)
        #print files
        fileCount = 0 #counts how many zip files have been collected
        
        for eachFile in files:
            #print eachFile
            
            extension = os.path.splitext(eachFile)[1]
            if (extension == '.zip'):

                #----problem here
                createCsvTempFolder()
                #----problem here
                
                fileCount += 1
                #print 'fileCount', fileCount
                src = folder + eachFile
                dst = collectorFolder + eachFile
                
                #this is a patch to prevent downloading a file
                #that is not yet ready, and get it corrupt.
                #find a better way of doing this.
                #time.sleep(1)
                shutil.copyfile(src, dst)
               
            #print eachFile
    except:
    
        #PROBLEM RETRIEVING ZIP FILES!
        logEntry('File Error', 'Can\'t Retrieve Files From', folder)
        
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



def get_csvFiles():
    #generator to get the names of ONLY files (not direstories)...
    #...in the csv collector folder
    files = ( csvFile for csvFile in os.listdir(csvTempFolder) 
         if os.path.isfile(os.path.join(csvTempFolder, csvFile)) )
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


#----------------------------------------
# modified this function to get ALL the csv files extracted
# to the csvTemp directory, and adding the SAN name

def getCsvData(options):
    #getCsvData extracts the specified columns from the csv file
    #the columns are a list of strings with the csv column letter identifier:
    #example: columns = ['a', 'c', 'd', 'l', 'm']
    #it returns a list of row-lists (one list per row) in that column order


    #csvFile = options.csvPath + options.csvFile + '.csv'
    columns = options.csvColumns
    shData = [] #initialise a list to store each row-list

    csvTarget = options.csvFile + '.csv'

    for san in options.sanList:
        shName, shFile, sanName, csvPath = san
        csvFile = csvPath + csvTarget
      
        with open(csvFile, 'rb') as f:
            reader = csv.reader(f)

            for rowIdx, shRow in enumerate(reader):
                #stop reading when we reach the end
                if len(shRow) == 0: break

                if rowIdx == 0:
                    # don't import the header row
                    continue
                
                
                #append each row data to a list
                row=[]
                
                row.append(sanName)
                
                #grab only the columns requested
                for col in columns:
                    row.append(shRow[col2num(col)-1]) #column index start at 1

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
        print '.',




