
#from slLib import *
from sqlLib import *
#import re
#import datetime
from slDeck_single import singleDeck
from slDeck_multi import multiDeck
from shLib import logEntry

#----------------------------------------------------------------------
# NOTES
#1) The reports' name (sanName), and date (shDate)
# are being added by shLib.getCsvData()
# <--sqLib.fill_dbTable() <--sqlib.csv_to_db() <--slDeck.loadDbTables()

#----------------------------------------------------------------------
def loadDbTables(tbl_options):
    '''
    1) Opens a connection to database that will remain open until
    all slide decks for a given customer are created.

    2) extracts the specified columns from the csv files and
    loads them into one db table per file.
    '''
    #!
    #customer, csvPath, shName, sanName, shYear = tbl_options.custData
    customer, csvPath, shName, sanName, shDate, shYear = tbl_options.custData
    #!

    #set desired table options
    #Is this needed? csvPathList is storing those now.
    tbl_options.csvPath = csvPath

    # Open the database
    conn = sql.connect(sqlite_file)
    #------------------------------
    conn.row_factory = sql.Row
    #-------------------------------
    tbl_options.dbConnection = conn
    c = conn.cursor()
#---------------------------------

##Import into dbtable: SwitchSummary.csv
    tbl_options.csvFile = 'SwitchSummary'
    tbl_options.csvColumns = ['a', 'k', 'c', 'f', 'i', 's', 't']

    tbl_options.dbTableName = 'switches'
    tbl_options.dbColNames = '''
    date TEXT,
    san TEXT,
    sw_name TEXT PRIMARY KEY,
    sw_sn TEXT,
    sw_model TEXT,
    sw_firmware TEXT,
    sw_fabric TEXT,
    sw_state TEXT,
    sw_status TEXT
    '''
    csv_to_db(tbl_options)

#-----------------------------
    
##Import into dbtable:  SwitchPortUsage.csv
    tbl_options.csvFile = 'SwitchPortUsage'
    tbl_options.csvColumns = ['a', 'f', 'k', 'l', 'm', 'p', 'q', 'r']

    tbl_options.dbTableName = 'ports'
    tbl_options.dbColNames = '''
    date TEXT,
    san TEXT,
    sw_name TEXT PRIMARY KEY,
    total_ports INT,
    unlic_ports INT,
    unused_ports INT,
    isl_ports INT,
    total_devices INT,
    disks INT,
    hosts INT
    '''
    csv_to_db(tbl_options)

#-----------------------------

##Import into dbtable: SwitchFRU.csv
    tbl_options.csvFile = 'SwitchFRUs'
    tbl_options.csvColumns = ['a', 'e', 'h', 'f', 'g']

    tbl_options.dbTableName = 'frus'
    tbl_options.dbColNames = '''
    date TEXT,
    san TEXT,
    sw_name TEXT,
    fru_type TEXT,
    fru_sn,
    fru_slot TEXT,
    fru_status TEXT
    '''
    csv_to_db(tbl_options)

#-----------------------------
    
##Import into dbtable:  FabricSummary.csv
    tbl_options.csvFile = 'FabricSummary'
    tbl_options.csvColumns = ['a', 'r', 'v', 'z', 'ad', 'ag']

    tbl_options.dbTableName = 'zones'
    tbl_options.dbColNames = '''
    date TEXT,
    san TEXT,
    sw_fabric TEXT,
    active_zoneCfg TEXT,
    hang_alias INT,
    hang_zones INT,
    hang_configs INT,
    zone_dbUsed TEXT
    '''
    csv_to_db(tbl_options)
    
#-----------------------------

##Import into dbtable:  PortErrorCounters.csv
    tbl_options.csvFile = 'PortErrorCounters'
    tbl_options.csvColumns = [
        'a', 'g', 'h', 'i', 'v', 'w', 'x', 'y', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 
        'r', 's', 't', 'u']

    tbl_options.dbTableName = 'PortErrorCnt'
    tbl_options.dbColNames = '''
    date TEXT,
    san TEXT,
    sw_name TEXT,
    slot_port TEXT,
    tx_frames TEXT,
    rx_frames TEXT,
    avPerf INT,
    pkPerf TEXT,
    buffReserved TEXT,
    buffUsed TEXT,
    error_type TEXT,
    error_count INT
    '''
    
    tbl_options.csvPivotCols = [
        'encInFrame',
        'crc',
        'shortFrame',
        'longFrame',
        'badEOF',
        'encOutFrame',
        'c3Discards',
        'linkFailure',
        'synchLost',
        'sigLost',
        'frameReject',
        'frameBusy']

    csv_to_db(tbl_options)
#-----------------------------
#-----------------------------
##Import into dbtable:  PortErrorChanges.csv
    tbl_options.csvFile = 'PortErrorChanges'
    tbl_options.csvColumns = [
        'a', 'g', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 
        'r', 's', 't', 'u']

    tbl_options.dbTableName = 'PortErrorChg'
    tbl_options.dbColNames = '''
    date TEXT,
    san TEXT,
    sw_name TEXT,
    slot_port TEXT,
    tx_frames_chg,
    rx_frames_chg,
    chg_encInFrame TEXT,
    chg_crc TEXT,
    chg_shortFrame TEXT,
    chg_longFrame TEXT,
    chg_badEOF TEXT,
    chg_encOutFrame TEXT,
    chg_c3Discards TEXT,
    chg_linkFailure TEXT,
    chg_synchLost TEXT,
    chg_sigLost TEXT,
    chg_frameBusy TEXT
    '''
    
    csv_to_db(tbl_options)
#-----------------------------

    #save imported data
    conn.commit()

    #leave the database open. It is closed by createSlide deck
    #conn.close()

#----------------------------------------------------------------
    # Now we have all the relevant data from the csv files
    # Ready to make slides
#----------------------------------------------------------------


    
#----------------------------------------------------------------------
#----------------------------------------------------------------------
def createSlideDeck(tbl_options):
    #This function creates a slide deck
    #from all the SAN Reports downloaded
 
    #IF THERE IS NO DATA FOR A SLIDE, PRINT A NOTE ON THE SLIDE: NO DATA!
    #--------------------------------------------------------
    #SLIDE CREATION
    #--------------------------------------------------------
    # 1) loop over the options.sanList tuplets (sanName, csvPath),
    # 2) store current san pointer into tbl_options.custData
    # 3) call slide deck creator (single or multi)

    for san in tbl_options.sanList:
        #retrieve next SAN data
        #!
        #customer, csvPath, shName, sanName, shYear = tbl_options.custData
        customer, csvPath, shName, sanName, shDate, shYear = tbl_options.custData
        #!

        shDate, shName, shFile, sanName, csvPath = san
        #shName: John_Morrison_170726_1640_Maiden_Prod
        #shFile: 7-27-2017_John_Morrison_170726_1640_Maiden_Prod.zip
        
        #and store it for the slide creator function
        #!
        #custData = (customer, csvPath, shName, sanName, shYear)
        custData = (customer, csvPath, shName, sanName, shDate, shYear)
        tbl_options.custData = custData
        #!
        
        #make and save slideDeck
        singleDeck(tbl_options)
        #print 'SAN', sanName
        logEntry('Slides Created', customer, shName)
        
    if len(tbl_options.sanList) > 1:
        # create a deck with the agregated data from all the downloaded reports
        #store multi SAN directive to use the customer name as the file name.
        #!
        sanName = 'ALL'
        #custData = (customer, csvPath, shName, sanName, shYear)
        custData = (customer, csvPath, shName, sanName, shDate, shYear)
        tbl_options.custData = custData
        #!

        multiDeck(tbl_options)
        #print 'SAN', sanName
        logEntry('Slides Created', customer, 'Agregate')
   # END OF SLIDES
   #note: the db connection is opened by loadDbTables
    tbl_options.dbConnection.close()

#/////////////////////////////////////////////////////////////////
#FOR TESTING JUST SLIDE DECK CREATION, EXECUTE THIS SCRIPT
#this uses default csv files already unzipped on the csvTemp directory

if __name__ == "__main__":

#to run this script by itself, first put the csv files for the customer referenced by custData in
#F:\Users\David\Desktop\SH-COLLECTOR\Downloads\csvTemp\
    custData = ('Customer A',
                'F:/Users/David/Desktop/SH-COLLECTOR/Downloads/csvTemp/John_Morrison_170726_1640_Maiden_Prod_',
                'John_Morrison_170726_1640_Maiden_Prod',
                'Maiden_Prod',
                '2017')

    createSlideDeck(custData)

