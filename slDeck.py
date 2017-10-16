
from slLib import *
from sqlLib import *
import re
import datetime
from slDeck_single import singleDeck
from slDeck_multi import multiDeck

#------------------------------------------------
def formatDbUsed(data):
    #convert the bytes value into a MB to shorten the width of the column
    newData = []

    #grab the dbUse value
    for row in data:        

        #Example: (row[6] = 11.7% of 1045274B)
        #Example: (row[6] = 11% of 1045274B)
        # extract 1045274B and convert to 1.0MB
        #items = re.match('(\d+\.\d+%)\s+of\s+(\d+)B', row[5])
        
        #1) grab everything before the % sign
        #2) grab everything between 'of' and 'B'
        items = re.match('(.+)%\s+of(.+)B', row[5])

        #print items.group(1)
        #print items.group(2)


        # get the two items in parenthesis (the usage % and the db size in B)
        usage = round( float(items.group(1)), 1)
        # convert Bytes to MB
        mb = float(items.group(2))/1000000

        lst = list(row[:5]) # save the first 4 elements
        # add the last element(db usage) to the tuple
        #(tuples are immutable so first convert to a list 
        lst.append(str(usage) + '%' + '  of ' + str(round(mb, 1)) + ' MB')
        newData.append(tuple(lst))
    return newData


#----------------------------------------------------------------------
#----------------------------------------------------------------------
def loadDbTables(tbl_options):
    '''
    1) Opens a connection to database that will remain open until
    all slide decks for a given customer are created.

    2) extracts the specified columns from the csv files and
    loads them into one db table per file.
    '''
    
    customer, csvPath, shName, sanName, shYear = tbl_options.custData


    #set desired table options
    #Is this needed? csvPathList is storing those now.
    tbl_options.csvPath = csvPath

    # Open the database
    conn = sql.connect(sqlite_file)
    tbl_options.dbConnection = conn
    c = conn.cursor()
#---------------------------------

##Import into dbtable: SwitchSummary.csv
    tbl_options.csvFile = 'SwitchSummary'
    tbl_options.csvColumns = ['a', 'k', 'c', 'f', 'i', 's', 't']

    tbl_options.dbTableName = 'switches'
    tbl_options.dbColNames = '''
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
        customer, csvPath, shName, sanName, shYear = tbl_options.custData
        sanName, csvPath = san
        #and store it for the slide creator function
        custData = (customer, csvPath, shName, sanName, shYear)
        tbl_options.custData = custData
        
        #make and save slideDeck
        singleDeck(tbl_options)
        print 'SAN', sanName
        
    if len(tbl_options.sanList) > 1:
        # create a deck with the agregated data from all the downloaded reports
        #store multi SAN directive to use the customer name as the file name.
        sanName = 'ALL'
        custData = (customer, csvPath, shName, sanName, shYear)
        tbl_options.custData = custData

        multiDeck(tbl_options)
        print 'SAN', sanName

#--------------------------------------------------------------------
   # END OF SLIDES
   #note: the db connection is opened by loadDbTables
    tbl_options.dbConnection.close()

#--------------------------------------------------------------------



def saveDeck(tbl_options):
    ''' Store a remote and local copy of the slide deck
    This is used by both, slDeck_single and slDeck_multi'''
    
    #close the connection to the database file
    #Note: the db may need to remain open if using the RAM file option
    #tbl_options.dbConnection.close()

    #save the slide deck to the customer's SH directory
    #using the current san health file name
    customer, csvPath, shName, sanName, shYear = tbl_options.custData
    folder = drive + startFolder + customer + shFolder
    
    #but if a slide deck with the same name already exists and it is open
    #add a timestamp to the name to make it unique    
    timestamp = datetime.datetime.now().strftime("%y-%m-%d-%H%M")

    prs = tbl_options.presentation

    if sanName == 'ALL':
        shName = customer + '_AGGREGATE_' + timestamp
        
    try:
        prs.save(folder + shName + '.pptx')
    except:
        #if slide deck with the same name is open...
        #... store a new one renamed with a timestamp 
        print folder + shName, 'Kept Open'
        prs.save(folder + shName + '-'+ timestamp + '.pptx')

    #save a LOCAL copy of the slide deck
    folder = archiveFolder
    try:
        prs.save(folder + shName + '.pptx')
    except:
        #if slide deck with the same name is open...
        #... store a new one renamed with a timestamp 
        print folder + shName, 'Kept Open'
        prs.save(folder + shName + '-'+ timestamp + '.pptx')
   
        
#-----------------------------------------------------------------

        
##def formatFruStatus(data):
##    ''' format the fru status to show either none (for blank, ok or enabled)
##        or other'''
##    
##    newData=[]
##    #grab the swtich Status value
##    for row in data:
##
##        status = row[5].lower()
##        if (status == 'ok') or\
##           (status == 'enabled') or\
##           (status == ''):
##            status = 'None'
##
##        lst = list(row[:5]) # save the first 4 elements
##        # add the last element(db usage) to the tuple
##        #(tuples are immutable so first convert to a list 
##        lst.append(str(status))
##        newData.append(tuple(lst))
##    return newData

        
#/////////////////////////////////////////////////////////////////
#FOR TESTING JUST SLIDE DECK CREATION, EXECUTE THIS SCRIPT
#this uses default csv files already unzipped on the csvTemp directory

if __name__ == "__main__":

#to run this script by itself, first put the csv files for the customer referenced by custData in
#F:\Users\David\Desktop\SH-COLLECTOR\Downloads\csvTemp
    custData = ('Customer A',
                'F:/Users/David/Desktop/SH-COLLECTOR/Downloads/csvTemp/John_Morrison_170726_1640_Maiden_Prod_',
                'John_Morrison_170726_1640_Maiden_Prod',
                'Maiden_Prod',
                '2017')

    createSlideDeck(custData)

