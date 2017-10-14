
from slLib import *
from sqlLib import *
import re
import datetime
from slDeck_single import singleDeck

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

#!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    #THE SLIDE'S CREATION NEED TO BE IN A LOOP TO GENERATE A SEPARATE DECK
    #FROM EACH SH REPORT
    #AND AFTER THAT ONE SINGLE DECK FOR ALL REPORTS COMBINED
    
#--------------------------------------------------------
    #SLIDES
#--------------------------------------------------------

    #make and save slideDeck
    singleDeck(tbl_options)

#--------------------------------------------------------------------
   # END OF SLIDES
   #note: the db connection is opened by loadDbTables
    tbl_options.dbConnection.close()

#--------------------------------------------------------------------



def saveDeck(tbl_options):
    ''' Store a remote and local copy of the slide deck
    This is used by slDeck_single and slDeck_multi'''
    
    #close the connection to the database file
    #Note: the db may need to remain open if using the RAM file option
    #tbl_options.dbConnection.close()

    #save the slide deck to the customer's SH directory
    #using the current san health file name
    customer, csvPath, shName, sanName, shYear = tbl_options.custData
    folder = drive + startFolder + customer + shFolder
    prs = tbl_options.presentation
    
    try:
        prs.save(folder + shName + '.pptx')
    except:
        #if slide deck with the same name is open...
        #... store a new one renamed with a timestamp 
        print folder + shName, 'Kept Open'
        prs.save(folder + shName + '-'+ timestamp + '.pptx')

    #save a local copy of the slide deck
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


##def createSlideDeck(tbl_options):
##    #This function creates a slide deck with the csv files extracted in the collector folder
##    
##    #-----------------------------------
##    # TO DO
##
##    #1) MAKE A FUNCTION TO READ a slide definition config file
##    #AND GENERATE SLIDE DECK
##
##    #2) IF THERE IS NO DATA FOR A SLIDE, PRINT A NOTE ON THE SLIDE: NO DATA!
##    #-----------------------------------
##
##
##    customer, csvPath, shName, sanName, shYear = tbl_options.custData
##
##    #Open Slide Deck Template
##    prs = Presentation(slides_template_path)
##
##    #set desired table options
##    tbl_options.csvPath = csvPath
##    tbl_options.presentation = prs
##
##
##    # Connect to the database file
##    conn = sql.connect(sqlite_file)
##    tbl_options.dbConnection = conn
##    c = conn.cursor()
##
###---------------------------------
##
####Import into dbtable: SwitchSummary.csv
##    tbl_options.csvFile = 'SwitchSummary'
##    tbl_options.csvColumns = ['a', 'k', 'c', 'f', 'i', 's', 't']
##
##    tbl_options.dbTableName = 'switches'
##    tbl_options.dbColNames = '''
##    sw_name TEXT PRIMARY KEY,
##    sw_sn TEXT,
##    sw_model TEXT,
##    sw_firmware TEXT,
##    sw_fabric TEXT,
##    sw_state TEXT,
##    sw_status TEXT
##    '''
##    csv_to_db(tbl_options)
##
###-----------------------------
##    
####Import into dbtable:  SwitchPortUsage.csv
##    tbl_options.csvFile = 'SwitchPortUsage'
##    tbl_options.csvColumns = ['a', 'f', 'k', 'l', 'm', 'p', 'q', 'r']
##
##    tbl_options.dbTableName = 'ports'
##    tbl_options.dbColNames = '''
##    sw_name TEXT PRIMARY KEY,
##    total_ports INT,
##    unlic_ports INT,
##    unused_ports INT,
##    isl_ports INT,
##    total_devices INT,
##    disks INT,
##    hosts INT
##    '''
##    csv_to_db(tbl_options)
##
###-----------------------------
##
####Import into dbtable: SwitchFRU.csv
##    tbl_options.csvFile = 'SwitchFRUs'
##    tbl_options.csvColumns = ['a', 'e', 'h', 'f', 'g']
##
##    tbl_options.dbTableName = 'frus'
##    tbl_options.dbColNames = '''
##    sw_name TEXT,
##    fru_type TEXT,
##    fru_sn,
##    fru_slot TEXT,
##    fru_status TEXT
##    '''
##    csv_to_db(tbl_options)
##
###-----------------------------
##    
####Import into dbtable:  FabricSummary.csv
##    tbl_options.csvFile = 'FabricSummary'
##    tbl_options.csvColumns = ['a', 'r', 'v', 'z', 'ad', 'ag']
##
##    tbl_options.dbTableName = 'zones'
##    tbl_options.dbColNames = '''
##    sw_fabric TEXT,
##    active_zoneCfg TEXT,
##    hang_alias INT,
##    hang_zones INT,
##    hang_configs INT,
##    zone_dbUsed TEXT
##    '''
##    csv_to_db(tbl_options)
###-----------------------------
##    
##
###----------------------------------------------------------------
##    # Now we have all the relevant data from the csv files
##    # Ready to make slides
###----------------------------------------------------------------
##
##
##
###--------------------------------------------------------
###--------------------------------------------------------
###--------------------------------------------------------
###--------------------------------------------------------
##    #SLIDES
###--------------------------------------------------------
###--------------------------------------------------------
##    ###SLIDE COPY
##    ##copy_slide(prs, prs, 0)
##
###--------------------------------------------------------------------
##    ###SLIDE MOVE
##    ##move_slide(prs, 0, 1)
##
###--------------------------------------------------------------------
##    #TITLE SLIDE
##
##    #grab first (title) slide (index=0)
##    slide1 = prs.slides[0]
##    shapes = slide1.shapes
##    
##    #add_textbox(left, top, width, height)
##    subtxt = shapes.add_textbox(Inches(2),Inches(4), Inches(10), Inches(1))
##    subtxt.text = 'SAN: ' + sanName
##    subtxt.text_frame.paragraphs[0].font.size = Pt(40)
##    subtxt.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255) # White?
##    
###--------------------------------------------------------------------
##    ##SLIDE WITH JUST TEXT
##    
##    ##title_slide_layout = prs.slide_layouts[1]
##    ##slide = prs.slides.add_slide(title_slide_layout)
##    ##title = slide.shapes.title
##    ##subtitle = slide.placeholders[1]
##    ##title.text = "Hello, World!"
##    ##subtitle.text = "python-pptx was here!"
###--------------------------------------------------------------------
##
####    #SLIDE WITH ONE TABLE - FROM SINGLE CSV FILE
####    tbl_options.title = 'Fabric Summary'
####    tbl_options.subtitle = 'SAN: '+ sanName
####    tbl_options.csvFile = 'FabricSummary'
####    tbl_options.csvColumns = ['a', 'c', 'd', 'e', 'f', 'g', 'k', 'l', 'm']
####    create_single_table(tbl_options)
##
###--------------------------------------------------------------------
##
##    #SLIDE: FABRIC SUMMARY TABLE
##
##    #FROM SQL QUERY WITH GROUP HEADERS ON MERGED CELLS
##    #using two db tables from two csv files
##
##    tbl_options.title = 'Fabric Summary'
##    tbl_options.subtitle = 'SAN: '+ sanName
##
##
##    c.execute('''
##    SELECT
##        sw_fabric,
##        sw_model,
##        COUNT(s.sw_name) AS count,
##        SUM(ports.total_ports),
##        SUM(ports.unlic_ports),
##        SUM(ports.unused_ports),
##        SUM(ports.isl_ports),
##        SUM(ports.hosts),
##        SUM(ports.disks),
##        SUM(ports.total_devices)
##    FROM
##        switches s
##    INNER JOIN ports ON s.sw_name = ports.sw_name
##    GROUP BY
##        s.sw_fabric, s.sw_model
##    ORDER BY
##        s.sw_fabric, s.sw_model, count
##    ''')
##
##    data = c.fetchall()
##    #format data with group headers (remove the group = first column data)
##    data = groupHeader(data)
##
##    
##    #Add column headers to print on the slide table
##    # this is a tuple with the column names
##    # as the very first record of the 'data' list
##    headers = [('Switch Model',
##                'Total Switches',
##                'Total Ports',
##                'Unlicensed Ports',
##                'Unused Ports',
##                'ISL Ports',
##                'Hosts',
##                'Disks',
##                'Total Devices')]
##
##    headers.extend(data)
##    data = headers
##    create_single_table_db(data, tbl_options)
##
##
###--------------------------------------------------------------------
##
##    #SLIDE: HANGING ZONES TABLE
##
##    tbl_options.title = 'Zoning Summary'
##    tbl_options.subtitle = 'SAN: '+ sanName
##
##    
##    c.execute('''
##    SELECT
##        sw_fabric,
##        active_zoneCfg,
##        hang_alias,
##        hang_zones,
##        hang_configs,
##        zone_dbUsed
##    FROM
##        zones
##    WHERE
##        zone_dbUsed != 'Unknown'
##     ORDER BY
##        sw_fabric
##   ''')
##    data = c.fetchall()
##    
##    #covert data on 'dbUsed' column from Bytes to MB
##    data = formatDbUsed(data)
##
##    #reformat dbUsed data
##    
##    #Add column headers to print on the slide table
##    # this is a tuple with the column names
##    # as the very first record of the 'data' list
##    headers = [('Fabric',
##                'Active Zone',
##                'Hanging Alias Mems',
##                'Hanging Zone Mems',
##                'Hanging Config Mems',
##                'Zone Database Use')]
##
##    headers.extend(data)
##    data = headers
##    create_single_table_db(data, tbl_options)
##
###--------------------------------------------------------------------
##  
##    #SLIDE: SWITCH SUMMARY TABLE
##
##    #FROM SQL QUERY WITH GROUP HEADERS ON MERGED CELLS
##    #using two db tables from two csv files
##
##    tbl_options.title = 'Switch Summary'
##    tbl_options.subtitle = 'SAN: '+ sanName
##
##
####    #Change all fru.fru_status column data to upper case
####      # not needed. Case change done on sql with the UPPER()
####    c.execute('''
####    UPDATE frus
####    SET fru_status = UPPER(fru_status)''')
##    
##    # If FRU Status is 'enabled' or 'ok', set to blank (OK)
##    c.execute('''
##    UPDATE frus
##    SET fru_status=''
##    WHERE UPPER(fru_status) = 'ENABLED' OR UPPER(fru_status) = 'OK' ''')
##    conn.commit()
##
##    
###----------------------
##    #This query reports the total of unique combinations of
##    #fabric, switch model, firmware, switch status and number of defective FRUs
##    c.execute('''
##    SELECT
##        s.sw_fabric,
##        s.sw_model,
##        COUNT(*) AS cnt,
##        s.sw_firmware,
##        s.sw_status,
##        SUM(f.fru_cnt) as tot_fru
##        FROM switches s
##        LEFT JOIN (
##            SELECT sw_name, COUNT(*) AS fru_cnt
##            FROM frus
##            WHERE fru_status != ''
##            GROUP BY sw_name
##            ) f
##        ON f.sw_name = s.sw_name
##    GROUP BY
##        s.sw_fabric, s.sw_model, s.sw_firmware, s.sw_status
##    ORDER BY
##        s.sw_fabric, s.sw_model, cnt
##    ''')
###--------------------------------
###--------------------------------
##
####    # this prints db column headers for the query, if used right after a query
####    for k in c.description:
####        print(k[0])
##
####    # format FRU status
####    # no longer needed since the formating is done with an SQL UPDATE command
####    #data = formatFruStatus(data)
##
##    #grab the results of the sql query
##    data = c.fetchall()
##    
##    #format data with group headers (remove the group = first column data)
##    data = groupHeader(data)
##
##    
##    #Add column headers to print on the slide table
##    # this is a tuple with the column names
##    # as the very first record of the 'data' list
##    # first column for 'fabric' will be printed on a single dividing row
##    
##    headers = [('Switch Model',
##                'Total Switches',
##                'Firmware',
##                'Switch Status',
##                'Faulty FRUs')]
##
##    headers.extend(data)
##    data = headers
##    create_single_table_db(data, tbl_options)
##   
##
##
##
###--------------------------------------------------------------------
###--------------------------------------------------------------------
##    # END OF SLIDES
###--------------------------------------------------------------------
###--------------------------------------------------------------------
##    #close the connection to the database file
##    conn.close()
##
##    #save the slide deck to the customer's SH directory
##    #using the current san health file name
##    folder = drive + startFolder + customer + shFolder
##    try:
##        prs.save(folder + shName + '.pptx')
##    except:
##        #if slide deck with the same name is open...
##        #... store a new one renamed with a timestamp 
##        print folder + shName, 'Kept Open'
##        prs.save(folder + shName + '-'+ timestamp + '.pptx')
##
##    #save a local copy of the slide deck
##    folder = archiveFolder
##    try:
##        prs.save(folder + shName + '.pptx')
##    except:
##        #if slide deck with the same name is open...
##        #... store a new one renamed with a timestamp 
##        print folder + shName, 'Kept Open'
##        prs.save(folder + shName + '-'+ timestamp + '.pptx')
##        


        
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

