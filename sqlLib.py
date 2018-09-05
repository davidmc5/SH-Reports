from itertools import islice
import sqlite3 as sql

#NEEDED ONLY FOR EXECUTING JUST THIS FILE
from shLib import getCsvData
#from shPaths import *
#from slLib import Table_Options

#SHOULD THIS ASSIGNEMENT BE IN SHPATH.PY?
#sqlite_file = ':memory'   # No file. Put the database on RAM
sqlite_file = 'shDB.db'    # name of the sqlite database file


def create_dbTable(options):
    ''' Creates a database table named 'tblName' on the open database file 'dbCursor'
        with the columns and type stated in the 'colNames' list of tuples.'''
    dbConn = options.dbConnection
    dbCursor = dbConn.cursor()
    tblName = options.dbTableName
    colNames = options.dbColNames

    # Delete existing table
    dbCursor.execute('DROP TABLE IF EXISTS {tn}'\
            .format(tn=tblName))
    #create new table
    dbCursor.execute('CREATE TABLE {tn} ({cols})'.format(tn=tblName, cols=colNames))



def fill_dbTable(options):

    dbConn = options.dbConnection
    dbCursor = dbConn.cursor()
    tblName = options.dbTableName

    csvData = getCsvData(options)

    #get the number of columns to import from csv file
    #nColumns = len(options.csvColumns)

    #if there is no data, do not create database table
    if len(csvData) == 0:
        return

    
    nColumns = len(csvData[0])
    #set the number of data values' placeholders (?, ?, ....?) based on the number of columns
    qmark = ','.join(['?'] * nColumns)

    #Populate the db table (switch) with the values from the csv file 

##-----------------------------
    #THIS HAS BEEN MODIFIED. Headers are now being removed by shLib\getCsvData()  
        #remove the first record since it is headers: (islice(data, 1, none)
        #this avoids making an expensive copy of the list to remove the header row.

        #inserts all rows from the list at once
##    dbCursor.executemany('INSERT OR IGNORE INTO {tn} VALUES ({q})'\
##                  .format(tn=tblName, q=qmark),islice(getCsvData(options), 1, None))
#-----------------------------
    #insert each row of csv records into database
    dbCursor.executemany('INSERT OR IGNORE INTO {tn} VALUES ({q})'\
                  .format(tn=tblName, q=qmark), csvData)

###-----------------------------
##    #insert each row of csv records into database
##    dbCursor.executemany('INSERT OR IGNORE INTO {tn} VALUES ({q})'\
##                  .format(tn=tblName, q=qmark),getCsvData(options))



def ColCount(db_cursor, table):
    '''Counts the number of columns in the given table'''
    dbCursor.execute("SELECT COUNT (*) FROM {tn}".format(tn=table))
    return len(dbCursor.description[0])


def csv_to_db(options):
    ''' creates a table named dbTableName with the felds dbColNames
        and populates it (fill_dbTable) with the data from the csv file
        '''
    dbConn = options.dbConnection
    dbCursor = dbConn.cursor()
    dbTableName = options.dbTableName
    dbColNames = options.dbColNames
    
    create_dbTable(options)
    fill_dbTable(options)
    dbConn.commit()


def groupHeader(data):
    #Prep data from a SQL Query to remove first column (group by)
    # grab the first data element (group) of each row
    # store it as the first row (with a single value tuple --column) before the group
    # and add to the list all the tuple rows with the same group value
    # but removing the group field (first value of the row)

    newData = []

    #grab the first value for the group
    group = None

    for item in data:
        row = tuple(item)
        if row[0] != group:
            group = row[0]
            #insert the group name as a single row
            newData.append((group,))
            #insert the rest of the row fields without the group field
            newData.append(row[1:])
        else:
            #insert the row fields without the group field 
            newData.append(row[1:])
    return newData






#---------------------------------------------------------------
#---------------------------------------------------------------
#------>>>>>>>>>>> BELOW HERE FOR TEST ONLY!
#---------------------------------------------------------------
#---------------------------------------------------------------
#---------------------------------------------------------------
#---------------------------------------------------------------

if __name__ == "__main__":
    #-------------REMOVE WHEN USING WITH THE COMPLETE SCRIPT
    #initialize table's default options
    options = Table_Options()
    #---------------------------------------------------------------
    #---------------------------------------------------------------


    #THIS NEEDS TO GO INTO slLib.createSlideDeck 
    # Connect to the database file
    conn = sql.connect(sqlite_file)
    c = conn.cursor()
    csv_to_db(c, options)

    c.execute('''
    SELECT
        sw_fabric,
        sw_model,
        COUNT(switches.sw_name),
        SUM(ports.total_ports),
        SUM(ports.unused_ports)
    FROM
        switches
    INNER JOIN ports ON switches.sw_name = ports.sw_name
    GROUP BY
        switches.sw_fabric, switches.sw_model
    ''')

    data = c.fetchall()
    print 'kaka', data


    #THIS NEEDS TO GO INTO slLib.createSlideDeck 
    #close the connection to the database file
    conn.close()







