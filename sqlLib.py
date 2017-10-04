from itertools import islice
import sqlite3 as sql

#NEEDED ONLY FOR EXECUTING JUST THIS FILE
from shLib import getCsvData
#from shPaths import *
#from slLib import Table_Options

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

    #get the number of columns to import from csv file
    nColumns = len(options.csvColumns)
    
    #set the number of data values' placeholders (?, ?, ....?) based on the number of columns
    qmark = ','.join(['?'] * nColumns)

    #Populate the db table (switch) with the values from the csv file 

    #insert each row of csv records into database
    #remove the first record since it is headers: (islice(data, 1, none)
    #this avoids making an expensive copy of the list to remove the header row.

    #inserts all rows from the list at once
    dbCursor.executemany('INSERT OR IGNORE INTO {tn} VALUES ({q})'\
                  .format(tn=tblName, q=qmark),islice(getCsvData(options), 1, None))



def ColCount(db_cursor, table):
    '''Counts the number of columns in the given table'''
    dbCursor.execute("SELECT COUNT (*) FROM {tn}".format(tn=table))
    return len(dbCursor.description[0])


def csv_to_db(options):
    ''' creates a table named dbTableName with the felds dbColNames
        and polulates it (fill_dbTable) with the data from the csv file
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

    for row in data:
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




    ###-----------------------------------------------------
    ###-----------------------------------------------------
    ## SQL TESTING
    ###-----------------------------------------------------
    ###-----------------------------------------------------

    ###-----------------------------------------------------
    ###-----------------------------------------------------
    #THIS NEEDS TO GO INTO slLib.createSlideDeck 
    ###-----------------------------------------------------
    ###-----------------------------------------------------

    #adding records

    ###how many total ports does a given switch has?
    ##c.execute("SELECT SUM(total_ports) FROM ports WHERE sw_name = 'nc-1412-f2-01'")
    ###data = c.fetchall()
    ##data = c.fetchone()
    ##print data[0]


    ###how many switches of a certain model are in fabric 1?
    ##c.execute("SELECT COUNT(sw_name) FROM switches WHERE sw_model = 'DCX-8510-4' and sw_fabric = 'Maiden F1'")
    ##data = c.fetchone()
    ##print data[0]

    ###List (SELECT) each type of switch (GROUP BY) and its number (COUNT)
    ##c.execute('''
    ##SELECT sw_model, COUNT(*)
    ##FROM switches
    ##GROUP BY sw_model
    ##''')
    ##data = c.fetchall()
    ##print 'Number of switches of each model' , data


    ###join two tables (INNER JOINT)
    ###and list (SELECT) the swiches' names and model
    ###for a given model AND fabric (WHERE)
    ##c.execute('''
    ##SELECT switches.sw_name, ports.total_ports
    ##FROM switches
    ##INNER JOIN ports
    ##ON switches.sw_name=ports.sw_name
    ##WHERE sw_model = 'DCX-8510-4' and sw_fabric = 'Maiden F1'
    ##''')

    ### the same but for just a given model (all fabrics)
    ##c.execute('''
    ##SELECT switches.sw_name, total_ports
    ##FROM switches
    ##INNER JOIN ports
    ##ON switches.sw_name = ports.sw_name
    ##WHERE sw_model = '5470'
    ##''')
    ##data = c.fetchall()
    ##print 'switches\'s names and total ports for the given model' , data


    #THIS IS NOT WORKING YET
    #https://www.w3resource.com/sql/aggregate-functions/sum-and-count-using-variable.php

    # the following statement  is using sql alias fields
    #(not real table fields): sCount and pCount
    # count and sum are grouped based on switch model

    ##c.execute('''
    ##SELECT switches.sw_model, switches.sCount, ports.pCount
    ##FROM switches
    ##INNER JOIN (
    ##SELECT sw_model,COUNT(*) AS sCount,
    ##SUM(total_ports) AS pCount
    ##FROM
    ##ON switches.sw_name=ports.sw_name
    ##WHERE sw_model = 'DCX-8510-4' and sw_fabric = 'Maiden F1'
    ##''')

    ##c.execute('''
    ##SELECT
    ##    switches.sw_model,
    ##    COUNT(switches.sw_name),
    ##    SUM(ports.total_ports)
    ##FROM
    ##    switches
    ##INNER JOIN ports ON switches.sw_name = ports.sw_name
    ##GROUP BY
    ##    switches.sw_model
    ##''')

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







