from slLib import *
#import slDeck
from sqlLib import *


def multiDeck(tbl_options):
    ''' Creates a slide deck with the agregated data from all
    the reports downloaded'''

    #Open Slide Deck Template
    prs = Presentation(slides_template_path)
    tbl_options.presentation = prs

    #connect to database
    conn = tbl_options.dbConnection
    c = conn.cursor()

    #retrieve this report's variables
    customer, csvPath, shName, sanName, shDate, shYear = tbl_options.custData
#--------------------------------------------------------

#--------------------------------------------------------------------
    #TITLE SLIDE

    #grab first (title) slide (index=0)
    slide = prs.slides[0]
    shapes = slide.shapes

    #add_textbox(left, top, width, height)
    subtxt = shapes.add_textbox(Inches(2),Inches(4), Inches(10), Inches(1))
    subtxt.text = customer
    subtxt.text_frame.paragraphs[0].font.size = Pt(40)
    subtxt.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255) # White?

#--------------------------------------------------------------------
#--------------------------------------------------------------------

    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #SLIDE: SAN Health Combined Report
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    # Slide to show the SH reports used in the agregate.

    #place all the report file names in a list
    #options.sanList =[ (shDate, shName, shFile, sanName, csvPath), (...), ]
    names = []
    for san in tbl_options.sanList:
        #print san[0]
        #!
        # append only the max date file names.
        #!
        names.append(san[1]) # shName
    count = len(names)
    title = "SAN Health Combined Report"
    subtitle = customer + ' - Reports Included (' + str(count) + '):'

    textSlide(prs, title, subtitle, names)

#--------------------------------------------------------------------
#--------------------------------------------------------------------
# added here the two slides from the compared report to unify in a single report
#--------------------------------------------------------------------
#--------------------------------------------------------------------
#--------------------------------------------------------------------

    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #SLIDE: PORT UTILIZATION 2
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    tbl_options.title = 'Port Summary'
    tbl_options.subtitle = '( Change From Previous Period )'
    tbl_options.subtitle_fontSize = Pt(20)

    c.execute('''
    SELECT
        p1.san,
        COUNT(p1.sw_name) AS p1count,
        printf('%d (%+d)', SUM(p1.total_ports), SUM(p1.total_ports)- SUM(p2.total_ports)),
        printf('%d (%+d)', SUM(p1.unused_ports), SUM(p1.unused_ports)- SUM(p2.unused_ports)),
        printf('%d %% (%+.1f)',
            100 * ( 1.0 * SUM(p1.total_ports) - SUM(p1.unused_ports) ) / SUM(p1.total_ports),
            100 * ( 1.0 * SUM(p1.total_ports) - SUM(p1.unused_ports) ) / SUM(p1.total_ports)-
            100 * ( 1.0 * SUM(p2.total_ports) - SUM(p2.unused_ports) ) / SUM(p2.total_ports)
            )
    FROM ports as p1
    INNER JOIN ports as p2
        ON p1.date > p2.date
        AND p1.sw_name = p2.sw_name

    GROUP BY
        p1.san, p1.date

    ORDER BY
        p1.san, p1.date
    --without the limit, the query returns one record for every combination of p1.date > p2.date
    --LIMIT 1

    ''')
    data = c.fetchall()

    #format data with group headers (remove the group = first column data)
    #data = groupHeader(data)

    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list

    if data:
        headers = [('SAN',
                    'Total Switches',
                    'Total Ports',
                    'Free Ports',
                    'Utilization',
                    )]

        headers.extend(data)
        data = headers
        tbl_options.font_size = Pt(10)
        create_single_table_db(data, tbl_options)
        #keep track if there are any non-blank slides in this deck
        #slidesExist +=1

#--------------------------------------------------------------------
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #SLIDE: ZONING SUMMARY (hanging zones)
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Modified to show changes between two dates

    tbl_options.title = 'Zoning Summary'
    tbl_options.subtitle = '( Change From Previous Period )'
    tbl_options.subtitle_fontSize = Pt(20)

    c.execute('''
    SELECT
        z1.san,
        z1.active_zoneCfg,
        printf('%d (%+d)', z1.zones, z1.zones-z2.zones),
        printf('%d (%+d)', z1.hang_zones, z1.hang_zones-z2.hang_zones)
        --z1.zone_dbUsed
    FROM
        zones as z1

    INNER JOIN zones as z2
        ON z1.date > z2.date
        AND z1.principalSw = z2.principalSw
        AND z1.active_zoneCfg != 'N/A'

    GROUP BY
        z1.san, z1.principalSw
    ORDER BY
        z1.san, z1.zones
   ''')
    data = c.fetchall()

    # if len(data) > 0:
    #     slidesExist +=1


    #covert data on 'dbUsed' column from Bytes to MB
    #data = formatDbUsed(data)
    #reformat dbUsed data

    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
    if data:
        headers = [('Fabric',
                    'Active Zone',
                    'Zones',
                    'Hanging Zone Members',
                    )]

        headers.extend(data)
        data = headers
        create_single_table_db(data, tbl_options)

#--------------------------------------------------------------------
#--------------------------------------------------------------------



#--------------------------------------------------------------------
#--------------------------------------------------------------------
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #SLIDE: FABRIC SUMMARY TABLE
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    tbl_options.title = 'Fabric Summary'
    tbl_options.subtitle = customer

    c.execute('''
    SELECT
        sw_fabric,
        sw_model,
        COUNT(s.sw_name) AS count,
        SUM(ports.total_ports),
        SUM(ports.unlic_ports),
        SUM(ports.unused_ports),
        SUM(ports.isl_ports),
        SUM(ports.hosts),
        SUM(ports.disks),
        SUM(ports.total_devices),
        printf('%.0d %', 100 * ( SUM(ports.total_ports) - SUM(ports.unused_ports) ) / SUM(ports.total_ports))
    FROM
        switches s

    INNER JOIN ports
        ON  s.sw_name = ports.sw_name
        AND s.date = ports.date
        -- Use only the max date for the fabric (san)
        AND s.date = (SELECT MAX(date) FROM switches s WHERE s.san = ports.san)
    GROUP BY
        s.sw_fabric, s.sw_model
    ORDER BY
        s.sw_fabric, s.sw_model, count
    ''')

    data = c.fetchall()

    #format data with group headers (remove the group = first column data)
    #data = groupHeader(data)

    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
    headers = [('Fabric',
                'Switch Model',
                'Total Switches',
                'Total Ports',
                'Unlic Ports',
                'Unused Ports',
                'ISL Ports',
                'Hosts',
                'Disks',
                'Total Devices',
                'Ports Used')]

    headers.extend(data)
    data = headers
    tbl_options.font_size = Pt(10)
    create_single_table_db(data, tbl_options)


#--------------------------------------------------------------------
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #SLIDE: ZONING SUMMARY (hanging zones)
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    tbl_options.title = 'Zoning Summary'
    tbl_options.subtitle = customer

    c.execute('''
    SELECT
        sw_fabric,
        active_zoneCfg,
        zones,
        hang_alias,
        hang_zones,
        hang_configs,
        zone_dbUsed
    FROM
        zones z
    WHERE
        active_zoneCfg != 'N/A'
        -- Use only the max date for the fabric (san)
        AND date = (SELECT MAX(date) FROM switches s WHERE s.san = z.san)
    ORDER BY
        sw_fabric
   ''')
    data = c.fetchall()

    #convert data on 'dbUsed' column from Bytes to MB
    #and add the 'MB' units
    data = formatDbUsed(data, column=7)

    #reformat dbUsed data

    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
    headers = [('Fabric',
                'Active Zone',
                'Total Zones',
                'Hanging Alias Mems',
                'Hanging Zone Mems',
                'Hanging Config Mems',
                'Zone Database Use')]

    headers.extend(data)
    data = headers
    create_single_table_db(data, tbl_options)

#--------------------------------------------------------------------

    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    #SLIDE: SWITCH SUMMARY TABLE
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    tbl_options.title = 'Switch Summary'
    tbl_options.subtitle = customer

    # If FRU Status is 'enabled' or 'ok', set to blank (OK)
    c.execute('''
    UPDATE frus
    SET fru_status=''
    WHERE
        UPPER(fru_status) = 'ENABLED'
        OR UPPER(fru_status) = 'OK'
        OR UPPER(fru_status) = 'N/A'
        ''')
    conn.commit()

    #----------------------
    #This query reports the total of unique combinations of
    #fabric, switch model, firmware, switch status and number of defective FRUs
    c.execute('''
    SELECT
        s.sw_fabric,
        s.sw_model,
        COUNT(*) AS cnt,
        s.sw_firmware,
        s.sw_status,
        SUM(f.fru_cnt) as tot_fru
        FROM switches s
        LEFT JOIN (
            SELECT sw_name, COUNT(*) AS fru_cnt
            FROM frus f
            WHERE fru_status != ''
                -- Use only the max date for the fabric (san)
                AND date = (SELECT max(date) FROM switches s WHERE f.san = s.san)
            GROUP BY sw_name
            ) f
        ON f.sw_name = s.sw_name
            -- Use only the max date for the fabric (san)
            WHERE date = (SELECT MAX(date) FROM switches s2 WHERE s2.san = s.san)

    GROUP BY
        s.sw_fabric, s.sw_model, s.sw_firmware, s.sw_status
    ORDER BY
        s.sw_fabric, s.sw_model, cnt
    ''')
    #--------------------------------

    data = c.fetchall()

    #Add column headers to print on the slide table
    # this is a single tuple with the column names
    # as the very first record of the 'data' list

    headers = [('Fabric',
                'Switch Model',
                'Total Switches',
                'Firmware',
                'Switch Status',
                'Faulty FRUs')]

    #add the sql fetched data[] elements to the headers,
    #so that the headers are the first element on the list
    headers.extend(data)
    #rename 'headers' as 'data'
    data = headers
    create_single_table_db(data, tbl_options)

# #--------------------------------------------------------------------

    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     #SLIDE: PORT ERRORS
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    # Shows all the ports with more than 1k errors and avPerf > 10

    tbl_options.title = 'Port Errors'
    tbl_options.subtitle = 'Showing Error Count > 1k and Avg Perf > 10MB'
    tbl_options.subtitle_fontSize = Pt(20)

    c.execute('''
    SELECT san, sw_name, slot_port, avPerf, error_type, error_count
        FROM
            PortErrorCnt p
        WHERE
            error_count > 999
            AND
            avPerf > 10
            --AND date = (SELECT max(date) FROM PortErrorCnt)
            AND date = (SELECT MAX(date) FROM PortErrorCnt p2 WHERE p.san = p2.san)

        ORDER BY
            error_count DESC
       ''')

    data = c.fetchall()

    headers = [('SAN',
                'Switch Name',
                'Slot / Port',
                'Avg Perf (MB)',
                'Error Type',
                'Error Count')]
    #Add table's headers row to data
    data = addHeaders(headers, data)
    if data:
        create_single_table_db(data, tbl_options)
    else:
        #place all the report file names in a list
        #options.sanList =[ (shDate, shName, shFile, sanName, csvPath), (...), ]
        title = "Port Errors"
        subtitle = 'Showing Error Count > 1k and Avg Perf > 10MB'
        result = ["No errors this period!",]

        textSlide(prs, title, subtitle, result, font_size = 40)

    #--------------------------------------------------------
    #--------------------------------------------------------
    #--------------------------------------------------------
    # END OF SLIDES
    #--------------------------------------------------------
    #--------------------------------------------------------
    #--------------------------------------------------------






#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#SQL TESTS
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    #--------------------------------------------------------

    # c.execute('''
    #     SELECT t1.sw_fabric, t1.hang_alias-t2.hang_alias
    #     FROM zones t1
    #     INNER JOIN zones t2
    #     ON t1.principalSw = t2.principalSw
    #     WHERE t1.date = '2017-09-12' AND  t2.date = '2017-10-03'
    # ''')

    c.execute('''
        SELECT t1.sw_fabric, t1.hang_alias-t2.hang_alias
        FROM zones t1
        INNER JOIN zones t2
        ON t1.principalSw = t2.principalSw
        WHERE t1.date = (SELECT max(date) FROM zones) AND  t2.date = '2017-09-12'
    ''')
    data = c.fetchall()
    #print data


# Another test
    # c.execute('''
    #     SELECT
    #         (SELECT max(date) FROM zones),
    #         (
    #         SELECT date
    #         FROM zones
    #         WHERE date < mxd
    #         ORDER BY date DESC
    #         LIMIT 1
    #         )
    #     FROM zones
    # ''')
    # data = c.fetchall()
    # print data

    c.execute('''
        SELECT max(date) FROM zones
    ''')
    data = c.fetchall()
    #print data


# #--------------------------------------------------------------------
    #SQL TESTS
    #prints all instances of values if a column has any letters

    # c.execute('''
    # WITH errors AS (
    #     SELECT sw_name, slot_port, err_c3Discards
    #     FROM PortErrorCnt
    #     WHERE CAST(err_c3Discards as decimal) > 600)
    #     SELECT * FROM errors
    #     ''')
    # c.execute('''
    # SELECT san,
    # FROM PortErrorCnt
    # ''')
    # data = c.fetchall()
    # print data
    #
    #

#--------------------------------------------------------------------
    # END OF SLIDES
    saveDeck(tbl_options)
    logEntry('Slides Created', customer, 'Combined')
#--------------------------------------------------------------------
#--------------------------------------------------------------------
