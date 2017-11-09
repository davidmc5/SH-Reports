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
    
    ###SLIDE COPY
    
    ##copy_slide(prs, prs, 0)

#--------------------------------------------------------------------
    
    ###SLIDE MOVE
    
    ##move_slide(prs, 0, 1)

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
    
    ##SLIDE: SAN Health Combined Report
    # Slide to show the SH reports used in the agregate.
    
    #place all the report file names in a list
    #options.sanList =[ (shDate, shName, shFile, sanName, csvPath), (...), ]
    names = []
    for san in tbl_options.sanList:
        names.append(san[1]) # shName
    count = len(names)
    title = "SAN Health Combined Report"
    subtitle = customer + ' - Reports Included (' + str(count) + '):'  
      
    textSlide(prs, title, subtitle, names)
                
#--------------------------------------------------------------------

    #SLIDE: FABRIC SUMMARY TABLE

    #FROM SQL QUERY WITH GROUP HEADERS ON MERGED CELLS
    #using two db tables from two csv files

    tbl_options.title = 'Fabric Summary'
    tbl_options.subtitle = customer

#THIS WORKS -- ORIGINAL
    # c.execute('''
    # SELECT
    #     sw_fabric,
    #     sw_model,
    #     COUNT(s.sw_name) AS count,
    #     SUM(ports.total_ports),
    #     SUM(ports.unlic_ports),
    #     SUM(ports.unused_ports),
    #     SUM(ports.isl_ports),
    #     SUM(ports.hosts),
    #     SUM(ports.disks),
    #     SUM(ports.total_devices)
    # FROM
    #     switches s
    # 
    # INNER JOIN 
    #     ports 
    #   ON s.sw_name = ports.sw_name
    # GROUP BY
    #     s.sw_fabric, s.sw_model
    # ORDER BY
    #     s.sw_fabric, s.sw_model, count
    # ''')
    #---------------------------------------------------------
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
    
    INNER JOIN 
        ports 
      ON s.sw_name = ports.sw_name
    GROUP BY
        s.sw_fabric, s.sw_model
    ORDER BY
        s.sw_fabric, s.sw_model, count
    ''')
        
#---------------------------------------------------------
    # c.execute('''
    # SELECT
    #     s.sw_fabric,
    #     s.sw_model,
    #     s.count,
    #     s.tot_p.t_ports,
    #     SUM(p.unlic_ports),
    #     SUM(p.unused_ports) as unused,
    #     SUM(p.isl_ports),
    #     SUM(p.hosts),
    #     SUM(p.disks),
    #     SUM(p.total_devices)
    # 
    # 
# FROM (SELECT
#         sw_fabric,
#         sw_name,
#         sw_model,
#         COUNT(sw_name) AS count,
#         (SELECT san, SUM(total_ports) as t_ports FROM ports GROUP BY san) tot_p
#     FROM switches
#     group by sw_fabric, sw_model
#         ) s
# 
#     INNER JOIN 
#         ports p
#     ON s.sw_name = p.sw_name
# 
#     GROUP BY
#         s.sw_fabric, s.sw_model
#     ORDER BY
#         s.sw_fabric, s.sw_model, s.count
#     ''')

    data = c.fetchall()
    print data
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

    #SLIDE: HANGING ZONES TABLE

    tbl_options.title = 'Zoning Summary'
    tbl_options.subtitle = customer

    
    c.execute('''
    SELECT
        sw_fabric,
        active_zoneCfg,
        hang_alias,
        hang_zones,
        hang_configs,
        zone_dbUsed
    FROM
        zones
    WHERE
        active_zoneCfg != 'N/A' and date = (SELECT max(date) FROM zones)
    ORDER BY
        sw_fabric
   ''')
    data = c.fetchall()
    
    #covert data on 'dbUsed' column from Bytes to MB
    data = formatDbUsed(data)

    #reformat dbUsed data
    
    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
    headers = [('Fabric',
                'Active Zone',
                'Hanging Alias Mems',
                'Hanging Zone Mems',
                'Hanging Config Mems',
                'Zone Database Use')]

    headers.extend(data)
    data = headers
    create_single_table_db(data, tbl_options)

#--------------------------------------------------------------------
  
    #SLIDE: SWITCH SUMMARY TABLE

    #FROM SQL QUERY WITH GROUP HEADERS ON MERGED CELLS
    #using two db tables from two csv files

    tbl_options.title = 'Switch Summary'
    tbl_options.subtitle = customer

    # If FRU Status is 'enabled' or 'ok', set to blank (OK)
    c.execute('''
    UPDATE frus
    SET fru_status=''
    WHERE UPPER(fru_status) = 'ENABLED' OR UPPER(fru_status) = 'OK' ''')
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
            FROM frus
            WHERE 
                fru_status != ''
            GROUP BY sw_name
            ) f
        ON f.sw_name = s.sw_name
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

    #add the data[] elements to the headers,
    #so that the headers are the first element 
    headers.extend(data)
    #rename 'headers' as 'data'
    data = headers
    create_single_table_db(data, tbl_options)
   
# #--------------------------------------------------------------------
# 
#     #SLIDE: PORT ERRORS
# 
    tbl_options.title = 'Port Errors'
    tbl_options.subtitle = customer

   #  
   #  c.execute('''
   #  SELECT
   #      san,
   #      sw_name,
   #      COUNT(*)
   #  FROM
   #      PortErrorCnt
   #  WHERE
   #      avPerf > 100 AND err_c3Discards > 100
   # 
   #  GROUP BY
   #      san, sw_name, slot_port
   #  ORDER BY
   #      san
   # ''')
   #  data = c.fetchall()
   #  
    #--------------------------------------------------------
    #table headers
    # headers = [('SAN',
    #             'Switch Name',
    #             'Num Errors')]
    # #Add table's headers row to data
    # data = addHeaders(headers, data)
    # #--------------------------------------------------------
    # if data:
    #     create_single_table_db(data, tbl_options)

# 
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
#WHERE err_c3Discards GLOB '*[A-Za-z]*' OR CAST(err_c3Discards as decimal) > 100
#WHERE CAST(err_c3Discards as decimal) > 800
#------------------------------------------------------
    # c.execute('''
    #  SELECT
    #      san,
    #      sw_name,
    #      slot_port,
    #      COUNT(*)
    #  FROM
    #      PortErrorCnt
    #  WHERE
    #     err_c3Discards != 0
    #     avPerf > 100 AND err_c3Discards > 100
    #  GROUP BY
    #      san, sw_name, slot_port
    #      ''')
    # data = c.fetchall()
#--------------------------------------------------------------------
# #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#     # SQL TESTS
#     # prints all instances of values if a column has any letters
#     
#     c.execute('''
#     SELECT san, sw_name, slot_port
#         FROM
#             PortErrorCnt
#         WHERE
#             CAST(err_c3Discards as decimal) > 700
#         GROUP BY
#             san, sw_name, slot_port
#         ORDER BY
#             san
#        ''')
#        
#        
#     #This moves the cursor c so the next fetchall returns nothing!
#     # for row in c:
#     #     print row.keys()
# 
#        
#     data = c.fetchall()
#        
#     headers = [('SAN',
#                 'Switch Name',
#                 'Slot / Port',
#                 'Errors')]
#     #Add table's headers row to data
#     data = addHeaders(headers, data)
#     if data:
#         create_single_table_db(data, tbl_options)
# 
# #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    #--------------------------------------------------------
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    # Port Error Slide
    # Shows all the ports with more than 1k errors and avPerf > 0
    
    tbl_options.title = 'Port Errors'
    tbl_options.subtitle = 'Showing Error Count > 1k and Avg Perf > 10MB'
    tbl_options.subtitle_fontSize = Pt(20)

    c.execute('''
    SELECT san, sw_name, slot_port, avPerf, error_type, error_count
        FROM
            PortErrorCnt
        WHERE 
            error_count > 999
            AND
            avPerf > 10
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
    print data


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
    print data


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
#--------------------------------------------------------------------
#--------------------------------------------------------------------
