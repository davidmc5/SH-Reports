from slLib import *
import slDeck
from sqlLib import *


def multiDeck(tbl_options):
    ''' Creates a slide deck with the agreagated data from all
    the reports downloaded'''

    #Open Slide Deck Template
    prs = Presentation(slides_template_path)
    tbl_options.presentation = prs

    #connect to database
    conn = tbl_options.dbConnection
    c = conn.cursor()
    
    #retrieve this report's variables
    customer, csvPath, shName, sanName, shYear = tbl_options.custData
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
    names = []
    for san in tbl_options.sanList:
        names.append(san[0])
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
        SUM(ports.total_devices)
    FROM
        switches s

    INNER JOIN 
        ports ON s.sw_name = ports.sw_name
    GROUP BY
        s.sw_fabric, s.sw_model
    ORDER BY
        s.sw_fabric, s.sw_model, count
    ''')

    data = c.fetchall()
    #format data with group headers (remove the group = first column data)
    data = groupHeader(data)

    
    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
    headers = [('Switch Model',
                'Total Switches',
                'Total Ports',
                'Unlicensed Ports',
                'Unused Ports',
                'ISL Ports',
                'Hosts',
                'Disks',
                'Total Devices')]

    headers.extend(data)
    data = headers
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
        active_zoneCfg != 'N/A'
    ORDER BY
        sw_fabric
   ''')
    data = c.fetchall()
    
    #covert data on 'dbUsed' column from Bytes to MB
    data = slDeck.formatDbUsed(data)

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
    
    #format data with group headers (remove the group = first column data)
    data = groupHeader(data)

    
    #Add column headers to print on the slide table
    # this is a single tuple with the column names
    # as the very first record of the 'data' list
    # first column for 'fabric' will be printed on a single dividing row
    
    headers = [('Switch Model',
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
   

#--------------------------------------------------------------------
#--------------------------------------------------------------------
    # END OF SLIDES
    slDeck.saveDeck(tbl_options)
#--------------------------------------------------------------------
#--------------------------------------------------------------------
