from slLib import *
#import slDeck
from sqlLib import *


def compDeck(tbl_options):
    ''' Creates a slide deck comparing data from all
    the reports with more that one date'''

    #Open Slide Deck Template
    prs = Presentation(slides_template_path)
    tbl_options.presentation = prs

    #connect to database
    conn = tbl_options.dbConnection
    c = conn.cursor()

    #retrieve this report's variables
    customer, csvPath, shName, sanName, shDate, shYear = tbl_options.custData
#--------------------------------------------------------

    slidesExist = 0
#--------------------------------------------------------
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
        #print san[0]
        #!
        # append only the max date file names.
        #!
        names.append(san[1]) # shName
    count = len(names)
    title = "SAN Health Report"
    subtitle = customer + ' - Reports Included (' + str(count) + '):'

    textSlide(prs, title, subtitle, names)

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
    LIMIT 1

    ''')
    data = c.fetchall()

    #print data

    # #keep track if there are any non-blank slides in this deck
    # if len(data) > 0:
    #     slidesExist +=1

    #format data with group headers (remove the group = first column data)
    #data = groupHeader(data)

    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
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

    if len(data) > 0:
        slidesExist +=1


    #covert data on 'dbUsed' column from Bytes to MB
    #data = formatDbUsed(data)
    #reformat dbUsed data

    #Add column headers to print on the slide table
    # this is a tuple with the column names
    # as the very first record of the 'data' list
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
    # END OF SLIDES

    if slidesExist > 0:
        saveDeck(tbl_options)
        logEntry('Slides Created', customer, 'Compared')
#--------------------------------------------------------------------
#--------------------------------------------------------------------
