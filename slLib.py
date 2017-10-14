#pptx slide manipulation
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

from shPaths import *
from shLib import getCsvData


class Table_Options:
    """ Sets either defaults or given options for a table"""
    def __init__(self):
        #set default options
        self.presentation = None

        self.custData = None
        #self.csvPath = None

        #self.csvPathList = None # a list of csvPath from file colector loop
        self.sanList = None # the list of SAN names and csvPath of retrieved csv files

        
        self.csvFile = None
        self.csvColumns = []

        self.dbConnection = None # db Connection handler: conn = sql.connect(sqlite_file)
        self.dbTableName = None #name of the db table to create
        self.dbColNames = None #Column names to use on the dbtable [(names,...)]
        
        self.hBand = False #sets horizontal shade banding
        self.vBand = False #sets vertical shade banding
        self.font_size = Pt(18) #<<<--------CHANGE to just use an integer
        self.font_bold = False
        self.font_italic = False
        self.font_name = 'Calibri'
        self.font_color = RGBColor(12, 34, 56) # Black
        self.text_hAlign = PP_ALIGN.CENTER #LEFT, CENTER, RIGHT --Try literal conversion ast.eval() 
        self.text_vAlign = MSO_ANCHOR.MIDDLE #TOP, MIDDLE, BOTTOM -- to enter only a string options = 'right'
        self.first_row = True
        self.first_col = False
        self.last_row = False
        self.last_col = False
        self.left = Inches(1.0)
        self.top = Inches(2.25)
        
        self.title = ' '
        self.subtitle = ' '

        #TODO?
        #from pptx.enum.dml import MSO_THEME_COLOR
        #font.color.theme_color = MSO_THEME_COLOR.ACCENT_1


    

def format_table(table, options):
    def iter_cells(table):
        for row in table.rows:
            for cell in row.cells:
                yield cell

    #set given table options            
    table.horz_banding = options.hBand
    table.vert_banding = options.vBand
    table.last_col = options.last_col
    table.first_row = options.first_row
    table.last_row = options.last_row
    table.first_col = options.first_col
    table.last_col = options.last_col

    for cell in iter_cells(table):
        
        #the row height adjusts automatically to font size.
        #set margins as a table options?
        cell.margin_left = cell.margin_right = Pt(10)
        cell.margin_top = cell.margin_bottom = 0
        cell.vertical_anchor = options.text_vAlign
        
        for paragraph in cell.text_frame.paragraphs:
            paragraph.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            run = cell.text_frame.paragraphs[0]

            # A run is each line of text in a paragraph
            # Just looking for line 0 because...
            # ...iterating over multiple runs in a cell disables horizontal alignment
            # (pptx bug?)
            run.font.size = options.font_size
            run.font.bold = options.font_bold
            run.font.italic = options.font_italic
            run.font.name = options.font_name
            run.alignment = options.text_hAlign
            #run.font.color.rgb = options.font_color

##------------------------------------------------------------------------------------VVVVVVVVV

def create_table(slide, data, options):
    title = options.title
    subtitle = options.subtitle
    width = Inches(10) #total width of table
    height = Inches(.1) #total height of table
    left = options.left
    top = options.top
    rows = len(data)
    cols = len(data[0])

    #create table
    shapes = slide.shapes
    shapes.title.text = title

    #Add subtitle
    #add_textbox(left, top, width, height)
    subtxt = shapes.add_textbox(Inches(1),Inches(1.5), Inches(10), Inches(0.5))
    subtxt.text = subtitle
    subtxt.text_frame.paragraphs[0].font.size = Pt(30)

                                  
    table = shapes.add_table(rows, cols, left, top, width, height).table

    #input data to table's cells
    for row_idx,row in enumerate(data):
        count = 0 #count the numner of non-empty cells
        for col_idx, cell in enumerate(row):
            table.cell(row_idx, col_idx).text = str(cell)
            if cell != None: count += 1

        #Merge group row cells

        #Check if all row cells are empty (optionally except the first) = GROUP HEADER ROW
        if count <= 1: # header row
            #mergeCellsHorizontally(table, row_idx, start_col_idx, end_col_idx)
            mergeCellsHorizontally(table, row_idx, 0, cols-1)

            #change row color
            table.cell(row_idx, 0).fill.solid()
            # set foreground (fill) color to a specific RGB color
            #table.cell(row_idx, 0).fill.fore_color.rgb = RGBColor(0xFB, 0x8F, 0x00) # Orange
            #table.cell(row_idx, 0).fill.fore_color.rgb = RGBColor(0x00, 0xFF, 0x00) #lima green
            table.cell(row_idx, 0).fill.fore_color.rgb = RGBColor(0x66, 0xFF, 0x66) #pastel green green


    # Manual way to set specific columns' widths
    #(if omited, all columns are equally sized)
    #table.columns[0].width = Inches(1.0)
    #table.columns[1].width = Inches(1.0)


    #Adjust column widths based on max width of all cells in a column
    maxLen = []
    for row in xrange(len(data)):
        for cel, value in enumerate(data[row]):

            #check to use only the largest string space-separated segment
            #This is because powerpoint will break the header row (containing spaces)
            #into two lines

            #For header row only, split words into multiple lines
            if row == 0:
                value = str(value)
                #For header row only, split words into multiple lines
                words = value.split(' ')
            else:
                #For the rest of rows, display the table cell as a single line string
                words = [value]
            
            for word in words:                    
                txtLen = len(str(word))
                try:
                    if txtLen > maxLen[cel]:
                        maxLen[cel] = txtLen
                except:
                    #maxLen[] is an empty list on first pass (header row) 
                    maxLen.append(txtLen)



    #<<<----------------------------
    # TO-DO: AUTO-ADJUST THE CORRECTION FACTOR FOR EACH POINT SIZE
    # BETWEEN 10 AND 18 points
    # PROBABLY NON-LINEAR
    #<<<----------------------------
    
    #for each column, set the width BASED ON THE MAX LENGHT OF THE DATA STRING
    #the number of columns is the len() of any row. Using first row, data[0].
    for col in xrange(len(data[0])):
        #Adjust column width to max length of text
        #table.columns[col].width = ( Pt(20)+ Pt(20 * maxLen[col] * 0.55) ) 
        table.columns[col].width = ( Pt(18)+ Pt(18 * maxLen[col] * 0.60) )
        #table.columns[col].width = ( Pt(12)+ Pt(12 * maxLen[col] * 0.65) )
        #table.columns[col].width = ( Pt(10)+ Pt(10 * maxLen[col] * 0.70) )

    #Format table

    #num_rows = len(data)
    
    if rows < 14:
        options.font_size = Pt(18)
        
    elif rows < 15:
        options.font_size = Pt(17)
        
    elif rows < 16:
        options.font_size = Pt(16)
        
    elif rows < 19:
        options.font_size = Pt(15)
        
    elif rows < 20:
        options.font_size = Pt(14)
        
    elif rows < 22:
        options.font_size = Pt(13)
         
    elif rows < 24:
        options.font_size = Pt(12)
        
    elif rows < 26:
        options.font_size = Pt(11)
        
    else:
        options.font_size = Pt(10)

       
    format_table(table, options)
    return table
##------------------------------------------------------------------------------------^^^^^^^^^




def create_single_table(options):
    #Wrapper to create a slide with just one table from csv data

    #generate slide for table
    prs = options.presentation
    title_only_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_only_slide_layout)
    
    #get data to populate table
    shData = getCsvData(options)

    #create table
    table = create_table(slide, shData, options)

    return table


def create_single_table_db(shData, options):
    #Wrapper to create a slide with just one table from db data

    #generate slide for table
    prs = options.presentation
    title_only_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_only_slide_layout)
    
    #create table
    table = create_table(slide, shData, options)

    return table


# merge cells vertically
#https://groups.google.com/forum/#!topic/python-pptx/cVRP9sSpEjA
def mergeCellsVertically(table, start_row_idx, end_row_idx, col_idx):
    row_count = end_row_idx - start_row_idx + 1
    column_cells = [r.cells[col_idx] for r in table.rows][start_row_idx:]

    column_cells[0]._tc.set('rowSpan', str(row_count))
    for c in column_cells[1:]:
        c._tc.set('vMerge', '1')

# merge cells horizontally
#https://groups.google.com/forum/#!topic/python-pptx/cVRP9sSpEjA
def mergeCellsHorizontally(table, row_idx, start_col_idx, end_col_idx):
    col_count = end_col_idx - start_col_idx + 1
    row_cells = [c for c in table.rows[row_idx].cells][start_col_idx:end_col_idx]
    row_cells[0]._tc.set('gridSpan', str(col_count))
    for c in row_cells[1:]:
        c._tc.set('hMerge', '1')



def copy_slide(pres,pres1,index):
    source = pres.slides[index]
    #blank_slide_layout = _get_blank_slide_layout(pres)
    blank_slide_layout = pres.slide_layouts[blank_slide_layout_index]
    dest = pres1.slides.add_slide(blank_slide_layout)

    for shp in source.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

    # NOTE: six.iteritems = dictionary.iteritems() on Python 2 and dictionary.items() on Python 3.
    for key, value in six.iteritems(source.part.rels):
        # Make sure we don't copy a notesSlide relation as that won't exist
        if not "notesSlide" in value.reltype:
            dest.part.rels.add_relationship(value.reltype, value._target, value.rId)
    return dest


def move_slide(presentation, old_index, new_index):
        xml_slides = presentation.slides._sldIdLst  # pylint: disable=W0212
        slides = list(xml_slides)
        xml_slides.remove(slides[old_index])
        xml_slides.insert(new_index, slides[old_index])



