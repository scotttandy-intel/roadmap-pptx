"""
annotations.py
JSON formatted data that describes the annotation between milestones

NOTE: Need to improve exception handling and reporting to make it clearer to the caller
what went wrong - usually in JSON formatting etc.

Scott Tandy
Copyright Intel Corporation 2020
"""

import json

SINGLE_QUOTE = "'"
DOUBLE_QUOTE = '"'

TRUE = 'True'
FALSE = 'False'
### Annoation Types
TYPE = 'type'          # Key
TYPE_TEXT = 'text'     # Value for Text annotation
TYPE_DELTA = 'delta'   # Value for Delta annotation

### General Keys
TEXT = 'text'
START = 'start'
END = 'end'
AFTER_TEXT = 'after'


### TextFormat related constants
FILL_RGB = 'fill_rgb'
LINE_RGB = 'line_rgb'
LINE_WIDTH = 'line_width'
LINE_DASH = 'line_dash'
FONT_NAME = 'font_name'
FONT_SIZE = 'font_size'
FONT_RGB = 'font_rgb'
FONT_BOLD = 'font_bold'
FONT_ITALIC = 'font_italic'
FONT_UNDERLINE = 'font_underline'

### Delta related constants
TYPE_PHASE = 'phase'
START_MS = 'start'
END_MS = 'end'

### Date format keys and values
DATE_FORMAT = 'df'
DF_D = 'date'
DF_WWs = 'ww'
DF_QTR = 'q'
DF_YRS = 'y'


def sq_to_dq( s ):
    return s.replace(SINGLE_QUOTE, DOUBLE_QUOTE)

def text_to_bool( s ):
    if not isinstance(s, str):
        raise ValueError('Not string type')

    s = s.strip().lower()
    if s == TRUE.lower():
        return True
    if s == FALSE.lower():
        return False
    
    raise ValueError('Neither true or false text found' + s )

def text_rgb_triple( s, length=3 ):
    if not isinstance(s, str):
        raise ValueError('Not string type')
    
    t = tuple( [int(s) for s in s.replace('(','').replace(')','').split(',')] )
    if len(t) != length:
        raise ValueError('Did not find tuple with length ' + str(length) + ' in ' + s )
    return t

def cell_to_json( c_str ):
    """
    Takes JSON in standards compliant format EXCEPT everywhere you would use a
    DOUBLE_QUOTE (") there is a SINGLE_QUOTE (') instead.

    This is to allow EXCEL users to enter the following into a cell:
    ="{'key':'value'}"
    Since excel only allows for " as the string delimeter - and JSON requires
    the " as the string delimeter - and we want to enter JSON into a Excel cell,
    we follow the convention that the single quote is used instead of the double quote.
    """

    c_str = sq_to_dq(c_str)

    return json.loads(c_str)

def an_key_in_dict(k, d):
    for key in d.keys():
        if k.lower() == key.lower():
            return key
    
    raise ValueError('Key not found'+k)

def an_value_for_key(k, d):
    return d[an_key_in_dict(k=k, d=d)]


def is_annotation_class( obj ):
    if isinstance(obj, Text):
        return True
    
    if isinstance(obj, Delta):
        return True
    
    return False

"""
Annotation types:
default: If the cell value is 
        ="'this is some text'" 
        then this annotation will be rendered to the right hand side of the prevous milestone.
        
        NOTE: just entering text into the cell will result in non valid json (there are no quotes)
        and will not parse and will be ignored.

PHASE: text between two (usually adjacent) milestones
   type':'phase', text:'text to show', 'first':{'col':} }

LINE_TO: a line between (usually the previous milestone - to a milestone in another line)
"""
class TextFormat:
    def __init__(self, fill_rgb=None, line_rgb = None, line_width= None, line_dash = None,\
             font_name=None, font_size = None, font_rgb=None, font_bold=None, font_italic=None,\
             font_underline=False ):
        self.fill_rgb = fill_rgb
        self.line_rgb = line_rgb
        self.line_width = line_width
        self.line_dash = line_dash
        self.font_name = font_name
        self.font_size = font_size
        self.font_rgb = font_rgb
        self.font_bold = font_bold
        self.font_italic = font_italic
        self.font_underline = font_underline
        return

    
    def __str__(self):
        s = 'fill_rgb '+str(self.fill_rgb)+'\n'+\
            'line_rgb '+str(self.line_rgb)+'\n'+\
            'line_width '+str(self.line_width)+'\n'+\
            'line_dash '+str(self.line_dash)+'\n'+\
            'font_name '+str(self.font_name)+'\n'+\
            'font_size '+str(self.font_size)+'\n'+\
            'font_rgb '+str(self.font_rgb)+'\n'+\
            'font_bold '+str(self.font_bold)+'\n'+\
            'font_italic '+str(self.font_italic)+'\n'+\
            'font_underline '+str(self.font_underline)+'\n'
        return s

def dict_to_TextFormat(d):
    #FILL_RGB = 'fill_rgb' 
    try:
        key = an_key_in_dict(FILL_RGB, j)
        fill_rgb = text_rgb_triple(s=j[key])  
    except:
        fill_rgb = None

    #LINE_RGB = 'line_rgb' 
    try:
        key = an_key_in_dict(LINE_RGB, j)
        line_rgb = j[key]
    except:
        line_rgb = None
    
    #LINE_WIDTH = 'line_width'   
    try:
        key = an_key_in_dict(LINE_WIDTH, j)
        line_width = j[key]
    except:
        line_width = None
    
    #LINE_DASH = 'line_dash'
    try:
        key = an_key_in_dict(LINE_DASH, j)
        line_dash = j[key]
    except:
        line_dash = None
    
    #FONT_NAME = 'font_name'
    try:
        key = an_key_in_dict(FONT_NAME, j)
        font_name = j[key]
    except:
        font_name = None

    #FONT_SIZE = 'font_size'
    try:
        key = an_key_in_dict(FONT_SIZE, j)
        font_size = j[key]
    except:
        font_size = None

    #FONT_RGB = 'font_rgb
    # Assumes '(128,128,128)' value format
    try:
        key = an_key_in_dict(FONT_RGB, j)
        font_rgb = text_rgb_triple(s=j[key])   
    except:
        font_rgb = None

    #FONT_BOLD = 'font_bold'
    try:
        key = an_key_in_dict(FONT_BOLD, j)
        font_bold = text_to_bool(s=j[key])  
    except:
        font_bold = None
    
    #FONT_ITALIC = 'font_italic'
    try:
        key = an_key_in_dict(FONT_ITALIC, j)
        font_italic = text_to_bool(s=j[key])  
    except:
        font_italic = None
    
    #FONT_UNDERLINE = 'font_underline'
    try:
        key = an_key_in_dict(FONT_UNDERLINE, j)
        font_underline = text_to_bool(s=j[key])  
    except:
        font_underline= None
    
    return TextFormat(fill_rgb=fill_rgb, line_rgb=line_rgb, line_width=line_width, line_dash=line_dash,\
        font_name=font_name, font_size=font_size, font_rgb=font_rgb,
        font_bold=font_bold, font_italic=font_italic, font_underline=font_underline)

class Text:
    def __init__(self, my_col, text, text_format=None):
        self.my_col = my_col
        self.text = text
        self.text_format = text_format
    
    def to_text(self):
        return self.text


def text_to_Text(my_col, j):
   
    text_format = dict_to_TextFormat(d=j)

    try:
        text = an_value_for_key(k=TYPE_TEXT, d=j)
    except:
        raise ValueError(TEXT+" key not found")

    return Text(my_col=my_col, text=text, text_format=text_format)

class Delta:
    def __init__(self, my_col, text, start, end, date_format, after_text, text_format  ):
        self.my_col = my_col
        self.text = text
        self.start = start
        self.end = end
        self.date_format = date_format
        self.after_text = after_text
        self.text_format = text_format
    
    def to_text(self):
        return self.text


def text_to_Delta(my_col, j):
    try:
        text = an_value_for_key(k=TEXT,d=j)
    except:
        text = ''

    try:
        start = an_value_for_key(k=START,d=j)
    except:
        start = None
    
    try:
        end = an_value_for_key(k=END, d=j)
    except:
        end = None
    
    try:
        date_format = an_value_for_key(k=DATE_FORMAT, d=j)
    except:
        date_format = None
    
    try:
        after_text = an_value_for_key(k=AFTER_TEXT, d=j)
    except:
        after_text = ''

    text_format = dict_to_TextFormat(d=j)

    return Delta(my_col=my_col, text=text, start=start, end=end, date_format=date_format, \
                after_text=after_text, text_format=text_format)

def cell_text_to_annotation( my_col, text ):
    try:
        j = cell_to_json(c_str=text)
    except:
        raise ValueError('Error converting to JSON: ' + text)
    
    if isinstance(j, str):
        return Text(my_col=my_col, text=j)
    
    try:
        t = an_value_for_key(k=TYPE, d=j)
    except:
        raise ValueError('No '+TYPE+' key found')
    
    t = t.lower()

    if t == TYPE_TEXT.lower():
        return text_to_Text(my_col=my_col, j=j)
    
    if t == TYPE_DELTA.lower():
        return text_to_Delta(my_col=my_col, j=j)

    raise ValueError('Valid JSON -> No Valid Annotation: ' + text)



if __name__ == '__main__':
    from pandas import read_excel

    df = read_excel('/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/json.xlsx')

    col_name = df.columns[0]

    for i, row in df.iterrows():
        j = row[col_name]
        #print(j, '---', type(cell_to_json(c_str=j)), '-----', cell_to_json(c_str=j))
        print(j, '---',cell_to_Phase(c_str=j))



