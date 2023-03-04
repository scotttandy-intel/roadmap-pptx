"""
Funcitons for manaing the "helper" excel spreadsheet that augments the table data
in the roadmap_table.

"""
import pandas as pd

ROADMAP_HELPER_ROWS_TO_SKIP = [0]

### Column Headings
UNNAMED__0 = 'Unnamed: 0'
IP_GENERATION = 'IP Generation'
BUSINESS = 'Business'
SIMPLESEGMENT = 'SimpleSegment'
SEGMENT = 'Segment'
SHORTHAND_NAME = 'Shorthand Name'
DEPRECATED = 'Deprecated'
IP_FAMILY = 'IP Family'
PRODUCT_CODE_NAME = 'Product Code Name'
BRAND_NAME = 'Brand Name'
GFX_FAMILY__CARBON_ID_ = 'GFX Family [Carbon ID]'
GFX_CONFIG__CARBON_ID_ = 'GFX Config [Carbon ID]'
MEDIA_FAMILY__CARBON_ID_ = 'Media Family [Carbon ID]'
MEDIA_CONFIG__CARBON_ID_ = 'Media Config [Carbon ID]'
DISPLAY_FAMILY__CARBON_ID_ = 'Display Family [Carbon ID]'
DISPLAY_CONFIG__CARBON_ID_ = 'Display Config [Carbon ID]'
COMPONENT_DIE__SPEED_ID_ = 'Component Die [SPEED ID]'
SI_PRODUCT__SPEED_ID_ = 'Si Product [SPEED ID]'
PLATFORM_NAME__SPEED_ID_ = 'Platform Name [SPEED ID]'
STATUS = 'Status'

ALL_COLS = [UNNAMED__0, IP_GENERATION, BUSINESS, SIMPLESEGMENT, SEGMENT, SHORTHAND_NAME, DEPRECATED,\
     IP_FAMILY, PRODUCT_CODE_NAME, BRAND_NAME, GFX_FAMILY__CARBON_ID_, GFX_CONFIG__CARBON_ID_,\
          MEDIA_FAMILY__CARBON_ID_, MEDIA_CONFIG__CARBON_ID_, DISPLAY_FAMILY__CARBON_ID_,\
               DISPLAY_CONFIG__CARBON_ID_, COMPONENT_DIE__SPEED_ID_, SI_PRODUCT__SPEED_ID_,\
                    PLATFORM_NAME__SPEED_ID_, STATUS, ]

SI_PRODUCT_ONLY = 'SiProductName'
SI_PRODUCT_SID = 'SiProductSID'

### Simple Segments Names
ENTHUSIAST = 'Enthusiast'
PERFORMANCE = 'Performance'
MAINSTREAM = 'Mainstream'
ENTRY = 'Entry'
COMPUTE = 'Compute'
MULTI_USE = 'Multi-Use'
PRO = 'Pro'
PERFORMANCE_PLUS = 'Performance Plus'
BASIC = 'Basic'
PREMIUM = 'Premium'

DC_GPU_SIMPLESEGMENTS = [COMPUTE, MULTI_USE]
DISCRETE_CLIENT_SIMPLESEGMENTS = [ENTRY, MAINSTREAM, PERFORMANCE, PERFORMANCE_PLUS,ENTHUSIAST, PRO]

class RoadmapHelper:
    """
    Reads in the Excel file that has program data that isnt' contianed in Atlas (yet)
    """
    def __init__( self, path_to_helper_file):
        self.df = None
        self.path_to_helper_file = path_to_helper_file
        self._load_excel_file()
        self._add_si_name_column()
        self._add_si_SID_column()
        return
    
    def _load_excel_file(self):
        self.df = pd.read_excel(io=self.path_to_helper_file,skiprows=ROADMAP_HELPER_ROWS_TO_SKIP)
        return
    
    def _add_si_name_column(self):
        self.df[SI_PRODUCT_ONLY] = self.df[SI_PRODUCT__SPEED_ID_].apply(series_split_name)
    
    def _add_si_SID_column(self):
        self.df[SI_PRODUCT_SID] = self.df[SI_PRODUCT__SPEED_ID_].apply(series_split_id).apply(safe_int)
    
    def lookup(self, value, lookup_col, return_col):
        try:
            value = self.df.loc[self.df[lookup_col] == value][return_col].values[0]
            return value
        except:
            return None
    
    def name_to_segment(self, name ):
        return self.lookup(value=name, lookup_col=SI_PRODUCT_ONLY , return_col=SEGMENT)
    
    def name_to_shorthand(self, name ):
        return self.lookup(value=name, lookup_col=SI_PRODUCT_ONLY , return_col=SHORTHAND_NAME)


def split_name_id(combined_cell_to_parse):

    combined = combined_cell_to_parse
    ## If multi-line just deal with the first one for now....
    if '\n'in combined:
        combined = combined.split('\n')[0]
    
    if '[' in combined:
        split = combined.split('[')
        name = split[0].strip()
        
        if ':' in name:
            name = name.split(':')[1].strip()
        
        try:
            sid = int(split[1].split(']')[0])
        except:
            sid = None

    elif '(TBA)'in combined:
        split = combined.split('(TBA)')
        name = split[0].strip()
        sid = None
    else:
        name = None
        sid = None
    
    return name, sid

def series_split_name( s ):
    return split_name_id(s)[0]

def series_split_id(s):
    return split_name_id(s)[1]

def safe_int( i ):
    try:
        i = int(i)
    except:
        i = -1

    return int(i)

def make_constants(col_name):
    """
    Helper function that prints out constants and their values so you can cut 
    and paste back into the code
    """
    col_name_name = col_name.upper()

    for c in " \'\",.:;!@#$%^&*(){}/\\[]-+":
        col_name_name = col_name_name.replace(c, '_')
    return col_name_name

if __name__ == "__main__":
    
    DATABASE_AND_NAMING_EXCEL = \
        '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ScottsWorkbench/Database and Naming POR Doc - Alternate - AddSimpleSegment.xlsx'
    rh = RoadmapHelper(DATABASE_AND_NAMING_EXCEL)
    for seg in rh.df[SIMPLESEGMENT].unique():
        print(f"{make_constants(seg)} = '{seg}'")
    """
    col_str = '['

    for col_name in list(rh.df.columns):
        col_name_name = col_name.upper()
        col_name_name = col_name_name.replace(' ', '_')
        col_name_name = col_name_name.replace('[', '_')
        col_name_name = col_name_name.replace(']', '_')
        col_name_name = col_name_name.replace(':', '_')  
        #print(f"{col_name_name} = '{col_name.strip()}'")
        col_str = col_str +col_name_name+', '
    col_str = col_str + ']'
    print(col_str)
 
    for i, row in rh.df.iterrows():
        combined = row[SI_PRODUCT__SPEED_ID_]

        print(split_name(combined), split_id(combined))

    rh.df[SI_PRODUCT_ONLY] = rh.df[SI_PRODUCT__SPEED_ID_].apply(series_split_name)
    rh.df[SI_PRODUCT_SID] = rh.df[SI_PRODUCT__SPEED_ID_].apply(series_split_id).apply(safe_int)
 
    for i, row in rh.df.iterrows():
        print(i,row[SI_PRODUCT_ONLY], row[SI_PRODUCT_SID], \
            rh.name_to_segment(row[SI_PRODUCT_ONLY]))
    """

