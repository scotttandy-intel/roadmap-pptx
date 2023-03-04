"""
Class that manages the translation of an Excel configuration file into the
keys / row_names / etc. that enable the parsing of the roadmap input data
and the formatting of the roadmap file
"""
from collections import OrderedDict
import pandas as pd

### Sheet names in the config file

COMMAND_COLUMNS_SHEET = 'COMMANDS'
COMMAND= 'COMMAND'
VALUE = 'VALUE'
COMMENT = 'COMMENT'
COMMAND_COLUMNS = [COMMAND,VALUE,COMMENT]

TRUE_VALUE_TEXT = 'TRUE'
FALSE_VALUE_TEXT = 'FALSE'


### Define roadmap_table column header meanings
ROADMAP_COLUMNS_SHEET = 'ROADMAP_COLUMNS'
TAG= 'TAG'
TYPE = 'TYPE'
## VALUE COLUMN IS DEFINED ABOVE ALREADY
CS_COLUMNS = [TAG, TYPE, VALUE]
## Types
NON_MILESTONE = 'Non Milestone'
MILESTONE = 'Milestone'


## Roadmap Tag Keys
SWIMLANES_SHEET = 'SWIMLANES'
BUSINESS = 'BUSINESS'
SEGMENT = 'SEGMENT'
SEGMENT_LABEL = 'SEGMENT LABEL'
SL_COLUMNS = [BUSINESS,SEGMENT, SEGMENT_LABEL]


class RoadmapConfig:
    def __init__(self, path_to_roadmap_config_excel):
        
        ef = pd.ExcelFile(path_to_roadmap_config_excel)

        ##### Parse the ROADMAP_COLUMNS SHEET 
        self.commands = self.parse_sheet(ef=ef, sheet_name=COMMAND_COLUMNS_SHEET,\
            col_names=COMMAND_COLUMNS, tag_column=COMMAND, value_column=VALUE, only_type=None)
        
        self.swimlanes_hierarchy, self.major_column_name, self.minor_column_name, self.name_col\
             = sheet_to_major_minor_name(ef=ef, sheet_name=SWIMLANES_SHEET)

    def parse_sheet(self, ef, sheet_name, col_names,tag_column=TAG, value_column=VALUE,\
         only_type=None):
        
        new_dict = OrderedDict()

        try:
            i = ef.sheet_names.index(sheet_name)
        except:
            raise ValueError(f"Error finding'{sheet_name}' sheet name")
        usecols = list(range(len(col_names)))

        df = ef.parse(sheet_name=ef.sheet_names[i], usecols=usecols)
        df.columns = col_names   ## overwrite the columns
        for i, row in df.iterrows():
            raw_row_value = row[value_column]
            if pd.isna( raw_row_value):
                print('Skipping:', row[usecols[0]])
                continue
            if only_type is not None:
                if row[TYPE] != only_type:
                    continue
            
            if type(raw_row_value) is str:
                raw_row_value=raw_row_value

            new_dict[row[tag_column]]=raw_row_value

        return new_dict

def sheet_to_major_minor_name( ef, sheet_name ):
    """
    Business   Segment  ProductName
    Datacenter Compute  ProductName
    Datacenter Compute  ProductName  
    """
    try:
        i = ef.sheet_names.index(sheet_name)
    except:
        raise ValueError(f"Error finding'{sheet_name}' sheet name")

    df = ef.parse(sheet_name=ef.sheet_names[i])

    df.columns = df.columns[:3]  ### Only 3 columns, and uppercase
    
    major_column, minor_column, name_col = (df.columns[0], df.columns[1], df.columns[2])

    df[minor_column].fillna(value='', inplace=True)
    df[name_col].fillna(value='', inplace=True)

    od = OrderedDict()
    for biz in df[major_column].unique():
        od[biz]=OrderedDict()
        for seg in df.loc[df[major_column] == biz][minor_column].unique():
            for i,row in df.loc[(df[major_column] == biz) & (df[minor_column] == seg)].iterrows():
                od[biz][seg] = seg  ## Note this level of the dict is now redudant....could refactor
    
    return (od, major_column, minor_column, name_col)

if __name__ == '__main__':
    RC_PATH = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/ElastiConfig.xlsx'

    rc = RoadmapConfig(path_to_roadmap_config_excel=RC_PATH)
