"""
roadmap_table.py
Functions related to reading roadmap information from excel tables

Author Scott Tandy
Copyright Intel Corporation 2020
"""
from collections import OrderedDict

import pandas as pd
import datetime
import math

import intel_ww as iw
import roadmap_helper as rh
import annotations as an

### Sheet Names
ANNOTATIONS_SHEET = 'Annotations'
TYPE = 'Type'
TAG = 'TAG'
MILESTONE = 'MILESTONE'
NON_MILESTONE = 'NON MILESTONE'

A = 'A'
B = 'B'
C = 'C'
D = 'D'

ALL = 'ALL'

ANNOTTION_SLIP = 'SLIP'
SLIP_START_SHORT_NAME = A
SLIP_START_MILESTONE = B
SLIP_END_SHORT_NAME = C
SLIP_END_MILESTONE = D

ANNOTATION_INLINEWWS = 'INLINEWWS'
INLINEWWS_BUSINESS = A
INLINEWWS_SEGMENT = B
INLINEWWS_PRODUCT = C
INLINEWWS_MILESTONE = D

ANNOTATIONS_SHEET = 'ANNOTATIONS'
COMMAND_COL = 'COMMAND'
VALUE_STRING = 'VALUE'

ANNOTTION_SLIP = 'SLIP'
SLIP_START_SHORT_NAME = A
SLIP_START_MILESTONE = B
SLIP_END_SHORT_NAME = C
SLIP_END_MILESTONE = D

### Default column headings 
BUSINESS = 'Business'
SEGMENT = 'Segment'
SPEED_ID = 'Speed ID'
PRODUCT_FAMILY = 'Product Family'
FULL_NAME_IN_SPEED_ATLAS = 'Full Name in SPEED/Atlas'
TYPE = 'Type'
CHILD_COMPONENTS = 'Child Components'
PHASE = 'Phase'
A0_TI = 'A0 TI'
PRQABLE_TI = 'PRQable TI'
PRQ = 'PRQ'
FCS_RTS = 'FCS/RTS'
IP_FAMILY = 'IP Family'
IP_FREEZE = 'IP Freeze'
GFX_CONFIG = 'Gfx Config'
MEDIA_IP = 'Media IP'
DISPLAY_IP = 'Display IP'
PROCESS = 'Process'
UPDATED = 'Updated'
COMMENTS = 'Comments'

SIMPLESEGMENT = rh.SIMPLESEGMENT
SHORTHAND_NAME = rh.SHORTHAND_NAME

SEGMENT_LABEL = 'Segment Label'

### TYPE Options
DIE = 'Die'
SI_PRODUCT = 'Si Product'   ## This is the packaged SOC
PLATFORM = 'Platform'
INTEGRATED_IP = 'Integrated IP'
GCD_DIE = 'GCD Die'

### Business Constants
DATACENTER = 'Datacenter'
CLIENT = 'Client'
CLIENT_DISCRETE = 'Client Discrete'
INTEGRATED = 'Integrated'


####
#### ConceptTable Defintions
####
CT_SHEET_PRECONCEPTS = 'PRECONCEPTS'
CT_CATEGORY = 'CATEGORY'
CT_BUSINESS = 'BUSINESS'
CT_SEGMENT = 'SEGMENT'
CT_DESCRIPTION = 'DESCRIPTION'
CT_NAME ='NAME'
CT_COMPONENTOF = 'COMPONENTOF'

def manage_sheets(excel_file_path, all=False, combine=False):
    """
    Returns a list of Data Frames based on the sheets in an Excel file
    """
    xl = pd.ExcelFile(excel_file_path)
    if combine == True:
        all = True  ## Let's assume the user wants to look at all the sheets
        raise NotImplementedError

    df_list = []
    print(xl.sheet_names)  # see all sheet names
    for sheet_name in xl.sheet_names:
        print("reading / parsing: ", sheet_name)

        df_list.append( xl.parse(sheet_name=sheet_name) )  # read a specific sheet to DataFrame

        if all is False:
            break

    return df_list, xl.sheet_names

class RoadmapTable:
    """
    Takes "golden doc" Excel file, parses it, and then provides helper functions to access the data
    in a product oriented way....
    """
    def __init__( self, df=None, path_to_roadmap=None, path_to_helper_file=None ):
        if df is None and path_to_roadmap is None:
            raise ValueError("Must supply dataframe or path to roadmap file to build Roadmap class")
        self.df = None
        self.a_df = None ### Annotations dataframe
        self.path_to_roadmap = path_to_roadmap
        self.is_concept_format = False

        if df is None:
                xl = pd.ExcelFile(self.path_to_roadmap)
                sheet_names_upper = [sn.upper() for sn in xl.sheet_names]

                if sheet_names_upper[0] == ANNOTATIONS_SHEET.upper():
                    raise ValueError('Cannot have ANNOTATIONS sheet as first sheet')
                
                if sheet_names_upper[0] == CT_SHEET_PRECONCEPTS.upper():
                    self.is_concept_format = True
                
                if self.is_concept_format:
                    skip_rows = [0]    ### Skip the first "description"
                else:
                    skip_rows = None
                
                self.df = xl.parse(sheet_name=xl.sheet_names[0], skiprows=skip_rows )

                print(self.df.columns)
                print(self.df.shape)

                if ANNOTATIONS_SHEET.upper() in sheet_names_upper:
                    a_index = sheet_names_upper.index(ANNOTATIONS_SHEET.upper())
                    self.a_df = xl.parse(sheet_name=xl.sheet_names[a_index])
                    print('Annotations shape:', self.a_df.shape )
                else:
                    self.a_df = None
                    print('No annotations found')
        else:
            self.df = df
        
        if self.is_concept_format is False:
            self._fix_up_business_names()

        if pd.notna(path_to_helper_file):
            self.rh = rh.RoadmapHelper(path_to_helper_file=path_to_helper_file)
            self._add_simplesegment_column()
            self._add_shorthand_name_column()
        else:
            self.rh = None
        return
    
    def _fix_up_business_names(self):
        ### Turn Client/Integrated into Integrated
        self.df.loc[(self.df[BUSINESS]==CLIENT) & (self.df[SEGMENT]==INTEGRATED ), BUSINESS] = INTEGRATED
        return
    
    def _add_simplesegment_column(self):
        self.df[SIMPLESEGMENT] = self.df[FULL_NAME_IN_SPEED_ATLAS].apply(self.rh.name_to_segment)
        self.df.loc[:,SIMPLESEGMENT].fillna(self.df.loc[:,SEGMENT], inplace=True)
        return
    
    def _add_shorthand_name_column(self):
        self.df[SHORTHAND_NAME] = self.df[FULL_NAME_IN_SPEED_ATLAS].apply(self.rh.name_to_shorthand)
        self.df.loc[:,SHORTHAND_NAME].fillna(self.df.loc[:,FULL_NAME_IN_SPEED_ATLAS], inplace=True)
        return
    
    def rows_matching_col_tuples(self, col_tuples):
        df = self.df
        for col_name, col_value in col_tuples:
            df = df.loc[df[col_name] == col_value]
        return df

    def rows_isin_col_tuples(self, col_tuples):
        df = self.df
        for col_name, col_value in col_tuples:
            df = df.loc[df[col_name].isin(col_value)]
        return df

    def business_list(self, business_column = BUSINESS):
        return list(self.df[business_column].unique())
    
    def segment_list(self, business, business_column = BUSINESS, segment_column = SEGMENT):  
        return list(self.df.loc[self.df[business_column]==business][segment_column].unique())
    
    def simple_segment_list(self, business,  business_column = BUSINESS, \
        segment_column=SEGMENT):
        
        if self.is_concept_format:
            raise ValueError("Simple segment list not available for concept_doc types")

        return list(self.df.loc[ (self.df[business_column]==business)]\
            [SIMPLESEGMENT].unique() )
     
    def by_type(self, business, segment, type_name, business_column = BUSINESS, \
            segment_column = SEGMENT, type_column = TYPE ):
        
        return self.df.loc[ (self.df[business_column] == business) & \
            (self.df[segment_column] == segment) & \
            (self.df[type_column] == type_name)]
    
    def by_type_simplesegment(self, business, simple_segment, type_name, business_column = BUSINESS, \
             ss_column = SIMPLESEGMENT, type_column = TYPE ):
        if self.is_concept_format:
            raise ValueError("By type simple segment list not available for concept_doc types")

        return self.df.loc[ (self.df[business_column] == business) & \
            (self.df[ss_column] == simple_segment) & \
            (self.df[type_column] == type_name)]

    def children_speed_ids(self, children_str ):
        if self.is_concept_format:
            raise ValueError("Children speed id list not available for non concept_doc types")
        try:
            children_split = children_str.split('\n')
        except:
            return []

        id_list = []
        for child in children_split:
            if len(child.split(':')) < 2:
                continue
            child_name = child.split(':')[1].strip()
            name_lookup = self.df.loc[self.df[FULL_NAME_IN_SPEED_ATLAS] == child_name][SPEED_ID]
            if len(name_lookup.values) != 0:
                id_list.append(name_lookup.values[0])
            else:
                #print('-'*10, "couldn't lookup ", child)
                pass
        return id_list
    
    def by_speed_id(self, speed_id):
        return self.df.loc[(self.df[SPEED_ID] == speed_id)]
    
    def by_shorthand_name(self, shorthand_name ):
        return self.df.loc[(self.df[SHORTHAND_NAME] == shorthand_name)]
    
    def annotations_list(self):
        if self.a_df is None:
            return None
        al = []
        for i, row in self.a_df.iterrows():
            ad = {}
            for col in self.a_df.columns:
                ad[col] = row[col]
            al.append(ad)  
        return al


class SiProduct:
    def __init__( self, shorthand_name, roadmap_table):
        """
        name - str - name of the SiProduct
        roadmap - type Roadmap
        """
        self.shorthand_name = shorthand_name
        self.roadmap = roadmap_table
        self.si_product_speed_id = self.my_row()[SPEED_ID].values[0]

        return
    
    def my_row(self):
        return self.roadmap.by_shorthand_name(self.shorthand_name)
    
    def _milestone(self, milestone_tag, latest_valid_milestone_WW = None, \
                if_empty_fill_with_earliest_child=True):
        """
        returns a iw.WW structure or None based on the tag
        """
        try:
            ms_text = self.my_row()[milestone_tag].values[0]
        except:
            return None
        
        delta_wws = None

        if latest_valid_milestone_WW is not None:
            try:
                delta_wws = iw.delta_string_to_wws(ms_text)
            except:
                pass
        
        if delta_wws is not None:
            ww = latest_valid_milestone_WW.add_wws(wws=delta_wws)
        else:
            ww = iw.WW_from_string(ms_text)

        if ww is None and if_empty_fill_with_earliest_child:
            ww = self.earliest(column_name=milestone_tag, override_si_prod_date=False)

        return ww
    
    def milestone_dict(self, tag_list, if_empty_fill_with_earliest_child=True ):
        ms_dict = {}

        for tag in tag_list:
            #print(tag, '-', end='')
            
            ms = self._milestone(milestone_tag=tag, latest_valid_milestone_WW=latest_valid_milestone_WW, \
                 if_empty_fill_with_earliest_child=if_empty_fill_with_earliest_child )
            
            if ms is not None:
                ms_dict[tag] = ms
                #print(str(ms))
            else:
                pass
                #print(' (not found)')
        
        return ms_dict
        
        
    def children_sids(self):
        c_sids = self.roadmap.children_speed_ids(children_str=self.my_row()[CHILD_COMPONENTS].values[0])
        #print('c_sids:', c_sids)
        return c_sids

    def early_late(self, column_name, earlist=True,  override_si_prod_date=False):
        """
        column_name - str - which column in df to look in
        earliest - bool - Check for earliest date if True else check for latest date
        ovveride_si_product_date - bool - 
                        look for earliest value even if defined in SiProduct row
        """
        earliest_latest_ww = None
        
        my_row = self.roadmap.by_speed_id(self.si_product_speed_id)
        
        ## Check to see if one is defined
        if len(my_row[column_name].values) > 0:
            early_late_ww = iw.WW_from_string(my_row[column_name].values[0])
            if early_late_ww is not None:
                earliest_latest_ww = early_late_ww
                ## We have a date, so only keep lookng if we are supposed to ovveride
                if not override_si_prod_date:
                    return earliest_latest_ww
        
        ## At this point we didn't find a date so we should check the children
        if self.children_sids() is not None:
            for sid in self.children_sids():
                my_row = self.roadmap.by_speed_id(sid)  ## Get the row from df matching sid

                if len(my_row[column_name].values) > 0:
                    early_late_ww = iw.WW_from_string(my_row[column_name].values[0])
                    
                    if early_late_ww is None:
                        continue
                    
                    #if early_late_ww is not None and earliest_latest_ww is not None:
                    #    print(earliest_latest_ww, early_late_ww, earliest_latest_ww.ww_delta_from(early_late_ww) )
                    
                    if earliest_latest_ww is None:
                        earliest_latest_ww = early_late_ww
                    else:
                        if earlist is True:
                            if earliest_latest_ww.ww_delta_from(early_late_ww) < 0:
                                earliest_latest_ww = early_late_ww
                        else:
                            if earliest_latest_ww.ww_delta_from(early_late_ww) > 0:
                                earliest_latest_ww = early_late_ww
        
        return earliest_latest_ww 

    def earliest(self, column_name=A0_TI, override_si_prod_date=False):
        return self.early_late(column_name=column_name, earlist=True, \
            override_si_prod_date=override_si_prod_date )

    def latest(self, column_name=PRQ, override_si_prod_date=False):
        return self.early_late( column_name=column_name, earlist=False, \
            override_si_prod_date=override_si_prod_date)

class ConceptSiProduct:
    """
    Removes dependency on SpeedID specifying unique row....
    """
    def __init__( self, major_value, minor_value, name, my_row):
        """
        name - str - name of the SiProduct
        roadmap - type Roadmap
        """
        self.major_value = major_value
        self.minor_value = minor_value
        self.name = name
        self.milestones_and_annotations = self._build_row_dictionary(my_row=my_row)

        return
    
    def _just_this_type(self, d, t):
        """
        Helper function to creat a subset dict with just this type of items
        """
        just_t = OrderedDict()
        for key, value in d.items():
            if isinstance(value, t):
                just_t[key] = value
        return just_t

    def _build_row_dictionary(self, my_row):
        """
        Looks at each cell in the row and figures out if it is a work week, relative work week
        or annotation.
        """
        my_row_dict = OrderedDict()
        for col_name in list( my_row.axes[0]):
            col_text = my_row.at[col_name]

            if pd.isnull(col_text):
                continue
            try:
                annotation = an.cell_text_to_annotation(my_col=col_name, text=col_text)
            except:
                annotation = None
            
            if annotation is not None and an.is_annotation_class(annotation):
                my_row_dict[col_name] = annotation
                continue
            
            ### This is a bit of a hack - and not taking full advantage of all
            ### the fields provide by the annotation class
            ### will need to re-write to put the annotation class object
            ### into the dict instead of just converting it to the equivalent of
            ### Text entered into a particular column
            if annotation is not None and an.is_milestone_class(annotation):
                col_name = annotation.name
                col_text = annotation.ww_text

            try:
                delta_wws = iw.delta_string_to_wws(col_text)
            except:
                delta_wws = None

            if delta_wws != None:
                just_ms = self._just_this_type(d=my_row_dict, t=iw.WW )
                current_index = len(just_ms.keys())

                if current_index == 0:
                    raise ValueError(\
                        'Cannot have relative milestone with no previous set milestone '+col_name+' - ' + col_text )
                
                previous_milestone =my_row_dict[list(just_ms.keys())[current_index-1]]
                
                my_row_dict[col_name] = previous_milestone.add_wws(delta_wws)
                continue

            ww = iw.WW_from_string(col_text)

            if ww is not None:        
                my_row_dict[col_name] = ww
            else:
                print(f'cell_text={col_text} - not date or milestone/annotation')

        return my_row_dict

    def milestones( self ):  
        ms_dict = OrderedDict()

        for key, value in self.milestones_and_annotations.items():
            if isinstance(value, iw.WW):
                ms_dict[key] = value
        
        return ms_dict

    def annotations( self ):   
        an_dict = OrderedDict()

        for key, value in self.milestones_and_annotations.items():
            if an.is_annotation_class(value):
                an_dict[key] = value  
        return an_dict
    
    def milestone_before( self, col_name):
        """
        Used when annotating and explict milestones are not provided in the annotation
        """
        all_cols = list(self.milestones_and_annotations.keys())
        index_of_col_name = all_cols.index(col_name)
        just_before = all_cols[:index_of_col_name]
        just_before_backwards = just_before[::-1]
        for col in just_before_backwards:
            if isinstance(self.milestones()[col], iw.WW):
                return col
        raise ValueError('No milestone found before'+col_name)
    
    def milestone_after( self, col_name):
        """
        Used when annotating and explict milestones are not provided in the annotation
        """
        all_cols = list(self.milestones_and_annotations.keys())
        index_of_col_name = all_cols.index(col_name)
        just_after = all_cols[index_of_col_name+1:]
        for col in just_after:
            if isinstance(self.milestones()[col], iw.WW):
                return col
        raise ValueError('No milestone found after'+col_name)
        
if __name__ == "__main__":
    import roadmap_config as rc
    
    CONCEPT_DOC_EXCEL = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/PreConceptBronzeDoc.xlsx'
    CONFIG_DOC_EXCEL = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/PreConceptConfig.xlsx'

    r_t = RoadmapTable(path_to_roadmap=CONCEPT_DOC_EXCEL)
    #r_c = rc.RoadmapConfig(path_to_roadmap_config_excel=CONFIG_DOC_EXCEL)

    for i, row in r_t.df.iterrows():
        major_value=row.at['Description']
        minor_value = row.at['Segment']
        name = row.at['Name']
        si = ConceptSiProduct(major_value=major_value, minor_value=minor_value,name=name, my_row=row)
        print('*'*5, si.name)
        #print(si.milestones_and_annotations)
        for ann in si.annotations():
            print(ann, si.milestone_before(col_name=ann))
            print(ann, si.milestone_after(col_name=ann))
        #print(si.milestones())
        #print(si.annotations())




    """
    GOLDEN_DOC_EXCEL = \
        "/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ScottsWorkbench/Golden Doc - W39'20 - Final.xlsx"
 
    ROADMAP_HELPER_EXCEL = \
        '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ScottsWorkbench/Database and Naming POR Doc - Alternate - AddSimpleSegment.xlsx'
   
    rt = RoadmapTable(path_to_roadmap=GOLDEN_DOC_EXCEL, path_to_helper_file=ROADMAP_HELPER_EXCEL)
    
    for biz in rt.business_list():
        if biz == INTEGRATED:
            continue
        for ss in rt.simple_segment_list(business=biz):
            rows = rt.by_type_simplesegment(business=biz, simple_segment=ss, type_name=SI_PRODUCT )
            if rows.shape[0] == 0:
                continue
            for i, row in rows.iterrows():
                sp = SiProduct(name=row[SHORTHAND_NAME], si_product_speed_id=row[SPEED_ID], roadmap_table=rt)
                print(sp.name, sp.si_product_speed_id, sp.earliest(), sp.latest())
             

    for biz in rt.business_list():
        for seg in rt.segment_list(business=biz):
            for i, si_prod in rt.by_type(business=biz, segment=seg, type_name = SI_PRODUCT).iterrows():
                #print(biz,'-',seg,'-',si_prod[FULL_NAME_IN_SPEED_ATLAS],'-',si_prod[SHORTHAND_NAME],'-', si_prod[SIMPLESEGMENT])
                sp = SiProduct(name=si_prod[SHORTHAND_NAME], si_product_speed_id=si_prod[SPEED_ID],\
                    roadmap=rt)
                print()
    """
                
    



    