"""
power_bi_gd.py

functions that allow access to Marcu's version of the Atlas Golden Doc
"""
import pandas as pd
import intel_ww as iw
from numpy import nan

BUSINESS = 'Business'.replace(' ','')
SEGMENT = 'Segment'.replace(' ','')
PRODUCT_FAMILY = 'Product Family'.replace(' ','')
FULL_NAME_IN_SPEEDATLAS = 'Full Name in SPEED/Atlas'.replace(' ','')
SPEED_ID = 'Speed ID'.replace(' ','')
TYPE = 'Type'.replace(' ','')
A0_TI = 'A0 TI'.replace(' ','')
PRQABLE_TI = 'PRQable TI'.replace(' ','')
PRQ = 'PRQ'.replace(' ','')
PLATFORM_PV = 'Platform PV'.replace(' ','')
CHILD_COMPONENTS = 'Child Components'.replace(' ','')
PHASE = 'Phase'.replace(' ','')
IP_FAMILY = 'IP Family'.replace(' ','')
IP_FREEZE = 'IP Freeze'.replace(' ','')
GFX_CONFIG = 'Gfx Config'.replace(' ','')
MEDIA__DISPLAY_IP = 'Media / Display IP'.replace(' ','')
PROCESS = 'Process'.replace(' ','')

SI_PRODUCT = 'Si Product'.strip()


def str_ww_or_none( col_str ):
    """
    Try to conver to a ww or return what was passed in
    """
    ww = iw.WW_from_string( col_str ) 
    if ww is None:
        return col_str
    return str(ww)

def strip( s ):
    if isinstance(s, str):
        return s.strip()
    
    return s

def dash_or_empty_to_nan( s ):
    if isinstance(s, str):
        if s == '--' or s == '':
            return nan
    return s

def power_bi_to_df( excel_path, skiprows=[0,1]):
    ef = pd.ExcelFile( excel_path)
    df = ef.parse(skiprows=skiprows)
    df.columns = [c.strip().replace(' ','') for c in df.columns]
    for col in df.columns:
        df[col] = df[col].apply(strip)
        df[col] = df[col].apply(str_ww_or_none)
        df[col] = df[col].apply(dash_or_empty_to_nan)
    
    #print(df.loc[df[TYPE] == SI_PRODUCT, [FULL_NAME_IN_SPEEDATLAS,A0_TI, PRQ]])
    
    df = populate_earliest_ti_from_children(df=df)

    #print(df.loc[df[TYPE] == SI_PRODUCT, [FULL_NAME_IN_SPEEDATLAS,A0_TI, PRQ]])
    return df

def types( df, type_col = TYPE):
    return list(df[type_col].unique())

def children_names_from_str( names ):
    try:
        children = [s[1:].strip() for s in names.split('\n')]
        return children
    except:
        return None

def populate_earliest_ti_from_children( df, which_type=SI_PRODUCT, type_col=TYPE, date_cols = [A0_TI]):
    for _, row in df.loc[df[type_col] == which_type].iterrows():
        children = children_names_from_str(row[CHILD_COMPONENTS])

        for col in date_cols:
            current_ww = iw.WW_from_string(ww_string=row[col])
            if current_ww is not None:
                continue
            for name in children:
                name_ww_str = df.loc[df[FULL_NAME_IN_SPEEDATLAS] == name, col ].values[0]
                name_ww = iw.WW_from_string(ww_string=name_ww_str)
                if name_ww is None:
                    continue
                if current_ww is None:
                    current_ww = name_ww
                    continue
                if current_ww > name_ww:
                    name_ww = current_ww
            if current_ww is None:
                print("No ", col, ' found for ', row[FULL_NAME_IN_SPEEDATLAS])
                continue
            df.loc[ df[SPEED_ID] == row[SPEED_ID], col ] = str(current_ww)

    return df

def not_na(l, r):
    if pd.notna(r):
        return r
    return l    
        

if __name__ == '__main__':
    GOLDEN_DOC = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ScottsWorkbench/PowerBIGD 23-Nov-20.xlsx'
    BRONZE_DOC = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - General/PreConceptBronzeDoc.xlsx'
    
    g_df = power_bi_to_df(excel_path=GOLDEN_DOC)
    b_df = pd.ExcelFile(BRONZE_DOC).parse(skiprows=[0])
    print('BRONZE: ----------')
    print(b_df.loc[b_df['Name'].str.contains('DG2'),['Name','A0TI']])
    print('GOLDEN: ----------')
    print(g_df.loc[g_df[FULL_NAME_IN_SPEEDATLAS].str.contains('DG2'),[FULL_NAME_IN_SPEEDATLAS, 'A0TI']])
    print('MERGED:----------')
    c_df = pd.merge(left=b_df, right=g_df, how='left', left_on='Name', right_on=FULL_NAME_IN_SPEEDATLAS, suffixes=['','_y'])
    print(c_df.columns)
    print(c_df.loc[b_df['Name'].str.contains('DG2'),['Name', 'A0TI', 'A0TI_y']])
    c_df['A0TI'] = c_df['A0TI'].combine(c_df['A0TI_y'], not_na )
    c_df['PRQ'] = c_df['PRQ'].combine(c_df['PRQ_y'], not_na )
    drop_cols = []
    for col in c_df.columns:
        if col.endswith('_y'):
            drop_cols.append(col)
    c_df = c_df.drop(labels=drop_cols, axis=1)
    print('Copied:----------')
    print(c_df.loc[b_df['Name'].str.contains('DG2'),['Name', 'A0TI','PRQ']])
    print(c_df.columns)





    





