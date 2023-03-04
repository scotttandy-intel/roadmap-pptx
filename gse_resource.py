"""
gsc_resource.py
Helper functions to parse GSE resource model data
"""

import pandas as pd
from collections import OrderedDict

### Sheet Names
YEARS_NROWS = 1
DATA_SKIP_ROWS = [0]


#PRE_MAP_COLUMNS

#POST MAP COLUMNS
MAJORCATEGORY="MajorCategory"
CATEGORY="Category"
PROJECT="Project"
DROP1='Drop1'
SHEET='Sheet'
NOTES="Notes"
Q1_2020="Q1'2020"
Q2_2020="Q2'2020"
Q3_2020="Q3'2020"
Q4_2020="Q4'2020"
Q1_2021="Q1'2021"
Q2_2021="Q2'2021"
Q3_2021="Q3'2021"
Q4_2021="Q4'2021"
Q1_2022="Q1'2022"
Q2_2022="Q2'2022"
Q3_2022="Q3'2022"
Q4_2022="Q4'2022"
Q1_2023="Q1'2023"
Q2_2023="Q2'2023"
Q3_2023="Q3'2023"
Q4_2023="Q4'2023"
Q1_2024="Q1'2024"
Q2_2024="Q2'2024"
Q3_2024="Q3'2024"
Q4_2024="Q4'2024"

REMAP_ONE = col_map = {'Unnamed: 0':MAJORCATEGORY, 'Unnamed: 1':CATEGORY, 'Unnamed: 2':PROJECT, \
    'Unnamed: 3':DROP1,\
           'Unnamed: 4':NOTES}

DROP_COLS = [DROP1,'Unnamed: 25', 'Unnamed: 26']


def roll_up( excel_path, sheet_names ):
    ef = pd.ExcelFile(excel_path)

    df = None

    for sheet in sheet_names:
        #print(sheet)
        sheet_df = parse_and_clean(pd_excel_file=ef, sheet_name=sheet)
        if df is None:
            df = sheet_df
        else:
            df = pd.concat([df, sheet_df])
    
    return df



def parse_and_clean(pd_excel_file, sheet_name):

    df = pd_excel_file.parse(sheet_name=sheet_name,skiprows=DATA_SKIP_ROWS).dropna(axis=0, how='all')

    years_df = pd_excel_file.parse(sheet_name=sheet_name,nrows=YEARS_NROWS)

    df = rename_qtr_columns_drop_extra(df=df, years_df=years_df)
    
    drop_rows = df[MAJORCATEGORY].notna() | df[CATEGORY].notna()  #Used later

    df = copy_categories_and_projects(df)

    df=df[~drop_rows]

    df = add_column(df=df, column_value=sheet_name, column_name=SHEET, col_after_index=3)

    return df


def rename_qtr_columns_drop_extra(df, years_df):
    l_rename = OrderedDict()
    year_track=0
    for l0,l1 in zip(years_df.columns, df.columns):
        if str(l0)[:2]=='20':
            year_track=l0
        if l1[0]=='Q':
            rename = l1[:2]+'\''+str(year_track)
            l_rename[l1]=rename
        else:
            l_rename[l1]=l1
    df.rename(mapper=l_rename, axis='columns', inplace=True)
    df.rename(mapper=REMAP_ONE,axis='columns', inplace=True)
    df.drop(labels=DROP_COLS, axis=1,inplace=True)

    return df

def copy_categories_and_projects(df):
    major_cat = ''
    cat=''
    for i, row in df.iterrows():
        if not pd.isna(row[MAJORCATEGORY]):
            major_cat = row[MAJORCATEGORY]
            cat=''
        if not pd.isna(row[CATEGORY]):
            cat=row[CATEGORY]

        df.at[i,MAJORCATEGORY]=major_cat
        df.at[i, CATEGORY]=cat
    
    return df

def add_column(df, column_value, column_name=SHEET, col_after_index=3):
    """
    df   pandas dataframe with GSE information
    column_name   str() name of the new column
    col_after_index: none->don't move it   0->make first column,\
         1-> make second column
    """

    df[column_name] = column_value

    if col_after_index is not None:
        re_order = []
        for i, c in enumerate(df.columns):
            if c == column_name:
                continue
            if i == col_after_index:
                re_order.append(column_name)
                continue
            re_order.append(c)

    df = df[re_order]
    return df

class Project:
    def __init__(self, project_name, project_column=PROJECT):
        self.project_name = project_name
        self.project_column = project_column
        return
    
    def groupby_team(self, df):
        return df.loc[df[self.project_column] == self.project_name].dropna().loc[:,SHEET:Q4_2024]
    
    def mountain_chart(self, df, figsize=(20,5)):
        ml_df = df.loc[df[self.project_column] == self.project_name].dropna().loc[:,SHEET:Q4_2024]
        row_labels = list(ml_df[SHEET].unique())
        ml_df = ml_df.loc[:,Q1_2020:Q4_2024].reset_index(drop=True)
        ml_df = ml_df.T
        ml_df.columns = row_labels
        return ml_df.plot( kind='area', title=self.project_name+' loading', figsize=figsize, x_compat=True)



if __name__ == '__main__':
    pass