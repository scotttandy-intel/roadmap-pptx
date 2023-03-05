"""
intel_ww.py
Set of classes and functions related to managing work weeks

NOTE: Like everything else - work weeks are complicated.

This code should be rewritten to better align with the ISO 8601 standard

Author: Scott Tandy   27-Sep-20

Copyright Intel Corporation 2020
"""
DAYS_PER_YEAR = 365   ##Ignoring leap year for most WW Applications
WW_PER_YEAR = 52
DAYS_PER_WW = 7
QTR_PER_YEAR = 4
WW_PER_QTR = 13
VALID_MAX_WW = 53

MIN_YEAR_SUPPORTED = 0
MAX_YEAR_SUPPORTED = 50


import pandas as pd
import datetime
import math
import random
from calendar import Calendar

from isoweek import Week

##### ISO CALENDAR ISO 8601 https://en.wikipedia.org/wiki/ISO_8601
"""

There are several mutually equivalent and compatible descriptions of week 01:

the week with the year's first Thursday in it (the formal ISO definition),
the week with 4 January in it,
the first week with the majority (four or more) of its days in the starting year, and
the week starting with the Monday in the period 29 December – 4 January.
As a consequence, if 1 January is on a Monday, Tuesday, Wednesday or Thursday, it is in week 01. 
If 1 January is on a Friday, Saturday or Sunday, it is in week 52 or 53 of the previous year 
(there is no week 00). 28 December is always in the last week of its year.
"""


class WW:
    Q1_WW_START = 1
    Q2_WW_START = Q1_WW_START + WW_PER_QTR
    Q3_WW_START = Q2_WW_START + WW_PER_QTR
    Q4_WW_START = Q3_WW_START + WW_PER_QTR

    ### Eventually we will support WW.n to specify the exact date
    MONDAY = 1
    TUESDAY = 2
    WEDNESDAY = 3
    THURSDAY = 4
    FRIDAY = 5
    SATURDAY = 6
    SUNDAY = 7


    def __init__(self, ww=None, year=None, datetime=None, override_range_checks=False):
        """
        Use factory method...
        """

        self.ww = ww
        self.year = year
        self._datetime = datetime
        self.override_range_checks = override_range_checks
        self.__check_ww__()

        return
    
    def __str__(self):
        if self.ww < 10:
            return f" {self.ww}'{self.year:02}"    
        
        return f"{self.ww:02}'{self.year:02}"
        #return str(self.ww)+"'"+str(self.year)
   
    def __eq__(self, other):
        if (self.ww == other.ww) and (self.year == other.year):
            return True

        return False
    
    def __lt__(self, other):
        if (self.year < other.year):
            return True
        
        if (self.year > other.year):
            return False
        
        if self.ww < other.ww:
            return True

        return False
    
    def __check_ww__(self):
        if self.ww is not None:
            if type(self.ww) is not int:
                raise ValueError('self.ww must be of type int')
            if not self.override_range_checks:
                if self.ww < 1 or self.ww > VALID_MAX_WW:
                    raise ValueError('WW must be between 1 and 53 inclusive')
        if self.year is not None:
            if self.year <MIN_YEAR_SUPPORTED or self.year >MAX_YEAR_SUPPORTED:
                raise ValueError('Year must be between '+str(MIN_YEAR_SUPPORTED)+' and '+str(MAX_YEAR_SUPPORTED))
    
    def to_datetime(self):
        """
        Uses a brute force attack to return at datetime for the Monday of the
        WW
        """
        c = Calendar()
        true_year = self.year + 2000

        for month in range(1,13):
            for day in c.itermonthdates(year=true_year, month=month):
                if day.isocalendar()[1] == self.ww:
                    return day
        
        raise ValueError("Error: could not find a valid datetime for " + str(self))



    def day_of_the_year(self, work_day=MONDAY):
        """
        The first day of WW1 is the "first day"
        """
        return (self.ww-1)*DAYS_PER_WW + work_day
    
    def qtr_of_year(self):
        if self.ww < WW.Q2_WW_START:
            return 1
        if self.ww < WW.Q3_WW_START:
            return 2
        if self.ww < WW.Q4_WW_START:
            return 3
        return 4

    def copy(self):
        return WW(ww=self.ww, year=self.year, datetime=self._datetime)
    """
    Deprecated - was confusing since it modified the ww sent in -
    the user can just reassign it if the want that behavior....
    def add_wws(self, wws):

        #Add number of WWs (can be >52) to this ww object

        years_to_add = math.trunc(wws / WW_PER_YEAR)

        wws_to_add = wws - (years_to_add * WW_PER_YEAR)

        ww_total = self.ww + wws_to_add

        if ww_total > WW_PER_YEAR:
            years_to_add = years_to_add + 1
            ww_total = ww_total - WW_PER_YEAR

        self.ww = ww_total
        self.year = self.year + years_to_add
        self.__check_ww__()
        
        return self
    """
    """
    def add_wws(self, wws):
        Add number of WWs (can be >52) to this ww object
 
        years_to_add = math.trunc(wws / WW_PER_YEAR)

        wws_to_add = wws - (years_to_add * WW_PER_YEAR)

        ww_total = self.ww + wws_to_add

        if ww_total > WW_PER_YEAR:
            years_to_add = years_to_add + 1
            ww_total = ww_total - WW_PER_YEAR
        
        ww = WW(ww=ww_total, year=self.year+years_to_add)

        return ww
    """
    def add_wws(self, wws):
        self_datetime = self.to_datetime()
        week_timedelta = datetime.timedelta(weeks=wws)
        self_datetime = self_datetime + week_timedelta
        return WW_from_date(date_field=self_datetime)
    
    def subtract_wws(self, wws):
        return self.subtract_wws(wws= -1 * wws)
    
    def ww_delta_from(self, ww):
        
        if self.year == ww.year:
            return ww.ww - self.ww
        
        if self.year > ww.year:
            whole_years_wws = (self.year - ww.year -1) * WW_PER_YEAR
            partial_ww = (WW_PER_YEAR-ww.ww + self.ww) 
            return -1 * (whole_years_wws + partial_ww)
        
         #### self.year < ww.year:
        whole_years_wws = (ww.year - self.year -1) * WW_PER_YEAR
        partial_ww = (WW_PER_YEAR-self.ww + ww.ww) 
        return whole_years_wws + partial_ww
    

    def ww_delta_from_dt(self, ww):
        ww_dt = ww.to_datetime()
        self_dt = self.to_datetime()
        td = ww_dt-self_dt
        return round(td.days / 7)

def WW_from_date( date_field ):
    if type(date_field) is pd._libs.tslib.Timestamp:
        ww, year = field_to_ww(date_field=date_field)
    
    elif type(date_field) is datetime.datetime or type(date_field) is datetime.date:
        pandas_datetime = pd.Timestamp(date_field)
        ww, year = field_to_ww(date_field=pandas_datetime)
    
    else:
        raise ValueError('Cannot parse date_field:', date_field)

    return WW(ww=ww, year=year)




def WW_from_string( ww_string, quarter_treatment= 0.5):
    return_tuple = field_to_ww(date_field = ww_string, quarter_treatment=quarter_treatment)

    if return_tuple is not None:
        return WW(ww=return_tuple[0], year=return_tuple[1])
    
    return None

def random_WW(min_year, max_year):
    return WW(ww=random.randint(1,52), year=random.randint(min_year,max_year))
        
def field_to_ww(date_field, quarter_treatment= 0.5):
    """
    date_field - might be a pandas date or text field: WWww:yy or Qn'yy
    quater_treatment = if it is in a "quarter" where to place it in the quarter
                        (0.5 is half way through the quarter, 0 is at the start)
    """
    ww = 0
    year = 0

    ## Handle if it is pandas datefield
    if type(date_field) is pd._libs.tslib.Timestamp:
        month = date_field.month
        ww = date_field.weekofyear
        year = date_field.year

        ### Special case when the date is actually in a work week 
        ### from last year or next year as compared to the year date
        if month == 12 and ww == 1:
            year = year + 1
        
        if month == 1:
            if ww == 52 and (date_field + pd.Timedelta(weeks=1)).weekofyear != 53:
                year = year -1        
            elif ww == 53:
                year = year - 1

        return (ww, abs(year) % 100)
    
    ## Handle if it is already a WW and YY
    try:
        ww,year = ww_text_to_tuple(date_field)
        year = abs(year) % 100
        return (ww, year)
    except:
        pass

    ## Handle if it is in Qn'YY
    try:
        ww, year = quarter_text_to_ww_year_tuple(date_field, quarter_treatment= quarter_treatment)
        year = abs(year) % 100
        return (ww, year)
    except:
        pass
    return None

def ww_text_to_tuple(ww_text):
    """
    Takes string in form WW25'20 or ww25'20 and returns
    (int work_week, int year) """
    SPLIT = "'"
    ALT_SPLITS = [ "`",'"',':']
    
    copy_text = ww_text

    ww_text = ww_text.strip().upper()
    ww_text = ww_text.replace('W','')
    
    ### Switch to standard split character
    for a_s in ALT_SPLITS:
        ww_text = ww_text.replace(a_s, SPLIT)

    ### Hack to remove (A0) from the end of the WW data
    if '(' in ww_text:
        ww_text = ww_text.split('(')[0]
    
    try:
        ww = int(ww_text.split(SPLIT)[0].strip())
        year = int(ww_text.split(SPLIT)[1].strip())

    except:
        raise ValueError('Bad format:'+ copy_text+":")
    
    if (ww < 1) or (ww > WW_PER_YEAR):
        raise ValueError('Invalid workweek:'+ str(ww))
    
    if (year < MIN_YEAR_SUPPORTED) or (year > MAX_YEAR_SUPPORTED):
        raise ValueError('Invalid year:' + str(year))
    
    return (ww, year)     

def quarter_text_to_ww_year_tuple(q_text, quarter_treatment= 0.5):
        """
        Takes string in form Q1'20 and returns
        (int work_week, int year) """

        if quarter_treatment < 0.0 or quarter_treatment > 1.0:
            raise ValueError('Quarter treatment must be between 0.0 and 1.0')
        
        SPLITS = ["’", "'"]
        
        copy_text = q_text

        q_text = q_text.replace('q','')
        q_text = q_text.replace('Q','')

        if len(q_text.split(SPLITS[0])) == 2:
            split = SPLITS[0]
        elif len(q_text.split(SPLITS[1])) == 2:
            split = SPLITS[1]
        else:
            raise ValueError('No quarter /  year delimeter found:'+ copy_text+":")
        
        try:
            quarter = int(q_text.split(split)[0])
            year = int(q_text.split(split)[1])

        except:
            raise ValueError('Bad format:'+ copy_text+":")
        
        if (quarter < 1) or (quarter > 4):
            raise ValueError('Invalid quarter:'+ str(quarter))
        
        if (year < MIN_YEAR_SUPPORTED) or (year > MAX_YEAR_SUPPORTED):
            raise ValueError('Invalid year:' + str(year))
        
        ww_per_q = int(WW_PER_YEAR / 4)

        whole_q_wws =(quarter - 1 ) * ww_per_q
        
        partial_wws = int(quarter_treatment * ww_per_q )

        ww = whole_q_wws + partial_wws

        return (ww, year)

if __name__ == '__main__':

    ww = WW(ww=1, year=19)
    for i in range(100):
        ww_before = ww.add_wws(-1*i)
        ww_after = ww.add_wws(i)
        print(ww.ww_delta_from(ww_before), ww.ww_delta_from_dt(ww_before), \
            ww.ww_delta_from(ww_after), ww.ww_delta_from_dt(ww_after))
    
    ww_before = ww.add_wws(-53)
    print( ww.ww_delta_from(ww_before))
    print(ww.ww_delta_from_dt(ww_before))
    """
    ww_list = []

    for i in range(10):
        ww1 = random_WW(min_year=19, max_year=27)
        ww2 = random_WW(min_year=19, max_year=27)
        ww_list = ww_list + [ww1, ww2]
    
    print([str(ww) for ww in ww_list])
    ww_list = sorted(ww_list)
    print([str(ww) for ww in ww_list])

    for year_count in range(MAX_YEAR_SUPPORTED-MIN_YEAR_SUPPORTED):
        year = MIN_YEAR_SUPPORTED + year_count
        for ww_count in range(WW_PER_YEAR):
            ww_known = WW(ww=ww_count+1, year=year)
            ww_str = f"Ww{ww_count+1}'{year}"
            ww_from_str = WW_from_string(ww_str)
            
            if (ww_known != ww_from_str):
                print("error: on ", ww_known, ww_str, ww_from_str)
        
        for ww_count in range(WW_PER_YEAR):
            ww_known = WW(ww=ww_count+1, year=year)
            ww_str = f"{ww_count+1}'{year}"
            ww_from_str = WW_from_string(ww_str)
            
            if (ww_known != ww_from_str):
                print("error: on ", ww_known, ww_str, ww_from_str)

    for quarter in range(4):
        q_text = 'Q'+str(quarter+1)+"'20"
        print(q_text, quarter_text_to_ww_year_tuple(q_text=q_text, quarter_treatment=0.0))
        print(q_text, quarter_text_to_ww_year_tuple(q_text=q_text, quarter_treatment=0.5))
        print(q_text, quarter_text_to_ww_year_tuple(q_text=q_text, quarter_treatment=1.0))

    print(field_to_ww(pd.Timestamp(year=2020, day=5, month=5), quarter_treatment= 0.5))
    print(field_to_ww("WW05'21", quarter_treatment= 0.5))
    print(field_to_ww("Q2'21", quarter_treatment= 0.5))

    for ww in range(1,53):
        for day in range(1,8):
            my_ww = WW(ww=ww, year=20)
            print(ww, day, my_ww.day_of_the_year(work_day=day), my_ww.qtr_of_year(), my_ww.to_datetime() )
    """
    
    
    
            
