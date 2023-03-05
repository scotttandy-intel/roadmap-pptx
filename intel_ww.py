"""
intel_ww.py

Set of classes and funcitons which wrap the isoweek open source
WorkWeek managment library

NOTE: Like everything else - work weeks are complicated.

Author: Scott Tandy   27-Sep-20

Copyright Intel Corporation 2020
"""
DAYS_PER_YEAR = 365   ##Ignoring leap year for most WW Applications
WW_PER_YEAR_USUALLY = 52
DAYS_PER_WW = 7
QTR_PER_YEAR = 4
WW_PER_QTR = 13
VALID_MAX_WW = 53
WW_PER_MONTH = 52.0 / 12.0

MIN_YEAR_SUPPORTED = 0
MAX_YEAR_SUPPORTED = 50
WW_YEAR_REAL_WORLD_YEAR = 2000

import pandas as pd
import datetime
import math
import random
from calendar import Calendar
from isoweek import Week

DELTA_PLUS_STR = '+'
DELTA_NEG_STR = '-'
DELTA_EQ_STR = '='

VALID_DELTA_STR = [DELTA_NEG_STR, DELTA_EQ_STR, DELTA_PLUS_STR]

DELTA_QTR_STR = 'Q'
DELTA_MONTH_STR = 'M'
DELTA_WEEK_STR = 'W'

def wws_in_year(year):
    """
    Full AD year (eg. 2001)
    """
    return len(list(Week.weeks_of_year(year)))

def wws_in_quarter(qtr, year):
    """
    In a 53 WW year - we add the extra ww in the last Q
    """
    if qtr != 4:
        return WW_PER_QTR
    
    if wws_in_year(year) == WW_PER_YEAR_USUALLY:
        return WW_PER_QTR  
    
    return WW_PER_QTR +1

def wws_to_text( wws ):
    if abs(wws) <=1:
        suffix = 'WW'
    else:
        suffix = 'WWs'
    
    return f"{wws}"+suffix

def wws_to_days( wws ):
    """
    wws - int - number of workweeks
    """
    return wws * DAYS_PER_WW

def wws_to_days_text( wws ):
    days = wws_to_qtr(wws=wws )

    if abs(days) <=1:
        suffix = 'day'
    else:
        suffix = 'days'

    return f"{days}"+suffix

def wws_to_qtr( wws, ndigits=1 ):
    """
    wws - int - number of workweeks
    """
    return round( float(wws) / float(WW_PER_QTR), ndigits)

def wws_to_qtr_text( wws, ndigits=1):
    qtrs = wws_to_qtr(wws=wws, ndigits=ndigits )

    if abs(qtrs) <=1:
        suffix = 'Q'
    else:
        suffix = 'Qs'

    if ndigits == 0:
        return f"{qtrs:.0f}"+suffix
    if ndigits == 1:
        return f"{qtrs:.1f}"+suffix
    if ndigits == 2:
        return f"{qtrs:.2f}"+suffix

    return f"{qtrs:.5f}" + suffix

def wws_to_years( wws, ndigits=1 ):
    """
    wws - int - number of workweeks
    """
    return round( float(wws) / float(WW_PER_YEAR_USUALLY), ndigits)

def wws_to_years_text( wws, ndigits=1):
    years = wws_to_years(wws=wws, ndigits=ndigits )

    if abs(years) <=1:
        suffix = 'Y'
    else:
        suffix = 'Ys'

    if ndigits == 0:
        return f"{years:.0f}"+suffix
    if ndigits == 1:
        return f"{years:.1f}"+suffix
    if ndigits == 2:
        return f"{years:.2f}"+suffix
    
    return f"{years:.5f}"+suffix

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

    def __init__(self, ww=None, year=None, date_time=None ):
        """
        Can supply ww/year or date_time
        """
        if ww is None and year is None and date_time is None:
            raise ValueError('Must supply ww/year or date_time to constructor')

        self.ww = ww
        self.year = year
        self._date_time = date_time
        self._isoweek = None

        if self._date_time is not None:
            self._isoweek = Week.withdate(self._date_time)
            self.ww = self._isoweek.week
            self.year = self._isoweek.year % 100
        else:
            self._isoweek = Week(year=year+WW_YEAR_REAL_WORLD_YEAR, week=ww)

        return
    
    def __str__(self):
        if self.ww < 10:
            return f" {self.ww}'{self.year:02}"    
        
        return f"{self.ww:02}'{self.year:02}"
        #return str(self.ww)+"'"+str(self.year)
   
    def __eq__(self, other):
        return self._isoweek == other._isoweek
    
    def __lt__(self, other):
        return self._isoweek < other._isoweek
    
    def __sub__(self, other):
        return self._isoweek-other._isoweek
    
    def contains(self, date_time):
        return self._isoweek.contains(date_time)
    
    def to_datetime(self, day_dot=1):
        """
        Leverages isoweek
        """
        return self._isoweek.day(day_dot + 1)

    def day_of_the_year(self, work_day=MONDAY):
        """
        The first day of WW1 is the "first day"

        You should use this very carefully - because the ww01'yy
        might actually be in yy-1 - and this is true on the other end.

        This function ends up returning the day of the year that the WW is actualy IN
        """
        day_from_start_td = self.to_datetime(day_dot=work_day) - \
            datetime.date(year=self.year+2000, month=1, day=1)
        
        return day_from_start_td.days
        
        #return (self.ww-1)*DAYS_PER_WW + work_day
    
    def qtr_of_year(self):
        if self.ww < WW.Q2_WW_START:
            return 1
        if self.ww < WW.Q3_WW_START:
            return 2
        if self.ww < WW.Q4_WW_START:
            return 3
        ### Doesn't consider the 53 WW year - but accidentially works OK...
        return 4

    def add_wws(self, wws):
        return WW_from_isoweek(self._isoweek + wws)
    
    def subtract_wws(self, wws):
        return self.add_wws(wws * -1)

    def ww_delta_from(self, ww):
        return ww._isoweek - self._isoweek
    
    def quater_delta_from(self, ww):
        """
        Used to figure out which quarter to position the WW....
        we don't care that the WW dates might span multiple calendar quarters...
        we care about which quarter the ww is positioned in.
        as defined in qtr_of_year().
        """
        ### same_year_case:
        self_qtr = self.qtr_of_year()
        ww_qtr = ww.qtr_of_year()

        if self.year == ww.year:
            return ww_qtr - self_qtr

        if ww.year > self.year:
            self_delta = QTR_PER_YEAR-self_qtr
            ww_delta = ww_qtr
            years = ww.year-self.year-1
            return self_delta + ww_delta + years * QTR_PER_YEAR
        
        self_delta = self_qtr
        ww_delta = QTR_PER_YEAR-ww_qtr
        years = self.year - ww.year - 1
        return -1 * (self_delta + ww_delta + years * QTR_PER_YEAR)


def WW_from_date( date_field ):
    return WW(date_time=date_field)

def WW_from_isoweek( iso_week, dot_day=1 ):
    date_time = iso_week.day(dot_day-1)
    return WW_from_date(date_field=date_time)

def WW_from_string( ww_string, quarter_treatment= 0.5):
    """
    ww_string -  WWww:yy or Qn'yy or YYYYwwWW
    quater_treatment = if it is in a "quarter" where to place it in the quarter
                        (0.5 is half way through the quarter, 0 is at the start)
    """
    ## Handle if it is YYYYwwWW(A0)

    try:
        ww, year = year_leading_ww_text_to_tuple( ww_string )
        return WW(ww=ww, year=year)
    except:
        pass

    ## Handle if it is already a WW and YY
    try:
        ww,year = ww_text_to_tuple(ww_string)
        year = abs(year) % 100
        return WW(ww=ww, year=year)
    except:
        pass

    ## Handle if it is in Qn'YY
    try:
        ww, year = quarter_text_to_ww_year_tuple(ww_string, quarter_treatment=quarter_treatment)
        year = abs(year) % 100
        return WW(ww=ww, year=year)
    except:
        pass
    
    return None

def random_WW(min_year, max_year):
    return WW(ww=random.randint(1,52), year=random.randint(min_year,max_year))


def year_leading_ww_text_to_tuple( ww_text ):
    """
    Handles dates in YYYYwwWW format introduced by our friend Marcus
    in his GoldenDoc output
    """
    did_split = ww_text.lower().split('ww')

    if len(did_split) != 2:
        raise ValueError("not in year_ww format")

    try:
        year = int(did_split[0])
    except:
        raise ValueError("not in year_ww format")
    
    if year >= WW_YEAR_REAL_WORLD_YEAR:
        year = year - WW_YEAR_REAL_WORLD_YEAR
    
    if year < MIN_YEAR_SUPPORTED or year > MAX_YEAR_SUPPORTED:
        raise ValueError("Year outside of bounds " + str(year))

    try:
        ww_str = did_split[1].split('(')[0].strip() # handle trailing '(B0)'
        ww = int(ww_str)
    except:
        raise ValueError("Error converting ww to year"+did_split[1])

    if ww < 1 or ww > VALID_MAX_WW:
        raise ValueError("Invalid work week " + str(ww))

    return (ww, year)


def ww_text_to_tuple(ww_text):
    """
    Takes string in form WW25'20 or ww25'20 and returns
    (int work_week, int year) """
    SPLIT = "'"
    ALT_SPLITS = [ "`",'"',':',]
    
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
    
    if (ww < 1) or (ww > WW_PER_YEAR_USUALLY):
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
        
        SPLIT = "'"
        ALT_SPLITS = [ "`",'"',':']
        
        copy_text = q_text
            ### Switch to standard split character
        
        for a_s in ALT_SPLITS:
            q_text = q_text.replace(a_s, SPLIT)

        q_text = q_text.replace('q','')
        q_text = q_text.replace('Q','')
        
        try:
            quarter = int(q_text.split(SPLIT)[0].strip())
            year = int(q_text.split(SPLIT)[1].strip())
        except:
            raise ValueError('Bad format:'+ copy_text+":")
        
        if (quarter < 1) or (quarter > 4):
            raise ValueError('Invalid quarter:'+ str(quarter))
        
        if (year < MIN_YEAR_SUPPORTED) or (year > MAX_YEAR_SUPPORTED):
            raise ValueError('Invalid year:' + str(year))
        
        ww_per_q = wws_in_quarter(qtr=quarter, year=year+WW_YEAR_REAL_WORLD_YEAR)

        whole_q_wws =(quarter - 1 ) * ww_per_q
        
        partial_wws = int(quarter_treatment * ww_per_q )

        ww = whole_q_wws + partial_wws

        return (ww, year)

def delta_string_to_wws( delta_string ):
    """
    Converts a delta string expression into a relative # of workweeks
    
    Raises an exception if not a valid delta string.....

    (+ or - or =)(int or float)(W or WW or M or  Q or QS)
    """
    delta_string_copy = delta_string

    delta_string = delta_string.strip().upper()

    if delta_string.strip()[0] not in VALID_DELTA_STR:
        raise ValueError('Not a valid delta')

    equality = delta_string[0]
    
    delta_string = delta_string[1:].strip()

    if equality == DELTA_EQ_STR:
        if len(delta_string) > 0:
            raise ValueError('No additional text allowed after equals sign:' + delta_string_copy)
        return 0

    if DELTA_WEEK_STR in delta_string:
        wws = int(delta_string.split(DELTA_WEEK_STR)[0])
    elif DELTA_MONTH_STR in delta_string:
        months = float(delta_string.split(DELTA_MONTH_STR)[0])
        wws = round(months * WW_PER_MONTH)
    elif DELTA_QTR_STR in delta_string:
        qtrs = float(delta_string.split(DELTA_QTR_STR)[0])
        wws = round(qtrs * WW_PER_QTR)
    else:
        raise ValueError('Error parsing delta string '+delta_string_copy)
    
    if equality == DELTA_NEG_STR:
        wws = wws * -1

    return wws


if __name__ == '__main__':
    TESTS = [DELTA_PLUS_STR+s for s in ['1WW', '1W','1.1M', '2.2M','1.1Q', '2.2Qs']]

    for i, ds in enumerate(TESTS):
        print(ds,' ------ ', delta_string_to_wws(delta_string=ds))
        print(ds.replace(DELTA_PLUS_STR, DELTA_NEG_STR), ' ------ ',\
            delta_string_to_wws(delta_string=ds.replace(DELTA_PLUS_STR, DELTA_NEG_STR)))
        print(' '*(i+1) + DELTA_EQ_STR+' '*(i+1), delta_string_to_wws(delta_string=' '*(i+1) + DELTA_EQ_STR+' '*(i+1)))

    
    
    
            
