"""
intel_ww.py

Set of classes and functions which wrap the isoweek open source
WorkWeek management library

NOTE: Like everything else - work weeks are complicated.

NOTE: This should be converted into a true package and related package installer
      as it is

Author: Scott Tandy   27-Sep-20

Copyright Intel Corporation 2020
"""

WORK_WEEK = "work_week"
MONTH = "month"
QUARTER = "quarter"

DAYS_PER_YEAR = 365  ##Ignoring leap year for most WW Applications
MONTHS_PER_YEAR = 12
MONTHS_PER_QTR = 3
WW_PER_YEAR_USUALLY = 52
WW_MAX_WW = 53
DAYS_PER_WW = 7
QTR_PER_YEAR = 4
WW_PER_QTR = 13
VALID_MAX_WW = 53
WW_PER_MONTH = float(WW_PER_YEAR_USUALLY) / float(MONTHS_PER_YEAR)

MIN_YEAR_SUPPORTED = 0
MAX_YEAR_SUPPORTED = 50
WW_YEAR_REAL_WORLD_YEAR = 2000

WW_PREFIX = "WW"
WW_PLURAL = "WWs"

from collections import OrderedDict
import datetime
import math
import random
import re

import pandas as pd

from calendar import Calendar
from isoweek import Week

DELTA_PLUS_STR = "+"
DELTA_NEG_STR = "-"
DELTA_EQ_STR = "="

VALID_DELTA_STR = [DELTA_NEG_STR, DELTA_EQ_STR, DELTA_PLUS_STR]

DELTA_QTR_STR = "Q"
DELTA_MONTH_STR = "M"
DELTA_WEEK_STR = "W"

GENERIC_WW_TO_Q = [4, 4, 5, 4, 4, 5, 4, 4, 5, 4, 4, 5]

def wws_in_year(year):
    """
    year: Full AD year (eg. 2001)

    Returns: Number of workweeks in the given year (52)
    """
    return len(list(Week.weeks_of_year(year)))

def num_wws_in_quarter(qtr, year):
    """
    Returns: Number of workweeks in the given quarter (13)

    In a 53 WW year - we add the extra ww in the last Q
    """
    if qtr != 4:
        return WW_PER_QTR

    if wws_in_year(year) == WW_PER_YEAR_USUALLY:
        return WW_PER_QTR

    return WW_PER_QTR + 1

def wws_to_text(wws):
    """
    Returns: '1WW', '2WWs'
    """

    if abs(wws) <= 1:
        suffix = WW_PREFIX
    else:
        suffix = WW_PLURAL

    return f"{wws}" + suffix

def wws_to_days(wws):
    """
    wws: int - number of workweeks

    Returns: number of days in wws (14)
    """
    return wws * DAYS_PER_WW

def wws_to_days_text(wws):
    """
    wws: int - number of workweeks

    Returns: number of days in wws ('1day', '7days')
    """
    days = wws_to_days(wws=wws)

    if abs(days) <= 1:
        suffix = "day"
    else:
        suffix = "days"

    return f"{days}" + suffix

def wws_to_qtr(wws, ndigits=1):
    """
    wws: int - number of workweeks

    Returns: number of quarters in wws (0.1, 1.0)
    """
    return round(float(wws) / float(WW_PER_QTR), ndigits)

def wws_to_qtr_text(wws, ndigits=1):
    """
    wws: int - number of workweeks

    Returns: number of quarters in wws ('0.1Q', '1.0Q', '1.1Qs')
    """
    qtrs = wws_to_qtr(wws=wws, ndigits=ndigits)

    if abs(qtrs) <= 1:
        suffix = "Q"
    else:
        suffix = "Qs"

    if ndigits == 0:
        return f"{qtrs:.0f}" + suffix
    if ndigits == 1:
        return f"{qtrs:.1f}" + suffix
    if ndigits == 2:
        return f"{qtrs:.2f}" + suffix

    return f"{qtrs:.5f}" + suffix

def wws_to_years(wws, ndigits=1):
    """
    wws: int - number of workweeks

    Returns: number of years in wws (1.0)
    """
    return round(float(wws) / float(WW_PER_YEAR_USUALLY), ndigits)

def wws_to_years_text(wws, ndigits=1):
    """
    wws: int - number of workweeks

    Returns: number of years in wws ('0.1Y', '1.1Ys')
    """

    years = wws_to_years(wws=wws, ndigits=ndigits)

    if abs(years) <= 1:
        suffix = "Y"
    else:
        suffix = "Ys"

    if ndigits == 0:
        return f"{years:.0f}" + suffix
    if ndigits == 1:
        return f"{years:.1f}" + suffix
    if ndigits == 2:
        return f"{years:.2f}" + suffix

    return f"{years:.5f}" + suffix

def qqyys_in_year(year):
    """
    Returns: list of QQ'YY in the given year
    """

    # use right 2 digits of year, eg '21'
    yy = str(year)[-2:]

    qqyys = []
    for q in [1,2,3,4]:
        qqyys.append(f"Q{q}'{yy}") # eg Q4'21

    return qqyys


class WW:
    """
    Does WW calculations
    """

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

    def __init__(self, ww=None, year=None, date_time=None):
        """
        Can supply ww/year or date_time
        """
        if ww is None and year is None and date_time is None:
            raise ValueError("Must supply ww/year or date_time to constructor")

        self.ww = ww
        self.year = year
        self._date_time = date_time
        self._isoweek = None

        if self._date_time is not None:
            self._isoweek = Week.withdate(self._date_time)
            self.ww = self._isoweek.week
            self.year = self._isoweek.year % 100
        else:
            self._isoweek = Week(year=year + WW_YEAR_REAL_WORLD_YEAR, week=ww)
        return

    def __str__(self):
        return WW.to_wwyyyy(self)

    def __eq__(self, other):
        if other is None:
            return False

        return self._isoweek == other._isoweek

    def __lt__(self, other):
        return self._isoweek < other._isoweek

    def __le__(self, other):
        return self._isoweek <= other._isoweek

    def __sub__(self, other):
        """
        Returns the difference in isoweeks, not a WW
        """
        if isinstance(other, int):
            return self.add_wws(-1 * other)

        if isinstance(other, float):
            return self.add_wws(-1 * int(other))

        return self._isoweek - other._isoweek

    def __add__(self, other):
        """
        Create behavior to adding ints or floats to a ww...
        """
        if isinstance(other, int):
            return self.add_wws(other)

        if isinstance(other, float):
            return self.add_wws(int(other))

        raise TypeError("Cannot add two intel WWs")

    def json_ok(self):
        """
        Friendlier format if you have to paste into different situations
        """
        return f"{self.ww:02}:{self.year:02}"

    def contains(self, date_time):
        return self._isoweek.contains(date_time)

    def AD_year(self):
        return self.year + WW_YEAR_REAL_WORLD_YEAR

    def to_datetime(self, day_dot=1):
        """
        Leverages isoweek
        """
        return self._isoweek.day(day_dot + 1)

    def to_wwyyyy(self):
        """
        str(intel_ww) == "52'2021"

        Returns: "ww'yyyy", "52'2021"
        """
        if self.ww < 10:
            return f" {self.ww}'{self.year:02}"

        return f"{self.ww:02}'{self.year:02}"
        # return str(self.ww)+"'"+str(self.year)

    def to_wwyy(self):
        """
        Returns ww'yy. 52'20
        """
        # truncate year to last 2 digits
        out = str(self.ww) + "'" + str(self.year)[-2:]
        return out

    def day_of_the_year(self, work_day=MONDAY):
        """
        The first day of WW1 is the "first day"

        You should use this very carefully - because the ww01'yy
        might actually be in yy-1 - and this is true on the other end.

        This function ends up returning the day of the year that the WW is actually IN
        """
        day_from_start_td = self.to_datetime(day_dot=work_day) - datetime.date(
            year=self.year + 2000, month=1, day=1
        )

        return day_from_start_td.days

    def month_of_year(self):
        return self.to_datetime().month

    def qtr_of_year(self):
        if self.ww < WW.Q2_WW_START:
            return 1
        if self.ww < WW.Q3_WW_START:
            return 2
        if self.ww < WW.Q4_WW_START:
            return 3
        ### Doesn't consider the 53 WW year - but accidentally works OK...
        return 4

    def qtr_year_str(self, delimiter='-'):
        """
        Formatted quarter of year

        qtr_year_str() == "Q1-21"
        qtr_year_str("'") == "Q1'21"
        """
        return f"Q{self.qtr_of_year()}{delimiter}{self.year:02d}"

    def add_wws(self, wws):
        return WW_from_isoweek(self._isoweek + wws)

    def subtract_wws(self, wws):
        return self.add_wws(wws * -1)

    def ww_delta_from(self, ww):
        return ww._isoweek - self._isoweek

    def quarter_delta_from(self, ww):
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
            self_delta = QTR_PER_YEAR - self_qtr
            ww_delta = ww_qtr
            years = ww.year - self.year - 1
            return self_delta + ww_delta + years * QTR_PER_YEAR

        self_delta = self_qtr
        ww_delta = QTR_PER_YEAR - ww_qtr
        years = self.year - ww.year - 1
        return -1 * (self_delta + ww_delta + years * QTR_PER_YEAR)

def WW_from_date(date_field):
    return WW(date_time=date_field)

def WW_from_date_string(date_string):
    '''
    date_string: yyyy-mm-dd format '2000-12-31'
    '''
    dt = datetime.datetime.strptime(date_string, '%Y-%m-%d')
    return WW_from_date(dt)

def WW_from_isoweek(iso_week, dot_day=1):
    date_time = iso_week.day(dot_day - 1)
    return WW_from_date(date_field=date_time)

def WW_from_string(ww_string, quarter_treatment=0.5):
    """
    ww_string -  WWww:yy, Qn'yy, "50'21", "50:21"
    quarter_treatment = if it is in a "quarter" where to place it in the quarter
                        (0.5 is half way through the quarter, 0 is at the start)
    """
    ## Handle if it is already a WW and YY
    try:
        ww, year = ww_text_to_tuple(ww_string)
        year = abs(year) % 100
        return WW(ww=ww, year=year)
    except:
        pass

    ## Handle if it is in Qn'YY
    try:
        ww, year = quarter_text_to_ww_year_tuple(
            ww_string, quarter_treatment=quarter_treatment
        )
        year = abs(year) % 100
        return WW(ww=ww, year=year)
    except:
        pass

    return None

def wwyys_in_wwyy_span(wwyy_start_str, wwyy_end_str):
    """
    ww_start: "1'20"
    ww_end: "14'20"

    Returns: Returns a list of wwyys (as str) that are in the span (inclusive)
    """
    ww_start = WW_from_string(wwyy_start_str)
    ww_end = WW_from_string(wwyy_end_str)

    # add 1 to include the final week
    total_weeks = ww_start.ww_delta_from(ww_end) + 1
    wwyys = []
    for i in range(total_weeks):
        wwyy = ww_start.add_wws(i).to_wwyy()
        wwyys.append(wwyy)

    # return list of wwyy strings
    return wwyys

def qqyys_in_year_span(start_year, end_year):
    """
    start_year: 2021
    end_year: 2024

    Returns: Returns a list of qqyys (as str) that are in the span (inclusive)
    """
    qqyys = []

    for year in range(start_year, end_year+1):
        qqyys  = qqyys + qqyys_in_year(year)

    # return list of qqyy strings
    return qqyys

def qqyys_in_wwyy_span(wwyy_start_str, wwyy_end_str):
    """
    ww_start: "1'20"
    ww_end: "14'20"

    Returns: Returns a list of qqyys (as str) that are in the ww span (inclusive)
    """
    # convert wwyy string to WW objects
    ww_start = WW_from_string(wwyy_start_str)
    ww_end = WW_from_string(wwyy_end_str)

    # convert WW objects to QQ objects
    qqyy_start_str = ww_start.qtr_year_str(delimiter="'")
    qqyy_end_str = ww_end.qtr_year_str(delimiter="'")
    qq_start = QQ(qqyy_start_str)
    qq_end = QQ(qqyy_end_str)

    # do until qqyy == qqyy_end_str
    qq:QQ = qq_start

    qqyys = []
    # qqyys = [ qq.to_qqyy() ]
    # qq = qq.add_1qtr()
    #while qq.to_qqyy() != qqyy_end_str:
    while True:
        qqyys.append(qq.to_qqyy())
        if qq.to_qqyy() == qqyy_end_str:
            break
        qq = qq.add_1qtr()

    # # add 1 to include the final week
    # total_weeks = ww_start.ww_delta_from(ww_end) + 1
    # wwyys = []
    # for i in range(total_weeks):
    #     wwyy = ww_start.add_wws(i).to_wwyy()
    #     wwyys.append(wwyy)

    # # return list of wwyy strings
    # return wwyys

    return qqyys

def is_valid_WW_string(ww_string):
    try:
        ww = WW_from_string(ww_string=ww_string)
    except:
        return False
    if ww is None:
        return False

    return True

def sort_WW_text_list(ww_text_list):
    """
    ww_test_list list(Valid_WW_Strings)
    """
    ww_tups = [(WW_from_string(s), s) for s in ww_text_list]
    return [t[1] for t in sorted(ww_tups)]

def random_WW(min_year, max_year):
    return WW(ww=random.randint(1, 52), year=random.randint(min_year, max_year))

def ww_text_to_tuple(ww_text):
    """
    Takes string in forms: "ww52 2000", WW25'20, ww25'20

    Returns: tuple (int work_week, int year)
    """
    SPLIT = "'"
    ALT_SPLITS = [
        "`",
        '"',
        ":",
    ]

    copy_text = ww_text

    # handle "ww52 2000", "WW52 2000"
    if re.match("^[wW]{2}[0-9]{2} [0-9]{4}$", ww_text):
        ww = int(ww_text[2:4])
        year = int(ww_text[-2:])
        return (ww, year)

    ww_text = ww_text.strip().upper()
    ww_text = ww_text.replace("W", "")

    ### Switch to standard split character
    for a_s in ALT_SPLITS:
        ww_text = ww_text.replace(a_s, SPLIT)

    ### Hack to remove (A0) from the end of the WW data
    if "(" in ww_text:
        ww_text = ww_text.split("(")[0]

    try:
        ww = int(ww_text.split(SPLIT)[0].strip())
        year = int(ww_text.split(SPLIT)[1].strip())

    except:
        raise ValueError("Bad format:" + copy_text + ":")

    if (ww < 1) or (ww > WW_MAX_WW):
        raise ValueError("Invalid workweek:" + str(ww))

    if (year < MIN_YEAR_SUPPORTED) or (year > MAX_YEAR_SUPPORTED):
        raise ValueError("Invalid year:" + str(year))

    return (ww, year)

def quarter_text_to_ww_year_tuple(q_text, quarter_treatment=0.5):
    """
    Takes string in form Q1'20 and returns
    (int work_week, int year)
    """

    if quarter_treatment < 0.0 or quarter_treatment > 1.0:
        raise ValueError("Quarter treatment must be between 0.0 and 1.0")

    SPLIT = "'"
    ALT_SPLITS = ["`", '"', ":", "-"]

    copy_text = q_text
    ### Switch to standard split character

    for a_s in ALT_SPLITS:
        q_text = q_text.replace(a_s, SPLIT)

    q_text = q_text.replace("q", "")
    q_text = q_text.replace("Q", "")

    try:
        quarter = int(q_text.split(SPLIT)[0].strip())
        year = int(q_text.split(SPLIT)[1].strip())
    except:
        raise ValueError("Bad format:" + copy_text + ":")

    if (quarter < 1) or (quarter > 4):
        raise ValueError("Invalid quarter:" + str(quarter))

    if (year < MIN_YEAR_SUPPORTED) or (year > MAX_YEAR_SUPPORTED):
        raise ValueError("Invalid year:" + str(year))

    if quarter_treatment == 0:
        if quarter == 1:
            ww = 1
        elif quarter == 2:
            ww = 14
        elif quarter == 3:
            ww = 27
        else:
            ww = 40
    else:
        ww_per_q = num_wws_in_quarter(qtr=quarter, year=year + WW_YEAR_REAL_WORLD_YEAR)

        whole_q_wws = (quarter - 1) * WW_PER_QTR

        partial_wws = int(quarter_treatment * ww_per_q)

        ww = whole_q_wws + partial_wws

    return (ww, year)

def delta_string_to_wws(delta_string):
    """
    Converts a delta string expression into a relative # of workweeks

    Raises an exception if not a valid delta string.....
    (+ or - or =)(int or float)(W or WW or M or  Q or QS)
    """
    delta_string_copy = delta_string

    delta_string = delta_string.strip().upper()

    if delta_string.strip()[0] not in VALID_DELTA_STR:
        raise ValueError("Not a valid delta")

    equality = delta_string[0]

    delta_string = delta_string[1:].strip()

    if equality == DELTA_EQ_STR:
        if len(delta_string) > 0:
            raise ValueError(
                "No additional text allowed after equals sign:" + delta_string_copy
            )
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
        raise ValueError("Error parsing delta string " + delta_string_copy)

    if equality == DELTA_NEG_STR:
        wws = wws * -1

    return wws

def dict_to_milestones(ms_dict, start_ww_str):
    """
    Helper function to turn milestone dicts
    ms_dict - {'milestone': int_delta_from_start_ww_str}
    """
    start_WW = WW_from_string(start_ww_str)
    milestones = OrderedDict()
    for milestone, delta in ms_dict.items():
        milestones[milestone] = start_WW.add_wws(delta)
    return milestones

def to_const(s):
    """
    Workbook helper function for creating Constant Code

    use with code:
    gb = df.groupby(by=[SOCS]).sum()
    gb = gb.loc[gb["Q1'21"] != 0.0]
    for name in list(gb.index):
        print(f"{to_const(name)}='{name}'")
    """
    return s.replace(" ", "_").upper().replace("-", "_").replace("''", "_")

class QQ:
    """
    class to handle Intel Quarters
    """

    qtr = None  # 4 (int)
    year = None # 2021 (int)

    def __init__(self, qqyy:str):
        """
        qqyy: "Q4'21" string
        """
        self.qtr = int(qqyy[1:2]) # 2nd char
        self.year = int(f"20{qqyy[-2:]}") # last 2 chars

    def __str__(self):
        return self.to_qqyy()

    def to_qqyy(self):
        """
        Returns: "Q4'21" string
        """
        s = f"Q{self.qtr}'{str(self.year)[-2:]}"
        return s

    def add_1qtr(self):
        """
        Returns: new QQ object 1 qtr later

        TODO improve to add any number of qtrs
        """
        new_qq = QQ(self.to_qqyy())

        # add 1 qtr
        new_qq.qtr += 1

        # handle special case
        if new_qq.qtr > 4:
            new_qq.qtr = new_qq.qtr - 4
            new_qq.year += 1

        return new_qq

    def wwyys_in_qtr(self):
        """
        Returns: list of wwyy's in the quarter

        [ "1'21", "2'21", ..., "13'21" ]
        """

        qqyy = self.to_qqyy()
        begin_ww = WW_from_string(qqyy, quarter_treatment=0)
        num_wws_in_qtr = num_wws_in_quarter(self.qtr, self.year)

        wwyys = []
        current_ww = begin_ww

        for _ in range(num_wws_in_qtr):
            wwyys.append(current_ww.to_wwyy())
            current_ww = current_ww.add_wws(1)

        return wwyys

class WWnnn:
    """
    Special class of WW's that are in the format WWnnn, WW001 etc up to WW999

    (999 / 52  = ~20 years so it seems that three digits is the sweet spot)
    """

    def __init__(self, www):
        self.www = int(www)
        if self.www < 0 or self.www > 999:
            raise ValueError("WWnnn constructor value out of range:" + str(www))

    def __str__(self):
        return f"WW{self.www:03}"

    def __eq__(self, other):
        return self.www == other.www

    def __lt__(self, other):
        return self.www < other.www

    def __add__(self, other):
        if isinstance(other, WWnnn):
            return WWnnn(www=(self.www + other.www))

        return WWnnn(www=(self.www + other))

    def __sub__(self, other):
        if isinstance(other, WWnnn):
            return WWnnn(www=(self.www + other.www))

        return WWnnn(www=(self.www - other))

def WW000_from_string(ww_string):
    """
    ww_string -  ww000 or WWnnn
    """
    ww_string = ww_string.upper().strip()

    try:
        www = int(ww_string.split(WW_PREFIX)[1].strip())
    except:
        raise ValueError("Error converting " + ww_string + " to WWnnn object")

    return WWnnn(www=www)

def is_valid_WW000_string(ww_string):
    try:
        _ = WW000_from_string(ww_string=ww_string)
    except:
        return False

    return True

def is_validM000_string(m000_string):
    """
    Check and see if a string is in the fromat of M000
    """
    m000_string = m000_string.strip().lower()

    try:
        if len(m000_string) != 4:
            return False

        if m000_string[0] != "m":
            return False

        i = int(m000_string[1:])

        if i >= 0 and i <= 999:
            return True
    except:
        pass

    return False

if __name__ == "__main__":

    # quick tests

    # # test WWs from quarter
    # print(QQ("Q1'20").wwyys_in_qtr())
    # print(QQ("Q4'20").wwyys_in_qtr()) # 14 weeks due to ww53
    # print(QQ("Q1'21").wwyys_in_qtr())
    # print(QQ("Q4'21").wwyys_in_qtr())

    # test WWs from quarter
    print(WW_from_string("Q1'21", quarter_treatment=0))
    print(WW_from_string("Q1'21", quarter_treatment=0.5))
    print(WW_from_string("Q1'21", quarter_treatment=0.99)) # 13'21
    print(WW_from_string("Q1'21", quarter_treatment=1))    # 14'21

    print(WW_from_string("Q4'21", quarter_treatment=0))
    print(WW_from_string("Q4'21", quarter_treatment=0.5))
    print(WW_from_string("Q4'21", quarter_treatment=0.99)) # 52'21
    print(WW_from_string("Q4'21", quarter_treatment=1))    # 53'21

    print(WW_from_string("Q4'20", quarter_treatment=0))
    print(WW_from_string("Q4'20", quarter_treatment=0.5))  # 50'20
    print(WW_from_string("Q4'20", quarter_treatment=0.99)) # 56'21
    print(WW_from_string("Q4'20", quarter_treatment=1))    # 57'21

    print(qqyys_in_wwyy_span("1'20", "14'20"))
    print(qqyys_in_wwyy_span("1'20", "13'20"))
    print(qqyys_in_wwyy_span("1'20", "27'20"))

    qq = QQ("Q1'20"); print(qq.to_qqyy())
    qq = qq.add_1qtr(); print(qq.to_qqyy())
    qq = qq.add_1qtr(); print(qq.to_qqyy())
    qq = qq.add_1qtr(); print(qq.to_qqyy())
    qq = qq.add_1qtr(); print(qq.to_qqyy())
    qq = qq.add_1qtr(); print(qq.to_qqyy())

    qqyys_in_year(2021)

    ww = WW(year=2020, ww=52)
    print(ww.to_wwyy())
    print(ww.to_wwyyyy())
    print(ww)

    wwyys_in_wwyy_span("1'20", "14'20")


    TESTS = [DELTA_PLUS_STR + s for s in ["1WW", "1W", "1.1M", "2.2M", "1.1Q", "2.2Qs"]]

    for i, ds in enumerate(TESTS):
        print(ds, " ------ ", delta_string_to_wws(delta_string=ds))
        print(
            ds.replace(DELTA_PLUS_STR, DELTA_NEG_STR),
            " ------ ",
            delta_string_to_wws(delta_string=ds.replace(DELTA_PLUS_STR, DELTA_NEG_STR)),
        )
        print(
            " " * (i + 1) + DELTA_EQ_STR + " " * (i + 1),
            delta_string_to_wws(
                delta_string=" " * (i + 1) + DELTA_EQ_STR + " " * (i + 1)
            ),
        )
