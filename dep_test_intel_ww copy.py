"""
test_intel_ww.py

Tests for the inte_ww.py classes and helper functions
"""

import pytest

import pandas as pd
import datetime
from calendar import Calendar
import intel_ww as iw

def build_valid_string_many_ways(ww, year):
    ww_strings = []
    ww_strings.append( (f"{ww}'{year}", ww, year) )
    ww_strings.append( (f"{ww}'{year}", ww, year) )
    ww_strings.append( (f"  {ww}'{year}", ww, year) )
    ww_strings.append( (f" {ww}'{year}", ww, year) )
    ww_strings.append( (f"  Ww{ww}'{year}  ", ww, year) )
    ww_strings.append( (f"  Ww{ww}'{year}  ", ww, year) )
    ww_strings.append( (f"  Ww{ww:02}'{year:02}  ", ww, year) )
    ww_strings.append( (f"{ww} ' {year}", ww, year) )
    ww_strings.append( (f"WW{ww} '{year}", ww, year) )
    ww_strings.append( (f"ww {ww}' {year}", ww, year) )
    ww_strings.append( (f"Ww {ww}' {year}", ww, year) )
    ww_strings.append( (f"  Ww{ww}'{year}  ", ww, year) )
    ww_strings.append( (f"  Ww{ww}'{year}  ", ww, year) )
    ww_strings.append( (f"  Ww{ww:02}'{year:02}  ", ww, year) )
    ww_strings.append( (f'Ww{ww:02}"{year:02}  ', ww, year) )
    ww_strings.append( (f'Ww{ww:02} : {year:02}  ', ww, year) )
    return ww_strings
    
def test_construct_all_wws():
    for year in range(0,iw.MAX_YEAR_SUPPORTED):
        for ww in range(1,53):
                ### Basic WW constructor
                my_ww = iw.WW(ww=ww, year=year)
                assert my_ww is not None
                assert my_ww.ww == ww
                assert my_ww.year == year
                
                string_tups = build_valid_string_many_ways(ww=ww, year=year)
                for ww_string, ww, year in string_tups:
                    my_ww = iw.WW_from_string(ww_string=ww_string)
                    assert my_ww is not None
                    assert my_ww.ww == ww
                    assert my_ww.year == year
         
    return

def test_check_all_the_dates():
    c = Calendar()
    for year in range(iw.MAX_YEAR_SUPPORTED):
        for month in range(1,13):
            for date in c.itermonthdates(year=2000+year, month=month):
                if date.month != month:
                    continue
                 
                pandas_datetime = pd.Timestamp(date)
                print(pandas_datetime, pandas_datetime.weekofyear)
                if month==12 and pandas_datetime.weekofyear == 1:
                    year_to_check = year + 1
                
                elif month==1 and pandas_datetime.weekofyear == 52:
                    next_week = pandas_datetime + pd.Timedelta(weeks=1)
                    if (next_week.weekofyear != 53):
                        year_to_check = year - 1
                
                elif month == 1 and pandas_datetime.weekofyear == 53:
                        year_to_check = year - 1
  
                else:
                    year_to_check = year

                if year_to_check < 0:  ### Special case of reverse wrap to year 99
                    continue

                ww = iw.WW_from_date(date_field=date)

                assert ww.year == year_to_check
                assert ww.ww == pandas_datetime.weekofyear

                ww = iw.WW_from_date(date_field=pandas_datetime)
                assert ww.year == year_to_check
                assert ww.ww == pandas_datetime.weekofyear

"""
def test_basic_constructor_WW():

def test_ww_text_to_WW():
    for ww_text, ww, yy in WW_TEXT_TUPLES:

    print(WW(23,21).add_wws(20))
    print(WW(35,21).add_wws(26))

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
            print(ww, day, my_ww.day_of_the_year(work_day=day), my_ww.qtr_of_year())
    
"""

if __name__ == '__main__':
    test_construct_all_wws()
    test_check_all_the_dates()



