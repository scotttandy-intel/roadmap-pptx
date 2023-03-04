"""
test_intel_ww.py

Tests for the intel_ww.py classes and helper functions
"""

import pytest

import pandas as pd
import datetime
from calendar import Calendar
import intel_ww as iw

def test_ww_in_yr_qtr():
    for year in range(0,iw.MAX_YEAR_SUPPORTED):
        add_up_wws = 0
        ww_in_year = iw.wws_in_year(2000+year)
        for qtr in range(1,5):
            ww_in_qtr =  iw.wws_in_quarter(qtr=qtr, year=2000+year)
            add_up_wws = add_up_wws + ww_in_qtr
        assert ww_in_year == add_up_wws

    return

def build_valid_string_many_ways_wws(ww, year):
    ww_strings = []
    ww_strings.append( (f"{ww}'{year:02}", ww, year) )
    ww_strings.append( (f"{ww}'{year:02}", ww, year) )
    ww_strings.append( (f"  {ww}'{year:02}", ww, year) )
    ww_strings.append( (f" {ww}'{year:02}", ww, year) )
    ww_strings.append( (f"  Ww{ww}'{year:02}  ", ww, year) )
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
                
                string_tups = build_valid_string_many_ways_wws(ww=ww, year=year)
                for ww_string, ww, year in string_tups:
                    my_ww = iw.WW_from_string(ww_string=ww_string)
                    assert my_ww is not None
                    assert my_ww.ww == ww
                    assert my_ww.year == year      
    return
def test_check_all_the_dates_ww():
    c = Calendar()
    for year in range(iw.MAX_YEAR_SUPPORTED):
        for month in range(1,13):
            for date in c.itermonthdates(year=2000+year, month=month):
                if date.month != month:
                    continue             
                pandas_datetime = pd.Timestamp(date)
                #print(pandas_datetime, pandas_datetime.weekofyear)
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

def build_valid_string_many_ways_qtr(qtr, year):
    qtr_strings = []
    qtr_strings.append( (f"Q{qtr}'{year}", qtr, year) )
    qtr_strings.append( (f"Q{qtr}'{year}", qtr, year) )
    qtr_strings.append( (f"  Q{qtr}'{year}", qtr, year) )
    qtr_strings.append( (f" Q{qtr}'{year}", qtr, year) )
    qtr_strings.append( (f"  Q{qtr}'{year}  ", qtr, year) )
    qtr_strings.append( (f"  Q{qtr}'{year}  ", qtr, year) )
    qtr_strings.append( (f"  Q{qtr:02}'{year:02}  ", qtr, year) )
    qtr_strings.append( (f"Q{qtr} ' {year}", qtr, year) )
    qtr_strings.append( (f"Q{qtr} '{year}", qtr, year) )
    qtr_strings.append( (f"Q {qtr}' {year}", qtr, year) )
    qtr_strings.append( (f"Q {qtr}' {year}", qtr, year) )
    qtr_strings.append( (f"  Q{qtr}'{year}  ", qtr, year) )
    qtr_strings.append( (f"  Q{qtr}'{year}  ", qtr, year) )
    qtr_strings.append( (f"  Q{qtr:02}'{year:02}  ", qtr, year) )
    qtr_strings.append( (f'Q{qtr:02}"{year:02}  ', qtr, year) )
    qtr_strings.append( (f'Q{qtr:02} : {year:02}  ', qtr, year) )
    return qtr_strings

def test_construct_all_qts():
    for year in range(0,iw.MAX_YEAR_SUPPORTED):
        for qtr in range(1,5):          
                string_tups = build_valid_string_many_ways_qtr(qtr=qtr, year=year)
                for qtr_string, ww, year in string_tups:
                    my_ww = iw.WW_from_string(ww_string=qtr_string)
                    #print(qtr_string, my_ww)
                    assert my_ww is not None
                    #assert my_ww.ww == ww
                    assert my_ww.year == year

def test_eq_lt():
    for i in range(100):
        ww = iw.random_WW(min_year = iw.MIN_YEAR_SUPPORTED, max_year=iw.MAX_YEAR_SUPPORTED)
        eq_ww = iw.WW(ww.ww, ww.year)
        assert ww == eq_ww

        lt_ww = ww.subtract_wws(10)
        assert lt_ww < eq_ww

        gt_ww = ww.add_wws(10)
        assert gt_ww > ww
    return

def test_datetime_round_trip():
    c = Calendar()
    for year in range(iw.MAX_YEAR_SUPPORTED):
        for month in range(1,13):
            for date in c.itermonthdates(year=2000+year, month=month):
                ### This needs to be in here because of a funny 
                ### feature of the Calendar module....
                if date.month != month:
                    continue     
                from_dt = iw.WW_from_date(date_field=date)
                to_dt = from_dt.to_datetime()
                assert from_dt.contains(date)
                assert abs((date-to_dt).days < 8)

def test_day_of_the_year():
    c = Calendar()
    for year in range(1, iw.MAX_YEAR_SUPPORTED):
        day_count = 0
        for month in range(1,13):
            for date in c.itermonthdates(year=2000+year, month=month):
                ### This needs to be in here because of a funny 
                ### feature of the Calendar module....
                if date.month != month:
                    continue  
                day_count = day_count + 1
                ww = iw.WW_from_date(date_field=date)
                day_of_year = ww.day_of_the_year()
                
                #### Special case for wrapping at start and end of year
                #print(day_count, day_of_year)
                if day_of_year < 7 and day_count > 350:
                    continue
                if day_of_year > 350 and day_count < 7:
                    continue
                assert abs(day_count - day_of_year) < 7
    return



if __name__ == '__main__':
    test_construct_all_wws()
    test_check_all_the_dates_ww()
    test_construct_all_qts()
    test_eq_lt()
    test_datetime_round_trip()
    test_ww_in_yr_qtr()
    test_day_of_the_year()



