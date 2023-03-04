"""
test_annotations.py

Tests for the annotations.py classes and helper functions
Scott Tandy
Copyright Intel Corporation 2020
"""

import pytest
import json

import annotations as an
SQ_TO_DQ_TUPELES = [("'",'"'), ("'ab':'cd'",'"ab":"cd"'),("{'ab':'cd'}",'{"ab":"cd"}'), ]

def test_sq_to_dq ():
    for t, a in SQ_TO_DQ_TUPELES:
        assert an.sq_to_dq(t) == a

TRUES_FOR_TO_BOOL = ['true',' true ', 'TRUE', ' TRue ', '  TruE   ']
FALSES_FOR_TO_BOOL = ['false', '  false ', 'False', '  FALSE  ', ' fAlSE ']
BADS_FOR_TO_BOOL = [True, False, 'asdfasd','notTrue','','kjlkjlkj',]

def test_text_to_bool():
    for t in TRUES_FOR_TO_BOOL:
        assert an.text_to_bool(t) == True
    for f in FALSES_FOR_TO_BOOL:
        assert an.text_to_bool(f) == False
    for b in BADS_FOR_TO_BOOL:
        with pytest.raises(ValueError):
            an.text_to_bool(b)


GOODS_FOR_RGB = [ ( '(0,0,0)', (0,0,0) ) ]
BADS_FOR_RGB = [ '', ' ', '(0,1)','asdfasdfasdf'] 
def test_text_rgb_triple():
    for t, a in GOODS_FOR_RGB:
        assert an.text_rgb_triple(t) == a
    
    for t in BADS_FOR_RGB:
        with pytest.raises(ValueError):
            an.text_rgb_triple(t)

GOODS_FOR_CELL_TO_J = [("'Hi there'",'Hi there'),("{'a':'b'}",{'a':'b'})]
BADS_FOR_CELL_TO_J = ['','garbage',"{'a':b}'"]
def test_cell_to_json():
    for t, a in GOODS_FOR_CELL_TO_J:
        assert an.cell_to_json(t) == a

    for t in BADS_FOR_CELL_TO_J:
        with pytest.raises(json.JSONDecodeError):
            an.cell_to_json(t)



if __name__ == '__main__':
    test_sq_to_dq()
    test_text_to_bool()
    test_text_rgb_triple()
    test_cell_to_json()
   