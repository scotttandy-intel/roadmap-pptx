"""
update.py

Updates one excel spreadsheet (like the bronze doc ) with dates from 
a second spreadsheet (formatted like the PowerBI Altas Download doc)
"""

import sys
from os.path import splitext, join

import pandas as pd 
import click

import power_bi_gd as pb 


### Valid Variable names to check values passed from external file
### The values in the 'def roadmap' must be kept consistent with these values
BRONZE_DOC_PATH = 'bronze_doc_path'
GOLDEN_DOC_PATH  = 'golden_doc_path'
OUTPUT_PATH = 'output_path'
UPDATED_TEXT = '_upd'




@click.command()
@click.argument( BRONZE_DOC_PATH, type=click.Path(exists=True))
@click.argument( GOLDEN_DOC_PATH, type=click.Path(exists=True))
@click.option('-o', OUTPUT_PATH, type=click.Path(exists=False), help='Full path to updated output file)')
def update(bronze_doc_path, golden_doc_path, output_path ):
    if output_path is None:
        output_path = join(splitext(bronze_doc_path)[0]+UPDATED_TEXT+splitext(bronze_doc_path)[1])

    pb_df = pb.power_bi_to_df( excel_path=golden_doc_path)

    

if __name__ == '__main__':
    update()