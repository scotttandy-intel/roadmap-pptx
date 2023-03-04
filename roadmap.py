"""
roadmap.py
Command line front end for the roadmap rendering package
Scott Tandy
Copyright Intel Corporation 2020
"""
import sys

import click
from pptx import Presentation

import intel_roadmap as ir
import intel_ww as iw
import roadmap_config as rc
import roadmap_table as rt
import roadmap_helper as rh
import roadmap_pptx as rp
import render_roadmap as rr

### Valid Variable names to check values passed from external file
### The values in the 'def roadmap' must be kept consistent with these values
ROAMDAP_CONFIG_PATH = 'roadmap_config_path'
GOLDEN_DOC_PATH = 'golden_doc_path'
CONCEPT_DOC_PATH = 'concept_doc_path'
ROADMAP_HELPER_PATH = 'roadmap_helper_path'
ROADMAP_TEMPLATE_PATH = 'roadmap_template_path'
OUTPUT_SIDE_PATH = 'output_side_path'
START_WW = 'start_ww'
END_WW = 'end_ww'
TITLE_TEXT = 'title_text'
ALIGN_ZERO = 'align_zero'
DO_NOT_ALIGN_ZERO = 'do_not_align_zero'

VALID_PARMETERS = [ROAMDAP_CONFIG_PATH,
    GOLDEN_DOC_PATH,
    ROADMAP_HELPER_PATH,
    ROADMAP_TEMPLATE_PATH,
    OUTPUT_SIDE_PATH,
    START_WW,
    END_WW,
    TITLE_TEXT,
    ALIGN_ZERO,
    DO_NOT_ALIGN_ZERO,]

TRUE_FALSE_FLAGS = [ALIGN_ZERO, DO_NOT_ALIGN_ZERO]

@click.command()
@click.argument( ROAMDAP_CONFIG_PATH, type=click.Path(exists=True))
@click.option('-g', GOLDEN_DOC_PATH, help='Full path to roadmap table (golden_doc_path)', \
    type=click.Path(exists=True), default=None)
@click.option('-rh', ROADMAP_HELPER_PATH,help='Optional path roadmap helper file (roadmap_helper_path)', \
    type=click.Path(exists=True), default=None)
@click.option('-t', ROADMAP_TEMPLATE_PATH, help='Full path to PowerPoint input file (roadmap_template_path)', \
    type=click.Path(exists=True), default=None)
@click.option('-o', OUTPUT_SIDE_PATH, help='Full path to PowerPoint output file (output_side_path)',\
    type=click.Path(writable=True), default=None)
@click.option('-s','--'+START_WW, default=None, help="start ww in the format of ww'yy (start_ww)")
@click.option('-e','--'+END_WW, default=None, help="end ww in the format of ww'yy (end_ww)")
@click.option('-t', TITLE_TEXT, default='Roadmap', help="Roadmap Title Text (title_text)")
@click.option('-az','--alignzero', ALIGN_ZERO, is_flag=True, default=False,\
     help="Set to true to normalize all first milestones to 01'00)")
@click.option('-daz','--donotalignzero', DO_NOT_ALIGN_ZERO, is_flag=True, default=False,\
     help="Set to true to override -a/--alignzero)")

def roadmap(golden_doc_path, roadmap_helper_path, roadmap_config_path, \
    roadmap_template_path, output_side_path, start_ww, end_ww,\
        title_text, align_zero, do_not_align_zero ):
    
    command_dict = rc.RoadmapConfig(path_to_roadmap_config_excel=roadmap_config_path).commands

    ## golden_doc_path is provided by the user and validity checked by the click package

    ### User provided command line values ovverride the config file values
    if golden_doc_path is None:
        if GOLDEN_DOC_PATH in command_dict.keys() and \
            command_dict[GOLDEN_DOC_PATH] != '' :
            golden_doc_path = command_dict[GOLDEN_DOC_PATH]

    if roadmap_helper_path is None:
        if ROADMAP_HELPER_PATH in command_dict.keys() and \
            command_dict[ROADMAP_HELPER_PATH] != '':
            roadmap_helper_path = command_dict[ROADMAP_HELPER_PATH]
    
    if roadmap_template_path is None:
        if ROADMAP_TEMPLATE_PATH in command_dict.keys() and \
            command_dict[ROADMAP_TEMPLATE_PATH] != '':
            roadmap_template_path = command_dict[ROADMAP_TEMPLATE_PATH]
    
    if output_side_path is None:
        if OUTPUT_SIDE_PATH in command_dict.keys() and \
            command_dict[OUTPUT_SIDE_PATH] != '':
            output_side_path = command_dict[OUTPUT_SIDE_PATH]
    
    if start_ww is None:
        if START_WW in command_dict.keys() and \
            command_dict[START_WW] != '':
            start_ww = command_dict[START_WW]
    
    if end_ww is None:
        if END_WW in command_dict.keys() and \
            command_dict[END_WW] != '':
            end_ww = command_dict[END_WW]
    
    if align_zero is False:
        if ALIGN_ZERO in command_dict.keys() and \
            command_dict[ALIGN_ZERO] != '':
            align_zero= (command_dict[ALIGN_ZERO].upper() == rc.TRUE_VALUE_TEXT.upper())
    
    ### You should probably never set do_not_align_zero in the roadmap config
    ### command tab, since its only purpose is to allow a command line ride override of
    ## the align_zero setting set in the configuration file.
    if do_not_align_zero is False:
        if DO_NOT_ALIGN_ZERO in command_dict.keys() and \
            command_dict[DO_NOT_ALIGN_ZERO] != '':
            do_not_align_zero = (command_dict[DO_NOT_ALIGN_ZERO].upper() \
                == rc.TRUE_VALUE_TEXT.upper())

    ### Again, this enables the user to ovverride the config file align_zero = True flag
    if do_not_align_zero is True:
        align_zero = False

    print('golden_doc_path',golden_doc_path)
    print('roadmap_helper_path',roadmap_helper_path)
    print('roadmap_config_path',roadmap_config_path)
    print('roadmap_template_path',roadmap_template_path)
    print('output_side_path',output_side_path)
    print('start_ww',start_ww)
    print('end_ww',end_ww)
    print('title_text',title_text)
    print('align_zero',align_zero)
    print('do_not_align_zero',do_not_align_zero)
    print('-------------------------------------------')

    start_ww_ww = iw.WW_from_string(start_ww)
    if start_ww_ww is None:
        sys.exit('Invalid start_ww date format:'+start_ww)
    
    end_ww_ww = iw.WW_from_string(end_ww)
    if end_ww_ww is None:
        sys.exit('Invalid end_ww date format:'+end_ww)
    
    print(start_ww_ww, end_ww_ww)

    pptx = rr.render_roadmap_from_paths(\
        golden_doc_path=golden_doc_path, 
        roadmap_helper_path=roadmap_helper_path,\
        roadmap_config_path=roadmap_config_path,\
        roadmap_template_path=roadmap_template_path, 
        start_ww=start_ww_ww, end_ww=end_ww_ww,
        roadmap_title=None, \
        roadmap_top_cm=1.5, input_slide_index=0, align_zero=align_zero )
    
    pptx.save(output_side_path)


if __name__ == '__main__':
    roadmap()