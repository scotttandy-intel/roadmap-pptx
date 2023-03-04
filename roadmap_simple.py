import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from collections import  OrderedDict

#### FILES #####
GOLDEN_DOC_XLS_PATH = \
"/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/ElastiScenariosGoldenDoc.xlsx"
ROADMAP_HELPER_EXCEL = \
        '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/ElastiScenariosHelperDoc.xlsx'
ROADMAP_CONFIG_EXCEL = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/ElastiConfig.xlsx'
ROADMAP_TEMPLATE = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ScottsWorkbench/BlankIAGSTemplate.pptx'
ROADMAP_OUT = '/Users/scotttan/Intel Corporation/Graphics Golden Doc Development - RoadmapPPTGeneration/ElastiScenarios/IAGS_OUT.pptx'

### Constant configuraiton - will need to be paramters
START_WW = iw.WW(27,20)
END_WW = iw.WW(51,23)
ROADMAP_TOP_CM = 1.5
ROADMAP_HEIGHT_CM =16.25
INPUT_SLIDE_INDEX = 0

pptx = render_roadmap_from_paths(golden_doc_path=GOLDEN_DOC_XLS_PATH, roadmap_helper_path=ROADMAP_HELPER_EXCEL,\
    roadmap_config_path=ROADMAP_CONFIG_EXCEL, 
    roadmap_template_path=ROADMAP_TEMPLATE, \
    start_ww=START_WW, end_ww=END_WW,\
    roadmap_top_cm=ROADMAP_TOP_CM, roadmap_height_cm=ROADMAP_HEIGHT_CM,\
            input_slide_index=INPUT_SLIDE_INDEX)

pptx.save(ROADMAP_OUT)