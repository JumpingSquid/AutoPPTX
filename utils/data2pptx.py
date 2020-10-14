import pandas as pd
from pptx import Presentation
from pptx.util import Inches

from autoppt_util import pptx_layout, smart_config
from data_preprocessing import chart_data_preprocessor

configwizard = smart_config.PowerpointConfig()
config = configwizard.configwizard()

if config["dataframe"][-3:] == "csv":
    df = pd.read_csv(config["dataframe"])
    print(config["dataframe"], "is read")
elif config["dataframe"][-4:] == "xlsx":
    df = pd.read_excel(config["dataframe"])
    print(config["dataframe"], "is read")
else:
    raise ValueError("Only support csv or xlsx file, please check data input format")

prs = Presentation()
# slide size: 16:9
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
print("create presentation file")

page_count = 1
for slide_config in config["slide"]:
    print("slide " + str(page_count) + " is processed")
    data = chart_data_preprocessor.ChartDataProcessor(df, slide_config["data"])
    layoutmanager = pptx_layout.PrsLayoutManager(prs, slide_config, data)
    prs = layoutmanager.add_chart_on_slide()
    page_count += 1

prs.save('test.pptx')
print("presentation file is created successfully")
