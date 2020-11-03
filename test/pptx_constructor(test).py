import pptx_construct

import pandas as pd

from pptx.util import Inches, Pt, Cm
from utils.pptx_params import PrsParamsManager

df1 = pd.DataFrame([[1, 2, 3] for _ in range(3)], columns=['a', 'b', 'c'])
df2 = pd.DataFrame([[4, 5, 6] for _ in range(3)], columns=['d', 'e', 'f'])
df3 = pd.DataFrame([[7, 8, 9] for _ in range(3)], columns=['g', 'h', 'i'])

params = PrsParamsManager()

constructor = pptx_construct.PptxConstructor({'prs_width': Inches(13),
                                              "prs_height": Inches(7)})

chart_loc_1 = ("a", Cm(0), Cm(0), Cm(10), Cm(10))
uid1 = constructor.add_object(data=df1, object_type='chart', position=chart_loc_1,
                              object_format={"chart_type": "bar", 'colormap': params.color_map('sunshine')})

chart_loc_2 = ("rr", uid1, Cm(0), Cm(0), Cm(10), Cm(10))
uid2 = constructor.add_object(data=df2, object_type='table', position=chart_loc_2, slide_page=1)

text_loc_1 = ("rd", uid2, Cm(0), Cm(0), Cm(4), Cm(1))
constructor.add_object(data='new chart', object_type='text', position=text_loc_1, slide_page=1,)

chart_loc_3 = ("rb", 0, 0, 0.3, 0.3)
constructor.add_object(data=df3, object_type='chart', slide_page=2,
                       position=chart_loc_3)

constructor.pptx_execute()
constructor.pptx_save("../test_storage/test.pptx")
