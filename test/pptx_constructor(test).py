import pptx_construct

import pandas as pd

from pptx.util import Inches, Pt
from utils.pptx_params import color_map

df1 = pd.DataFrame([[1, 2, 3] for _ in range(3)], columns=['a', 'b', 'c'])
df2 = pd.DataFrame([[4, 5, 6] for _ in range(3)], columns=['d', 'e', 'f'])
df3 = pd.DataFrame([[7, 8, 9] for _ in range(3)], columns=['g', 'h', 'i'])

constructor = pptx_construct.PptxConstructor({'prs_width': Inches(13),
                                              "prs_height": Inches(7)})

chart_loc_1 = ("a", Inches(0), Inches(0), Inches(6), Inches(6))
uid1 = constructor.add_object(data=df1, object_type='chart', position=chart_loc_1,
                              object_format={"chart_type": "bar", 'colormap': color_map('sunshine')})

chart_loc_2 = ("ro", uid1, Inches(0), Inches(0), Inches(6), Inches(6))
constructor.add_object(data=df2, object_type='chart', position=chart_loc_2, slide_page=1,
                       object_format={"chart_type": "line", "font_size": Pt(20)})

chart_loc_3 = ("rb", 0, 0, 1, 1)
constructor.add_object(data=df3, object_type='chart', slide_page=10,
                       position=chart_loc_3)

constructor.pptx_execute()
constructor.pptx_save()
