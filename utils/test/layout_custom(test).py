from autoppt_util import pptx_construct
from autoppt_util.pptx_params import color_map
from autoppt_util.layout_scanner import LayoutScanner
import pandas as pd


hotel_name = {"bjkd": "北京凱迪克", "ncnc": "南昌國際", "szzh": "深圳中航城", "szhy": "深圳花園", "jxgz": "贛州國際"}
hotel = "szhy"
hotel = (hotel, hotel_name[hotel])
month = ("11", "nov")

filename = hotel[0] + "_" + month[0] + ".xlsx"
df = pd.read_excel("C:/Users/user/PycharmProjects/GS_analytics/data/project_2019/survey/"+filename)
col_map = pd.read_excel("C:/Users/user/PycharmProjects/GS_analytics/data/project_2019/survey/column_name_map.xlsx")
df.columns = col_map.iloc[:len(df.columns), :].column_english
column_name = df.columns
col_alias_dict = {r["column_english"]: r["column_alias"] for i, r in col_map.iterrows()}

df.loc[df.nps.isin(['9分', '10分（非常愿意）']), "nps"] = "promoter"
df.loc[df.nps.isin(['7分', '8分']), "nps"] = "passive"
df.loc[df.nps.isin(['0分（非常不愿意）', '1分', '2分', '3分', '4分', '5分', '6分']), "nps"] = "detractor"

pcor = pptx_construct.PptxConstructor()
pcor.set_datapath(df)

sc = pcor.slide_config_create(1)
sc["title"] = hotel[1]
sc["chart"].append(pcor.chart_config_create(category_col=["rating_overall"],
                                            type="pie",
                                            agg="percentage",
                                            label_number_format='0%',
                                            colormap=color_map("coldwarm"),
                                            charttitle="入住体验整体评价"))

sc["chart"].append(pcor.chart_config_create(category_col=["rating_competitive"],
                                            type="pie",
                                            agg="percentage",
                                            label_number_format='0%',
                                            colormap=color_map("sunshine"),
                                            charttitle="与同价位的酒店相比的评价"))

sc["chart"].append(pcor.chart_config_create(category_col=["revisit"],
                                            type="pie",
                                            agg="percentage",
                                            label_number_format='0%',
                                            colormap=color_map("sunshine"),
                                            charttitle="再次光临"))

sc["comment"] = pcor.comment_config_create(cat_col="rating_overall", col_alias="整体评价", comment_type="%max")
sc["custom_layout"] = LayoutScanner().chart_layout_export()
pcor.slide_config_append(sc)
pcor.pptx_execute()
pcor.pptx_save("C://Users/user/Desktop/custom_layout.pptx")
