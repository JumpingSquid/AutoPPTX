from autoppt_util import pptx_construct

pptxconstructer = pptx_construct.PptxConstructor()
pptxconstructer._set_config_datapath("C:/Users/user/Desktop/p1876510037_emil.hristov_249794536.xlsx")

slide_config = pptxconstructer.slide_config_create(1)
chart_config = pptxconstructer.chart_config_create(category_col=["S1"], value_col=["S13"], type="stackbar", fillna="1")
slide_config["chart"] = slide_config["chart"] + [chart_config]
chart_config = pptxconstructer.chart_config_create(category_col=["S1"],
                                                   value_col=["Product2"],
                                                   type="hist",
                                                   fillna="1",
                                                   agg="percentage")
slide_config["chart"] = slide_config["chart"] + [chart_config]
pptxconstructer.slide_config_append(slide_config)

slide_config = pptxconstructer.slide_config_create(2)
for brand_id in [147, 138, 134]:
    chart_config = pptxconstructer.chart_config_create(category_col=["S1"],
                                                       value_col=["Q2_"+str(brand_id), "Q3_"+str(brand_id), "Q4"],
                                                       type="multiple_category_multiple_value_column_bar",
                                                       fillna="2",
                                                       pos_indicator=[1, 1, brand_id],
                                                       ignorena=False,
                                                       label_number_format="0.00")
    slide_config["chart"] = slide_config["chart"] + [chart_config]
pptxconstructer.slide_config_append(slide_config)

slide_config = pptxconstructer.slide_config_create(3)
chart_config = pptxconstructer.chart_config_create(category_col=["Q4"], type="pie", fillna="1")
slide_config["chart"] = slide_config["chart"] + [chart_config]
pptxconstructer.slide_config_append(slide_config)

pptxconstructer.pptx_save()
