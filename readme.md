# AutoPPTX
AutoPPTX is a wrapper based on python-pptx(https://python-pptx.readthedocs.io/en/latest/).  
The package provides an API that allows user to create presentation with python-pptx more easily.
  
## I. Modules
Since the purpose of AutoPPTX is to minimize the effort to utilize python-pptx, it reduces the flexibility
to manipulate the entire presentation file. The API provides three ways to produce a presentation file:   
* PPTX_Construct: a module with more flexibility that is designed to produce a presentation file with  
* Data2Chart: a specific module to create slides with only charts  
* Data2Table: a specific module to create slides with only tables  
* Data2Text: a specific module to create slides with only text box  
standard structure and layout (typically a weekly or daily report with fixed layout but varied data source).  
 
## II. Example 
### 1. Creating charts on pptx with Pandas
    
#### Data2Chart
    pptx = ChartCreator()
    pptx.add_slide("main_page") 
    pptx.add_slide("slide_2")  
    pptx.add_slide("slide_3")  
    # df is a pandas dataFrame
    pptx.add_chart(data=df, slide_id="main_page", chart_type='line')  
    pptx.add_chart(data=df, slide_id="slide_2", chart_type='bar')  
    pptx.add_chart(data=df, slide_id="slide_3", chart_type='pie')
    pptx.save("new.pptx")
    
#### PPTX_Constructor
    constructor = pptx_construct.PptxConstructor({'prs_width': Inches(13), "prs_height": Inches(7)})
    # absolute position
    loc_1 = ("a", Inches(0), Inches(0), Inches(6), Inches(6)) 
    uid1 = constructor.add_object(data=df1, object_type='chart', position=loc_1,
                                  object_format={"chart_type": "bar", 'colormap': color_map('sunshine')})
    
    # relative position to boundary
    loc_2 = ("rb", 0, 0, 1, 1)
    constructor.add_object(data=df2, object_type='chart', position=loc_2, slide_page=1,
                           object_format={"chart_type": "line", "font_size": Pt(20)})
    
    # relative position to other object(given uid)
    loc_3 = ("ro", uid1, Inches(0), Inches(0), Inches(6), Inches(6))
    constructor.add_object(data=df3, object_type='chart', slide_page=10,
                           position=loc_3)
    constructor.pptx_execute()
    constructor.pptx_save()

 ___
Modified under The MIT License (MIT) Copyright (c) 2013 Steve Canny(https://github.com/scanny)