# AutoPPTX
AutoPPTX is a wrapper based on python-pptx(https://python-pptx.readthedocs.io/en/latest/).  
The package provides an interface that allows user to create presentation with python-pptx more easily.
  
## I. Modules
Since the purpose of AutoPPTX is to minimize the effort to utilize python-pptx, it reduces the flexibility
to manipulate the entire presentation file. The API provides several ways to produce a presentation file:   
* PPTX_Construct: a module with more flexibility that is designed to produce a presentation file with standard
 structure and layout (typically a weekly or daily report with fixed layout but varied data source).  
* Data2Chart: a specific module to create slides with only charts  
* Data2Table: a specific module to create slides with only tables  
* Data2Text: a specific module to create slides with only text box  

 
## II. Example 
### 1. Creating charts on pptx with Pandas
    
#### PPTX_Constructor
    constructor = PptxConstructor({'prs_width': Inches(13), "prs_height": Inches(7)})
    constructor.add_object(data=df1, object_type='chart', object_format={"chart_type": "line"})
    constructor.add_object(data=df2, object_type='chart', slide_page=1)
    constructor.add_object(data=df3, object_type='chart', position=('a', 0, 0, Cm(13), Cm(13)))
    constructor.pptx_execute()
    constructor.pptx_save()
    
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

 ___
Modified under The MIT License (MIT) Copyright (c) 2013 Steve Canny(https://github.com/scanny)