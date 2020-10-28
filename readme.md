# AutoPPTX
AutoPPTX is a wrapper based on python-pptx(https://python-pptx.readthedocs.io/en/latest/).  
The package provides an API that allows user to create presentation with python-pptx more easily.
  
## I. Main functions
Since the purpose of AutoPPTX is to minimize the effort to utilize python-pptx, it reduces the flexibility
to manipulate the entire presentation file. The API provides three ways to produce a presentation file:   
(1) Data2Chart: a specific module to create slides with only charts  
(2) Data2Table: a specific module to create slides with only tables  
(3) PPTX_Construct: a module with more flexibility that is designed to produce a presentation file with
standard structure and layout (typically a weekly or daily report with fixed layout but varied data source).  


## II. Other tools
Besides, two tools that may be useful for some scenarios will be available in the future:  
(1) chart data preprocessor: handle multiple types of data (e.g. pandas Dataframe or Series),
 transform them into proper format for adding charts or others.  
(2) layout scanner:  scan the existing presentation file, and analyze the layout and structure so the user
can manipulate the data source of each chart or table while retaining the original presentation structure.  
 
## III. Example 
### 1. Data2Chart: Creating charts on pptx with Pandas
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
Modified under The MIT License (MIT)

The MIT License (MIT)
Copyright (c) 2013 Steve Canny, https://github.com/scanny

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
