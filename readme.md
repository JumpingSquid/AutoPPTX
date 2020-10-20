# AutoPPTX
AutoPPTX is a package based on python-pptx(https://python-pptx.readthedocs.io/en/latest/).  
The package provides an API for creating presentation charts or other objects more easily,
 especially the case that multiple slides with multiple objects needed to be created.
 
## Example 
### 1. Creating charts on pptx with Pandas
    pptx = ChartCreator()
    pptx.add_slide("main_page") 
    pptx.add_slide("slide_2")  
    pptx.add_slide("slide_3")  
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