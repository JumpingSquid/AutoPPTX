"""
PptxConstructor is the user interface to create the presentation file.
The steps are described as follows:
(1) Specifying the layout of the presentation file (width, height,...)

(2) Adding object by using add_object() method:
    add_object(data, object_type, object_format, slide_page, location)
    data: str or pandas dataframe
    object_type: "text", "chart", "table"
    object_format: format for above type
    slide_page: int or None; if None, create a new slide at the end
    location: (x, y, w, h) or None; if None, handled by the PrsLayoutDesigner

(3) call execute() to create the presentation
(4) DataProcessor will transform the data based on corresponding object type, and packaged all data into data_container
(5) PrsLayoutDesigner provides layout_design_container
(6) layout_design_container and data_container are passed to PrsLayoutManager
(7) PrsLayoutManager will call data2chart, data2table, data2text to create the object, using data from data_container,
     and put the object at the location
(8) Optional: Auditor will check each object and issue warning if necessary (e.g. insufficient sample size)
(9) call save() to save the presentation file
"""

import pandas as pd
import data2table

from pptx import Presentation
from pptx.util import Inches
from utils import pptx_layout, pptx_auditor
from preprocessing import data_preprocessor

from collections import namedtuple


class PptxConstructor:

    def __init__(self, config):
        self.config = config
        self.presentation = Presentation()

        self.presentation.slide_width = config['prs_width']
        self.presentation.slide_height = config['prs_height']

        self.prs_object_pool = {}
        self.page_stack = {0: []}

    def add_object(self, data, object_type: str,
                   object_format=None, slide_page=None, position=None):

        # assure the object type is text, chart, or table
        assert (object_type == 'text') or (object_type == 'chart') or (object_type == 'table')

        if position is not None:
            # three types of position representation:
            # 1. absolute position ("a", x, y, w, h)
            # 2. relative position to boundary ("rb", x%, y%, w%, h%)
            # 3. relative position to object ("ro", uid, x, y, w, h)
            # assure the location is in the format of three
            assert isinstance(position, tuple)
            assert (position[0] == 'a') or (position[0] == 'rb') or (position[0] == 'ro')

        if slide_page is None:
            # if no slide_page is given, create a new slide directly
            # TODO: slide_page can be str or int, but if string how to determine choose the page number
            slide_page = max(self.page_stack) + 1
            self.page_stack[slide_page] = []

        # TODO: In the future, uid is produced by a hashmap based on the object argument.
        # If an object already exists, then the ObjectContainer can be reused to save space.
        # Currently do not apply this since no mechanism to handle duplicate objects with different location.
        uid = f"{object_type}_{len(self.prs_object_pool)}"

        ObjectContainer = namedtuple("obj", ['uid', 'data', "obj_type", "obj_format", "slide_page", 'position'])
        self.prs_object_pool[uid] = ObjectContainer(uid, data, object_type, object_format, slide_page, position)
        self.page_stack[slide_page].append(uid)
        return uid

    def pptx_execute(self):
        layout_designer = pptx_layout.PrsLayoutDesigner(self.config, self.prs_object_pool, self.page_stack)
        layout_designer.execute()
        design_structure = layout_designer.layout_design_export()

        data_processor = data_preprocessor.DataProcessor(self.prs_object_pool)
        data_container = data_processor.data_container_export()

        layout_manager = pptx_layout.PrsLayoutManager(presentation=self.presentation,
                                                      object_stack=self.prs_object_pool,
                                                      layout_design_container=design_structure,
                                                      data_container=data_container)
        result = layout_manager.layout_execute()
        return result

    def pptx_save(self, filepath='C://Users/user/Desktop/untitled.pptx'):
        self.presentation.save(filepath)
        print(f"INFO: presentation file is saved successfully as {filepath}")

    def _set_presentation_size(self, width, height):
        self.presentation.slide_width = width
        self.presentation.slide_height = height

    def _presentation_reset(self):
        self.presentation = Presentation()
        self._set_presentation_size(Inches(13.33), Inches(7.5))
