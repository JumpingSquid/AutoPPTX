from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.util import Inches, Pt

from autoppt_util import data2chart
from autoppt_util.data2table import TableCreator
from autoppt_util.pptx_params import textbox
from data_preprocessing import chart_data_preprocessor


class PrsLayoutManager:

    def __init__(self, presentation, config):
        self.presentation = presentation
        self.prs_height = presentation.slide_height
        self.prs_width = presentation.slide_width
        self.dataprocessor = None

        # read in the title config
        # if no title in the config, use the blank slide instead
        if config["title"] is not None:
            self.title_config = config["title"]
            bullet_slide_layout = self.presentation.slide_layouts[5]
        else:
            bullet_slide_layout = self.presentation.slide_layouts[6]

        self.slide = self.presentation.slides.add_slide(bullet_slide_layout)

        # read in the chart config
        self.chart_config = config["chart"]

        # without specified chart id, the chart is created based on the order of the chart list
        if not config["setting"]["chart_id_sepcifed"]:
            self._chart_id = 0
            self.chart_rescale = self.chart_config[self._chart_id]["rescale"]
            self.chart_num_on_slide = len(self.chart_config)

        # beta: allow for customized layout for charts, but require the user to provide fully defined layout
        # the user should provide a list contains (1) origin and (2) the box size for each chart
        if "custom_layout" not in config:
            self.custom_layout_flag = False
            self.chart_origin_anchor = [Inches(0), Inches(1.65)]
            print("INFO: ", self.chart_num_on_slide, "chart(s) required")
            self.chart_box_size = [self.prs_width / max(3, self.chart_num_on_slide) * self.chart_rescale[0],
                                   self.prs_height * 0.7 * self.chart_rescale[1]]
        elif "custom_layout" in config:
            self.custom_layout_flag = True
            self.custom_layout_config = config["custom_layout"]
            self.chart_origin_anchor = config["custom_layout"]["origin"][self._chart_id]
            print("INFO: ", "Self-defined layout is used")
            self.chart_box_size = config["custom_layout"]["size"][self._chart_id]

        # store the object that exists on the slide
        self.object_pool = []

    def read_dataprocessor(self, dataprocessor: chart_data_preprocessor):
        self.dataprocessor = dataprocessor

    def slide_chart_layout(self, slide):

        # setting the limit for showing sample size warning
        sample_size_limit = 20

        # load the format of the chart
        chart_created_config = self.chart_config[self._chart_id]
        if "format" not in chart_created_config:
            chart_created_config["format"] = None

        # basic chart type: hist, stackbar, stackcolumn
        if chart_created_config["type"] in ["hist", "stackbar", "stackcolumn"]:
            print("INFO: chart type", chart_created_config["type"], "require only 1 chart, full space will be used")
            chartcreator = data2chart.ChatCreator(chart_category=self.dataprocessor.single_category_push()["category"],
                                                  chart_series=self.dataprocessor.single_category_push()["series"],
                                                  origin=self.chart_origin_anchor,
                                                  size=self.chart_box_size,
                                                  chart_type=chart_created_config["type"],
                                                  chart_format=chart_created_config["format"])
            chartcreator.chart_create_on_slide(slide)
            self.sample_size_legend(slide, self.dataprocessor.samplecount, left=self.chart_origin_anchor[0])
            if int(self.dataprocessor.totalsample) < sample_size_limit:
                self.sample_size_warning(self.chart_origin_anchor[0],
                                         self.chart_origin_anchor[1] + self.chart_box_size[1] * 0.5)
            self.object_pool.append((self.chart_origin_anchor, self.chart_box_size, chart_created_config["type"]))
            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        # basic chart type: pie
        elif chart_created_config["type"] in ["pie", "p"]:
            print("INFO: chart type", chart_created_config["type"], "require only 1 chart, full space will be used")
            chartcreator = data2chart.ChatCreator(chart_category=self.dataprocessor.pie_push()["category"],
                                                  chart_series=self.dataprocessor.pie_push()["series"],
                                                  origin=self.chart_origin_anchor,
                                                  size=self.chart_box_size,
                                                  chart_type="pie",
                                                  chart_format=chart_created_config["format"])
            chartcreator.chart_create_on_slide(slide)
            self.sample_size_legend(slide, self.dataprocessor.totalsample, left=self.chart_origin_anchor[0])

            if int(self.dataprocessor.totalsample) < sample_size_limit:
                self.sample_size_warning(self.chart_origin_anchor[0],
                                         self.chart_origin_anchor[1] + self.chart_box_size[1] * 0.5)
            self.object_pool.append((self.chart_origin_anchor, self.chart_box_size, chart_created_config["type"]))

            # new function: add appendix under the chart created
            if chart_created_config["appendix"] is not None:
                appendix_bias = self.chart_origin_anchor[1] + self.chart_box_size[1]
                for appendix_ele in chart_created_config["appendix"]:
                    appendix_bias = self.chart_appendix(slide,
                                                        appendix_ele,
                                                        left=self.chart_origin_anchor[0],
                                                        top=appendix_bias,
                                                        width=self.chart_box_size[0])

            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["sep_bar", "sep_b"]:
            print("INFO: chart type", chart_created_config["type"], "require only 1 chart, full space will be used")
            chartcreator = data2chart.ChatCreator(chart_category=self.dataprocessor.sep_bar_push()["category"],
                                                  chart_series=self.dataprocessor.sep_bar_push()["series"],
                                                  origin=self.chart_origin_anchor,
                                                  size=self.chart_box_size,
                                                  chart_type="bar",
                                                  chart_format=chart_created_config["format"])
            chartcreator.chart_create_on_slide(slide)
            self.sample_size_legend(slide, self.dataprocessor.totalsample, left=self.chart_origin_anchor[0])

            if int(self.dataprocessor.totalsample) < sample_size_limit:
                self.sample_size_warning(self.chart_origin_anchor[0],
                                         self.chart_origin_anchor[1] + self.chart_box_size[1] * 0.5)
            self.object_pool.append((self.chart_origin_anchor, self.chart_box_size, chart_created_config["type"]))

            # new function: add appendix under the chart created
            if chart_created_config["appendix"] is not None:
                appendix_bias = self.chart_origin_anchor[1] + self.chart_box_size[1]
                for appendix_ele in chart_created_config["appendix"]:
                    appendix_bias = self.chart_appendix(slide,
                                                        appendix_ele,
                                                        left=self.chart_origin_anchor[0],
                                                        top=appendix_bias,
                                                        width=self.chart_box_size[0])

            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] == ["line", "l"]:
            print("INFO: chart type", chart_created_config["type"], "require only 1 chart, full space will be used")
            chartcreator = data2chart.ChatCreator(chart_category=self.dataprocessor.line_push()["category"],
                                                  chart_series=self.dataprocessor.line_push()["series"],
                                                  origin=self.chart_origin_anchor,
                                                  size=self.chart_box_size,
                                                  chart_type="line",
                                                  chart_format=chart_created_config["format"])
            chartcreator.chart_create_on_slide(slide)
            self.sample_size_legend(slide, self.dataprocessor.samplecount, left=self.chart_origin_anchor[0])
            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["single_category_multiple_bar", "scmb"]:
            chart_num = self.dataprocessor.category_num
            print("INFO: chart type", chart_created_config["type"], "requires", chart_num, "charts")
            self.sample_size_legend(slide, self.dataprocessor.samplecount, left=self.chart_origin_anchor[0])
            chart_size_split_num = chart_num // 3
            if chart_num % 3 != 0:
                chart_size_split_num += 1
            self.chart_box_size = [self.chart_box_size[0] / chart_size_split_num, self.chart_box_size[1] / 3.2]
            for chart_id in range(0, chart_num):
                chartcreator = data2chart.ChatCreator(chart_category=self.dataprocessor.multi_category_push(chart_id)["category"],
                                                      chart_series=self.dataprocessor.multi_category_push(chart_id)["series"],
                                                      origin=self.chart_origin_anchor,
                                                      size=self.chart_box_size,
                                                      chart_type="bar",
                                                      chart_format=chart_created_config["format"])
                slide = chartcreator.chart_create_on_slide(slide)
                if chart_id % 3 != 2:
                    self.chart_origin_anchor = [self.chart_origin_anchor[0],
                                                self.chart_origin_anchor[1] + self.chart_box_size[1]]
                else:
                    self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0],
                                                Inches(1.65)]
            if chart_num % 3 != 0:
                self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]
            else:
                self.chart_origin_anchor = [self.chart_origin_anchor[0], Inches(1.65)]

        elif chart_created_config["type"] in ["multiple_value_column_bar", "mv"]:
            chartcreator = data2chart.ChatCreator(chart_category=self.dataprocessor.multi_value_column_push()["category"],
                                                  chart_series=self.dataprocessor.multi_value_column_push()["series"],
                                                  origin=self.chart_origin_anchor,
                                                  size=self.chart_box_size,
                                                  chart_type="bar",
                                                  chart_format=chart_created_config["format"])
            slide = chartcreator.chart_create_on_slide(slide)
            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["multiple_category_multiple_value_column_bar", "mcmv"]:
            chart_num = self.dataprocessor.category_num
            print("INFO: chart type", chart_created_config["type"], "requires", chart_num, "charts")

            chart_size_split_num = chart_num

            self.chart_box_size = [self.chart_box_size[0], self.chart_box_size[1] / chart_size_split_num]
            for chart_id in range(0, chart_num):
                chartcreator = data2chart.ChatCreator(
                    chart_category=self.dataprocessor.multi_category_multi_value_column_push(chart_id)["category"],
                    chart_series=self.dataprocessor.multi_category_multi_value_column_push(chart_id)["series"],
                    origin=self.chart_origin_anchor,
                    size=self.chart_box_size,
                    chart_type="bar",
                    chart_format=chart_created_config["format"])
                slide = chartcreator.chart_create_on_slide(slide)
                self.chart_origin_anchor = [self.chart_origin_anchor[0],
                                            self.chart_origin_anchor[1] + self.chart_box_size[1]]

            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["multiple_dummy_one_column_bar", "mdc"]:
            chart_num = len(self.chart_config[self._chart_id]["cat_columns"])
            print("INFO: chart type", chart_created_config["type"], "requires", chart_num, "charts")

            self.chart_box_size = [size_ele / chart_num for size_ele in self.chart_box_size]

            for chart_id in range(0, chart_num):
                chartcreator = data2chart.ChatCreator(
                    chart_category=self.dataprocessor.multiple_dummy_one_column_push(chart_id)["category"],
                    chart_series=self.dataprocessor.multiple_dummy_one_column_push(chart_id)["series"],
                    origin=self.chart_origin_anchor,
                    size=self.chart_box_size,
                    chart_type="bar",
                    chart_format=chart_created_config["format"])
                slide = chartcreator.chart_create_on_slide(slide)
                self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        else:
            raise ValueError("ERROR: Unknow type of chart,", chart_created_config["type"], "is introduced")

        return slide

    def add_chart_on_slide(self):
        # try to create chart, if there is ZeroDivisionError for creating chart data, skip the chart first
        # otherwise stop the task
        try:
            self.slide_chart_layout(self.slide)
            self.chart_rescale = self.chart_config[self._chart_id]["rescale"]
            print(self._chart_id, "chart is added")
        except ZeroDivisionError:
            print("WARN: the chart-", self._chart_id, "incurs ZeroDivisionError, skipped automatically")
            pass
        if self._chart_id < self.chart_num_on_slide - 1:
            self._chart_id += 1
            # reset the chart size and anchor after creating the chart if no pre-defined layout
            # o/w navigating to the next object's spec
            if not self.custom_layout_flag:
                self.chart_box_size = [self.prs_width / max(3, self.chart_num_on_slide) * self.chart_rescale[0],
                                       self.prs_height * 0.7 * self.chart_rescale[1]]
            else:
                self.chart_origin_anchor = self.custom_layout_config["origin"][self._chart_id]
                self.chart_box_size = self.custom_layout_config["size"][self._chart_id]
        return self.presentation

    def add_title_on_slide(self):
        # create title
        shapes = self.slide.shapes
        title_shape = shapes.title
        text_frame = title_shape.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.text = self.title_config
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        p.font.name = textbox("title_font")
        p.font.size = Pt(40)
        p.font.language_id = MSO_LANGUAGE_ID.TRADITIONAL_CHINESE
        return self.presentation

    def add_comment_on_slide(self, commentcreator):
        # the position of the comment
        left = top = Inches(0)
        width = Inches(13.33)
        height = Inches(0.4)
        txBox = self.slide.shapes.add_textbox(left, top, width, height)

        # the color of the comment
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(200, 200, 200)

        # the content and the format of the comment
        tf = txBox.text_frame
        tf.clear()  # not necessary for newly-created shape
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = commentcreator.comment_create()
        font = run.font
        font.name = textbox("comment_font")
        font.size = Pt(16)
        font.language_id = MSO_LANGUAGE_ID.TRADITIONAL_CHINESE
        return self.presentation

    def add_table_on_slide(self, tablecreator: TableCreator):
        # the position of the comment
        left = tablecreator.origin[0]
        top = tablecreator.origin[1]
        width = tablecreator.width
        height = tablecreator.height
        row_num = tablecreator.row_num
        col_num = tablecreator.col_num
        table = self.slide.shapes.add_table(row_num, col_num, left, top, width, height).table
        tablecreator.table_data_fill(table)
        return self.presentation

    # create a warning tag on the chart where the sample size is below 15
    def sample_size_warning(self, left=Inches(8.33), top=Inches(1)):
        # the position of the comment
        width = Inches(2.66)
        height = Inches(0.37)
        txBox = self.slide.shapes.add_textbox(left, top, width, height)

        # the color of the comment
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(200, 200, 200)

        # the content and the format of the comment
        tf = txBox.text_frame
        tf.clear()  # not necessary for newly-created shape
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = textbox("sample_warn_text")
        font = run.font
        font.name = textbox("sample_warn_font")
        font.size = Pt(16)
        return self.presentation

    # draw the layout setting on the screen, assist checking the number and the location of object created
    def layout_describe(self):

        def square_painter(draw_lst, x, y, lx, ly):
            draw_lst += [(x, y+y_bias) for y_bias in range(ly)]
            draw_lst += [(x+lx, y+y_bias) for y_bias in range(ly)]
            draw_lst += [(x+x_bias, y) for x_bias in range(lx)]
            draw_lst += [(x+x_bias, y+ly) for x_bias in range(lx)]
            return draw_lst

        draw_point = []
        draw_point = square_painter(draw_point, 0, 0, 55, 17)
        for object in self.object_pool:
            draw_point = square_painter(draw_point,
                                        int(object[0][0]/Inches(13.33)*55),
                                        int(object[0][1]/Inches(7.5)*17),
                                        int(object[1][0]/Inches(13.33)*55),
                                        int(object[1][1]/Inches(7.5)*17))

        for y_painter in range(0, 18):
            drawer = ""
            for x_painter in range(0, 56):
                if (x_painter, y_painter) in draw_point:
                    drawer += "#"
                else:
                    drawer += " "
            print(drawer)

    # add the sample size legend for the chart created
    @staticmethod
    def sample_size_legend(slide, samplecount, left=Inches(0), top=Inches(7.2)):
        width = Inches(2)
        height = Inches(0.27)
        txBox = slide.shapes.add_textbox(left, top, width, height)

        # the color of the comment
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)

        # the content and the format of the comment
        tf = txBox.text_frame
        tf.clear()  # not necessary for newly-created shape
        p = tf.paragraphs[0]
        run = p.add_run()
        if isinstance(samplecount, str):
            run.text = samplecount
        else:
            run.text = textbox("sample_size_text") + str(samplecount)
        font = run.font
        font.name = textbox("sample_size_font")
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        font.size = Pt(10)

    # chart appendix: a function that append additional objects below the indicated chart
    def chart_appendix(self, slide, appendix, left=Inches(0), top=Inches(7.2), width=Inches(2), height=Inches(0.27)):
        # used to create additional ppt object under the chart
        # right now it is in hard written style
        # input: tuple(appendix_type, appendix_content)
        if appendix[0] == "text":
            chart_appendix_tbox = slide.shapes.add_textbox(left, top, width, height)
            word_per_line = int(width/Inches(0.25))+1

            # the color of the comment
            fill = chart_appendix_tbox.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 255, 255)

            # the content and the format of the comment
            tf = chart_appendix_tbox.text_frame
            tf.clear()  # not necessary for newly-created shape
            p = tf.paragraphs[0]
            run = p.add_run()

            raw_text = appendix[1]
            raw_text = raw_text.split("\n")
            raw_text = [x[:word_per_line]+"\n"+x[word_per_line:] if len(x) > word_per_line else x for x in raw_text]
            run.text = "\n".join(raw_text)

            font = run.font
            font.name = textbox("chart_appendix_font")
            font.language_id = MSO_LANGUAGE_ID.TRADITIONAL_CHINESE
            p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
            font.size = Pt(12)

            appendix_bias = top + (appendix[1].count("\n") + 1)*Inches(0.25)

        elif appendix[0] == "chart":
            from pptx.chart.data import CategoryChartData
            from pptx.enum.chart import XL_CHART_TYPE
            chart_data = CategoryChartData()
            chart_data.categories = [appendix[2]]
            chart_data.add_series("series_1", tuple([appendix[1]]))
            height = Inches(1)
            chart_frame = slide.shapes.add_chart(
                XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data
            )
            chart = chart_frame.chart
            chart.chart_style = 29
            value_axis = chart_frame.chart.value_axis
            value_axis.minimum_scale = 0.0
            value_axis.maximum_scale = 100.0
            chart_frame.chart.font.size = Pt(12)
            chart_frame.chart.plots[0].has_data_labels = True

            appendix_bias = top + height

        self.object_pool.append(([left, top], [width, height], "chart_appendix"))

        return appendix_bias
