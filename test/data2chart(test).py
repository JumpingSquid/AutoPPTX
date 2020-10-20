import unittest

from data2chart import ChartCreator, ChartFormatSetter
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
from pptx.util import Pt
from pptx import Presentation
from pptx.util import Inches


class TestChart(unittest.TestCase):
    def setUp(self) -> None:
        self.chart_creator = ChartCreator()


class Testinit(TestChart):
    def test_prs_initializaion(self):
        self.assertIsNotNone(self.chart_creator.prs)

    def test_format_initialization(self):
        self.assertIsNotNone(self.chart_creator.chart_format)

    def test_one_slide_pool(self):
        self.assertEqual(len(self.chart_creator.slide_pool), 1)