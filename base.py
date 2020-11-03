from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx import Presentation
from pptx.util import Inches


class ObjectWorker:

    def __init__(self, prs):
        if prs:
            # initializing a basic presentation file when no existing file is given
            self.prs = self.create_prs()

    @staticmethod
    def create_prs() -> Presentation():
        prs = Presentation()
        # slide size: 16:9
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        return prs
