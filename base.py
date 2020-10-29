from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.lang import MSO_LANGUAGE_ID


class DataCreator:

    def __init__(self):
        self.uid_pool = []
        self.alignment = {"left": PP_PARAGRAPH_ALIGNMENT.LEFT,
                          "right": PP_PARAGRAPH_ALIGNMENT.RIGHT,
                          "center": PP_PARAGRAPH_ALIGNMENT.CENTER}
        self.text_lang = {"tc": MSO_LANGUAGE_ID.TRADITIONAL_CHINESE}