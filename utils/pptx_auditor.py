"""
PrsAuditor is used to check the potential invalidity in the presentation (e.g. insufficient sample size),
it will scan the data_container, layout design structure, and other meta data to issue warnings.
"""


class PrsAuditor:
    def __init__(self):
        self.data = None
