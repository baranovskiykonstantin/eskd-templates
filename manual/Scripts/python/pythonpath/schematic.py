"""Объектное представление схемы."""

import re
import sys

def init(scriptcontext):
    pass


class Schematic():
    """Данные о схеме."""

    def __init__(self):
        self.title = ""
        self.number = ""
        self.company = ""
        self.developer = ""
        self.verifier = ""
        self.inspector = ""
        self.approver = ""
