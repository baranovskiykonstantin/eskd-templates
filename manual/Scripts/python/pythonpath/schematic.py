"""Объектное представление схемы."""

import sys

kicadnet = None

def init(scriptcontext):
    global kicadnet
    kicadnet = sys.modules["kicadnet" + scriptcontext.getDocument().RuntimeUID]


class Schematic():
    """Данные о схеме."""

    def __init__(self, netlistName):
        self.title = ""
        self.number = ""
        self.company = ""
        self.developer = ""
        self.verifier = ""
        self.inspector = ""
        self.approver = ""

        netlist = kicadnet.Netlist(netlistName)
        for sheet in netlist.items("sheet"):
            if sheet.attributes["name"] == "/":
                title_block = netlist.find("title_block", sheet)
                for item in title_block.items:
                    if item.name == "title":
                        self.title = item.text if item.text is not None else ""
                    elif item.name == "company":
                        self.company = item.text if item.text is not None else ""
                    elif item.name == "comment":
                        if item.attributes["number"] == "1":
                            self.number = item.attributes["value"]
                        elif item.attributes["number"] == "2":
                            self.developer = item.attributes["value"]
                        elif item.attributes["number"] == "3":
                            self.verifier = item.attributes["value"]
                        elif item.attributes["number"] == "4":
                            self.approver = item.attributes["value"]
                        elif item.attributes["number"] == "6":
                            self.inspector = item.attributes["value"]
                break
