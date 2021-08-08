import sys

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

def toggleRevTable(*args):
    """Добавить/удалить таблицу регистрации изменений"""
    doc = XSCRIPTCONTEXT.getDocument()
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()

def togglePageRevTable(*args):
    """Добавить/удалить таблицу изменений на текущей странице"""
    doc = XSCRIPTCONTEXT.getDocument()
    frameName = "Изм_стр_%d" % doc.CurrentController.ViewCursor.Page
    if frameName in doc.TextFrames:
        common.removePageRevTable()
    else:
        common.addPageRevTable()
