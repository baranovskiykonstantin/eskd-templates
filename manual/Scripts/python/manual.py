import sys

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

def toggleRevTable(*args):
    """Добавить/удалить таблицу регистрации изменений"""
    doc = XSCRIPTCONTEXT.getDocument()
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()
