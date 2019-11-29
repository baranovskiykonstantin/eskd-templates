import re
import sys

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
config = sys.modules["config" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
textwidth = sys.modules["textwidth" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

def syncCommonFields():
    """Обновить значения общих граф.

    Обновить значения граф форматной рамки последующих листов, которые
    совпадают с графами форматной рамки и основной надписи первого листа.
    Необходимость в обновлении возникает при изменении значения графы
    на первом листе.
    На втором и последующих листах эти графы защищены от записи.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    doc.UndoManager.lock()
    doc.lockControllers()
    for name in common.STAMP_COMMON_FIELDS:
        firstFrame = doc.TextFrames["Перв.1: " + name]
        for prefix in ("Прочие: ", "РегИзм: "):
            otherFrame = doc.TextFrames[prefix + name]
            otherFrame.String = firstFrame.String

            firstCursor = firstFrame.createTextCursor()
            otherCursor = otherFrame.createTextCursor()
            otherCursor.gotoEnd(True)
            otherCursor.CharHeight = firstCursor.CharHeight
            otherCursor.CharScaleWidth = firstCursor.CharScaleWidth
    doc.unlockControllers()
    doc.UndoManager.unlock()

def setFirstPageFrameValue(name, value):
    """Установить значение поля форматной рамки первого листа.

    Имеется четыре стиля для первого листа и у каждого содержится
    свой набор полей. Они должны иметь одинаковые значения, чтобы
    при смене стиля значения сохранялись.

    Атрибуты:
    name -- имя текстового поля (без первых 4-х символов);
    value -- новое значение поля.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    firstPageStyleName = doc.Text.createTextCursor().PageDescName
    doc.UndoManager.lock()
    doc.lockControllers()
    for firstPageVariant in "1234":
        fullName = "Перв.{}: {}".format(firstPageVariant, name)
        if fullName in doc.TextFrames:
            frame = doc.TextFrames[fullName]
            # Записать в буфер действий для отмены
            # только изменения текущего стиля
            if firstPageStyleName.endswith(firstPageVariant):
                doc.UndoManager.unlock()
            frame.String = value
            if firstPageStyleName.endswith(firstPageVariant):
                doc.UndoManager.lock()
            if name in common.ITEM_WIDTHS:
                cursor = frame.createTextCursor()
                for line in value.splitlines(keepends=True):
                    widthFactor = textwidth.getWidthFactor(
                        line,
                        cursor.CharHeight,
                        common.ITEM_WIDTHS[name] - 1
                    )
                    cursor.goRight(len(line), True)
                    cursor.CharScaleWidth = widthFactor
                    cursor.collapseToEnd()
    doc.unlockControllers()
    doc.UndoManager.unlock()

def clean(*args):
    """Очистить основную надпись.

    Удалить содержимое полей основной надписи и форматной рамки.

    """
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    firstPageStyleName = doc.Text.createTextCursor().PageDescName
    doc.UndoManager.lock()
    doc.lockControllers()
    for frame in doc.TextFrames:
        if frame.Name.startswith("Перв.") \
            and not frame.Name.endswith("(наименование)") \
            and not frame.Name.endswith("7 Лист") \
            and not frame.Name.endswith("8 Листов"):
                # Записать в буфер действий для отмены
                # только изменения текущего стиля
                if firstPageStyleName[-1] == frame.Name[5]:
                    doc.UndoManager.unlock()
                frame.String = ""
                if firstPageStyleName[-1] == frame.Name[5]:
                    doc.UndoManager.lock()
                cursor = frame.createTextCursor()
                cursor.CharScaleWidth = 100
    if config.getboolean("stamp", "place doc id to table title") \
        and "Спецификация" in doc.TextTables:
            amountTitleCell = doc.TextTables["Спецификация"].getCellByName("F1")
            amountTitleCell.String = "Кол. на исполнение"
    syncCommonFields()
    doc.unlockControllers()
    doc.UndoManager.unlock()

def fill(*args):
    """Заполнить основную надпись.

    Считать данные из файла списка цепей и заполнить графы основной надписи.
    Имеющиеся данные не удаляются из основной надписи.
    Новое значение вносится только в том случае, если оно не пустое.

    """
    if common.isThreadWorking():
        return
    schematic = common.getSchematicData()
    if schematic is None:
        return
    doc = XSCRIPTCONTEXT.getDocument()
    doc.lockControllers()
    # Наименование документа
    docTitle = schematic.title.replace('\\n', '\n')
    if config.getboolean("stamp", "convert doc title"):
        tailPos = docTitle.find("Схема электрическая")
        if tailPos > 0:
            docTitle = docTitle[:tailPos]
        docTitle = docTitle.strip()
    setFirstPageFrameValue("1 Наименование документа", docTitle)
    # Наименование организации
    setFirstPageFrameValue("9 Наименование организации", schematic.company)
    # Обозначение документа
    docId = schematic.number
    idParts = re.match(
        r"([А-ЯA-Z0-9]+(?:[\.\-]\d+)+\s?)(Э\d)",
        docId
    )
    if config.getboolean("stamp", "convert doc id") \
        and idParts is not None:
            docId = idParts.group(1).strip()
    setFirstPageFrameValue("2 Обозначение документа", docId)
    if config.getboolean("stamp", "place doc id to table title") \
        and docId and "Спецификация" in doc.TextTables:
            amountTitleCell = doc.TextTables["Спецификация"].getCellByName("F1")
            amountTitleCell.String = "Кол. на исполнение {}-".format(docId)
    # Первое применение
    if config.getboolean("stamp", "fill first usage") \
        and idParts is not None:
            setFirstPageFrameValue("25 Перв. примен.", idParts.group(1).strip())
    # Разработал
    setFirstPageFrameValue("11 Разраб.", schematic.developer)
    # Проверил
    setFirstPageFrameValue("11 Пров.", schematic.verifier)
    # Нормативный контроль
    setFirstPageFrameValue("11 Н. контр.", schematic.inspector)
    # Утвердил
    setFirstPageFrameValue("11 Утв.", schematic.approver)

    syncCommonFields()
    doc.unlockControllers()

g_exportedScripts = clean, fill
