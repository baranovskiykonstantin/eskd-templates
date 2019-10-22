import re
import sys

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
config = sys.modules["config" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
textwidth = sys.modules["textwidth" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

def syncCommonFields():
    """Обновить значения общих граф.

    Обновить значения граф форматной рамки последующих листов, которые
    совпадают с графами форматной рамки и основной надписи первого листа.
    К таким графам относятся:
    - 2. Обозначение документа;
    - 19. Инв. № подл.;
    - 21. Взам. инв. №;
    - 22. Инв. № дубл.
    Необходимость в обновлении возникает при изменении значения графы
    на первом листе.
    На втором и последующих листах эти графы защищены от записи.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    doc.getUndoManager().lock()
    doc.lockControllers()
    for name in common.STAMP_COMMON_FIELDS:
        firstFrame = doc.getTextFrames().getByName("1.1." + name)
        otherFrame = doc.getTextFrames().getByName("N." + name)
        otherFrame.setString(firstFrame.getString())

        firstCursor = firstFrame.createTextCursor()
        otherCursor = otherFrame.createTextCursor()
        otherCursor.gotoEnd(True)
        otherCursor.CharHeight = firstCursor.CharHeight
        otherCursor.CharScaleWidth = firstCursor.CharScaleWidth
    doc.unlockControllers()
    doc.getUndoManager().unlock()

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
    firstPageStyleName = doc.getText().createTextCursor().PageDescName
    doc.getUndoManager().lock()
    doc.lockControllers()
    for i in range(1, 5):
        fullName = "1.{}.{}".format(i, name)
        if doc.getTextFrames().hasByName(fullName):
            frame = doc.getTextFrames().getByName(fullName)
            # Записать в буфер действий для отмены
            # только изменения текущего стиля
            if firstPageStyleName.endswith(str(i)):
                doc.getUndoManager().unlock()
            frame.setString(value)
            if firstPageStyleName.endswith(str(i)):
                doc.getUndoManager().lock()
            if name in common.ITEM_WIDTHS:
                cursor = frame.createTextCursor()
                cursor.gotoEnd(True)
                longestLine = max(value.splitlines(), key=len) if value else ""
                cursor.CharScaleWidth = textwidth.getWidthFactor(
                    longestLine,
                    cursor.CharHeight * (cursor.CharEscapementHeight / 100),
                    common.ITEM_WIDTHS[name] - 1
                )
    doc.unlockControllers()
    doc.getUndoManager().unlock()

def clean(*args):
    """Очистить основную надпись.

    Удалить содержимое полей основной надписи и форматной рамки.

    """
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    firstPageStyleName = doc.getText().createTextCursor().PageDescName
    doc.getUndoManager().lock()
    doc.lockControllers()
    for frame in doc.getTextFrames():
        if frame.Name.startswith("1.") \
            and not frame.Name.endswith("(наименование)") \
            and not frame.Name.endswith(".7 Лист") \
            and not frame.Name.endswith(".8 Листов"):
                # Записать в буфер действий для отмены
                # только изменения текущего стиля
                if firstPageStyleName[-1] == frame.Name[2]:
                    doc.getUndoManager().unlock()
                frame.setString("")
                if firstPageStyleName[-1] == frame.Name[2]:
                    doc.getUndoManager().lock()
                cursor = frame.createTextCursor()
                cursor.CharScaleWidth = 100
    syncCommonFields()
    doc.unlockControllers()
    doc.getUndoManager().unlock()

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
