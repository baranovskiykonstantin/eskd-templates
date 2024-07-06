"""Вспомогательные функции и общие данные.

Модуль содержит вспомогательные функции и данные, которые могут использоваться
различными макросами.

"""

import os
import sys
import traceback
import threading
import uno

XSCRIPTCONTEXT = None
schematic = None
kicadnet = None
config = None
textwidth = None

def init(scriptcontext):
    global XSCRIPTCONTEXT
    global schematic
    global kicadnet
    global config
    global textwidth
    XSCRIPTCONTEXT = scriptcontext
    schematic = sys.modules["schematic" + scriptcontext.getDocument().RuntimeUID]
    kicadnet = sys.modules["kicadnet" + scriptcontext.getDocument().RuntimeUID]
    config = sys.modules["config" + scriptcontext.getDocument().RuntimeUID]
    textwidth = sys.modules["textwidth" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

STAMP_COMMON_FIELDS = (
    "2 Обозначение документа",
    "19 Инв. № подл.",
    "20 Подп. и дата",
    "21 Взам. инв. №",
    "22 Инв. № дубл.",
    "23 Подп. и дата"
)

ITEM_WIDTHS = {
    "1 Наименование документа": 70,
    "2 Обозначение документа": 120,
    "4 Лит.1": 5,
    "4 Лит.2": 5,
    "4 Лит.3": 5,
    "9 Наименование организации": 50,
    "10": 17,
    "11": 23,
    "11 Н. контр.": 23,
    "11 Пров.": 23,
    "11 Разраб.": 23,
    "11 Утв.": 23,
    "13": 10,
    "13 Н. контр.": 10,
    "13 Пров.": 10,
    "13 Разраб.": 10,
    "13 Утв.": 10,
    "19 Инв. № подл.": 25,
    "20 Подп. и дата": 35,
    "21 Взам. инв. №": 25,
    "22 Инв. № дубл.": 25,
    "23 Подп. и дата": 35,
    "24 Справ. №": 60,
    "25 Перв. примен.": 60,
    "27": 14,
    "28": 53,
    "29": 53,
    "30": 120,

    "ТабРИ.A": 8,
    "ТабРИ.B": 20,
    "ТабРИ.C": 20,
    "ТабРИ.D": 20,
    "ТабРИ.E": 20,
    "ТабРИ.F": 20,
    "ТабРИ.G": 25,
    "ТабРИ.H": 25,
    "ТабРИ.I": 15,
    "ТабРИ.J": 12,

    "ТабТИ.A": 7,
    "ТабТИ.B": 10,
    "ТабТИ.C": 23,
    "ТабТИ.D": 15,
    "ТабТИ.E": 10,
}

SKIP_MODIFY_EVENTS = False

def showMessage(text, title="Сообщение"):
    """Показать текстовое сообщение.

    Аргументы:
    text -- текст сообщения;
    title -- заголовок окна сообщения.

    """
    window = XSCRIPTCONTEXT.getDocument().CurrentController.Frame.ContainerWindow
    msgbox = window.Toolkit.createMessageBox(
        window,
        uno.Enum("com.sun.star.awt.MessageBoxType", "MESSAGEBOX"),
        uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK"),
        title,
        text
    )
    msgbox.execute()

def showFilePicker(filePath="", title="Выбор файла с данными о схеме", **fileFilters):
    """Показать диалоговое окно выбора файла.

    Аргументы:

    filePath -- имя файла по умолчанию;
    fileFilters -- перечень фильтров для выбора файлов в формате:
        {"Текстовые файлы": "*.txt", "Документы": "*.odt;*.ods"}

    Возвращаемое значение -- полное имя файла или None, если файл не выбран.

    """
    context = XSCRIPTCONTEXT.getComponentContext()
    if os.path.isfile(filePath):
        directory, file = os.path.split(filePath)
    else:
        docUrl = XSCRIPTCONTEXT.getDocument().URL
        if docUrl:
            directory = os.path.dirname(uno.fileUrlToSystemPath(docUrl))
        else:
            directory = os.path.expanduser('~')
        file = ""
    filePicker = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.ui.dialogs.OfficeFilePicker",
        context
    )
    filePicker.setTitle(title)
    pickerType = uno.getConstantByName(
        "com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE"
    )
    filePicker.initialize((pickerType,))
    filePicker.setDisplayDirectory(
        uno.systemPathToFileUrl(directory)
    )
    filePicker.setDefaultName(file)
    for filterTitle, filterValue in fileFilters.items():
        filePicker.appendFilter(filterTitle, filterValue)
        if not filePicker.getCurrentFilter():
            # Установить первый фильтр в качестве фильтра по умолчанию.
            filePicker.setCurrentFilter(filterTitle)
    result = filePicker.execute()
    OK = uno.getConstantByName(
        "com.sun.star.ui.dialogs.ExecutableDialogResults.OK"
    )
    if result == OK:
        sourcePath = uno.fileUrlToSystemPath(filePicker.Files[0])
        return sourcePath
    return None

def getSourceFileName():
    """Получить имя файла с данными о схеме.

    Попытаться найти файл с данными о схеме в текущем каталоге.
    В случае неудачи, показать диалоговое окно выбора файла.

    Для KiCad источником данных о схеме является список цепей.

    Возвращаемое значение -- полное имя файла или None, если файл
        не найден или не выбран.

    """
    sourcePath = config.get("doc", "source")
    if os.path.exists(sourcePath):
        return sourcePath
    sourceDir = ""
    sourceName = ""
    docUrl = XSCRIPTCONTEXT.getDocument().URL
    if docUrl:
        docPath = uno.fileUrlToSystemPath(docUrl)
        sourceDir = os.path.dirname(docPath)
        for fileName in os.listdir(sourceDir):
            if fileName.endswith(".pro"):
                sourceName = fileName.replace(".pro", ".net")
        if sourceName:
            sourcePath = os.path.join(sourceDir, sourceName)
            if os.path.exists(sourcePath):
                config.set("doc", "source", sourcePath)
                config.save()
                return sourcePath
    sourcePath = showFilePicker(
        os.path.join(sourceDir, sourceName),
        **{"Список цепей KiCad": "*.net;*.xml", "Все файлы": "*.*"}
    )
    if sourcePath is not None:
        config.set("doc", "source", sourcePath)
        config.save()
        return sourcePath
    return None

def getSchematicData():
    """Подготовить необходимые данные о схеме.

    Выбрать из файла данные о компонентах и данные для заполнения
    основной надписи.

    Возвращаемое значение -- объект класса Schematic или None, если
        файл не найден или данные в файле отсутствуют.

    """
    sourceFileName = getSourceFileName()
    if sourceFileName is None:
        # Отменено пользователем
        #showMessage(
        #    "Не удалось получить данные о схеме.",
        #    "Пояснительная записка"
        #)
        return None
    try:
        return schematic.Schematic(sourceFileName)
    except kicadnet.ParseException as error:
        showMessage(
            "Не удалось получить данные о схеме.\n\n" \
            "При разборе файла обнаружена ошибка:\n" \
            + str(error),
            "Пояснительная записка"
        )
    except:
        showMessage(
            "Не удалось получить данные о схеме.\n\n" \
            + traceback.format_exc(),
            "Пояснительная записка"
        )
    return None

def appendRevTable():
    """Добавить таблицу регистрации изменений."""
    doc = XSCRIPTCONTEXT.getDocument()
    if "Лист_регистрации_изменений" in doc.TextTables:
        return
    global SKIP_MODIFY_EVENTS
    SKIP_MODIFY_EVENTS = True
    doc.lockControllers()
    doc.UndoManager.lock()
    text = doc.Text
    text.insertControlCharacter(
        text.End,
        uno.getConstantByName(
            "com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"
        ),
        False
    )
    # Таблица
    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(4, 10)
    text.insertTextContent(text.End, table, False)
    table.Name = "Лист_регистрации_изменений"
    table.BreakType = uno.Enum("com.sun.star.style.BreakType", "PAGE_BEFORE")
    table.PageDescName = "Лист регистрации изменений"
    table.HoriOrient = uno.getConstantByName("com.sun.star.text.HoriOrientation.LEFT_AND_WIDTH")
    table.Width = 18500
    table.LeftMargin = 2000
    columnSeparators = table.TableColumnSeparators
    columnSeparators[0].Position = 432   # int((8)/185*10000)
    columnSeparators[1].Position = 1512  # int((8+20)/185*10000)
    columnSeparators[2].Position = 2594  # int((8+20+20)/185*10000)
    columnSeparators[3].Position = 3675  # int((8+20+20+20)/185*10000)
    columnSeparators[4].Position = 4756  # int((8+20+20+20+20)/185*10000)
    columnSeparators[5].Position = 5837  # int((8+20+20+20+20+20)/185*10000)
    columnSeparators[6].Position = 7189  # int((8+20+20+20+20+20+25)/185*10000)
    columnSeparators[7].Position = 8540  # int((8+20+20+20+20+20+25+25)/185*10000)
    columnSeparators[8].Position = 9351  # int((8+20+20+20+20+20+25+25+15)/185*10000)
    table.TableColumnSeparators = columnSeparators
    # Заголовок
    table.RepeatHeadline = True
    table.HeaderRowCount = 3
    table.Rows[0].Height = 1030
    table.Rows[0].IsAutoHeight = False
    table.Rows[1].Height = 600
    table.Rows[1].IsAutoHeight = False
    table.Rows[2].Height = 1900
    table.Rows[2].IsAutoHeight = False
    cellCursor = table.createCursorByCellName("A1")
    cellCursor.gotoCellByName("J1", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("A2", False)
    cellCursor.gotoCellByName("E2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("B2", False)
    cellCursor.gotoCellByName("F3", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("C2", False)
    cellCursor.gotoCellByName("G3", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("D2", False)
    cellCursor.gotoCellByName("H3", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("E2", False)
    cellCursor.gotoCellByName("I3", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("F2", False)
    cellCursor.gotoCellByName("J3", True)
    cellCursor.mergeRange()
    headerNames = (
        ("A1", "Лист регистрации изменений"),
        ("A2", "Номера листов (страниц)"),
        ("B2", "Всего\nлистов\n(страниц)\nв докум."),
        ("C2", "№\nдокумен-\nта"),
        ("D2", "Входящий\n№ сопрово-\nдительно-\nго докум.\nи дата"),
        ("E2", "Подп."),
        ("F2", "Да-\nта"),
        ("A3", "Изм."),
        ("B3", "изменен-\nных"),
        ("C3", "заменен-\nных"),
        ("D3", "новых"),
        ("E3", "аннули-\nрован-\nных")
    )
    for cellName, headerName in headerNames:
        cell = table.getCellByName(cellName)
        cellCursor = cell.createTextCursor()
        cellCursor.ParaStyleName = "Заголовок графы таблицы"
        fontSize = 16
        if cellName == "A1":
            fontSize = 18
        cellCursor.CharHeight = fontSize
        cellCursor.ParaAdjust = uno.Enum(
            "com.sun.star.style.ParagraphAdjust",
            "CENTER"
        )
        if cellName == "D2":
            lineSpacing = cellCursor.ParaLineSpacing
            lineSpacing.Mode = uno.getConstantByName(
                "com.sun.star.style.LineSpacingMode.FIX"
            )
            lineSpacing.Height = 490
            cellCursor.ParaLineSpacing = lineSpacing
        if cellName == "A3":
            cellCursor.CharScaleWidth = 85
        cell.TopBorderDistance = 0
        cell.BottomBorderDistance = 0
        cell.LeftBorderDistance = 0
        cell.RightBorderDistance = 0
        cell.VertOrient = uno.getConstantByName(
            "com.sun.star.text.VertOrientation.CENTER"
        )
        cell.String = headerName
    # Обрамление
    border = table.TableBorder
    noLine = uno.createUnoStruct("com.sun.star.table.BorderLine")
    normalLine = uno.createUnoStruct("com.sun.star.table.BorderLine")
    normalLine.OuterLineWidth = 50
    border.TopLine = noLine
    border.LeftLine = noLine
    border.RightLine = noLine
    border.BottomLine = normalLine
    border.HorizontalLine = normalLine
    border.VerticalLine = normalLine
    table.TableBorder = border
    # Строки
    table.Rows[3].Height = 815
    table.Rows[3].IsAutoHeight = False
    for i in range(10):
        cell = table.getCellByPosition(i, 3)
        cursor = cell.createTextCursor()
        cursor.ParaStyleName = "Значение графы таблицы"
        cell.TopBorderDistance = 0
        cell.BottomBorderDistance = 0
        cell.LeftBorderDistance = 50
        cell.RightBorderDistance = 50
        cell.VertOrient = uno.getConstantByName(
            "com.sun.star.text.VertOrientation.CENTER"
        )
    table.Rows.insertByIndex(3, 28)
    # Дабы избежать образования пустой страницы после листа рег.изм.
    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.ParaStyleName = "Пустой"
    doc.refresh()
    viewCursor = doc.CurrentController.ViewCursor
    viewCursor.gotoRange(table.getCellByName("A4").Start, False)
    doc.UndoManager.unlock()
    doc.UndoManager.clear()
    doc.unlockControllers()
    SKIP_MODIFY_EVENTS = False
    return

def removeRevTable():
    """Удалить таблицу регистрации изменений."""
    doc = XSCRIPTCONTEXT.getDocument()
    if "Лист_регистрации_изменений" not in doc.TextTables:
        return
    global SKIP_MODIFY_EVENTS
    SKIP_MODIFY_EVENTS = True
    doc.lockControllers()
    doc.UndoManager.lock()
    doc.TextTables["Лист_регистрации_изменений"].dispose()
    cursor = doc.Text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.goLeft(1, True)
    cursor.String = ""
    doc.refresh()
    viewCursor = doc.CurrentController.ViewCursor
    viewCursor.gotoStart(False)
    doc.UndoManager.unlock()
    doc.UndoManager.clear()
    doc.unlockControllers()
    SKIP_MODIFY_EVENTS = False
    return

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
    for name in STAMP_COMMON_FIELDS:
        if ("Перв.1: " + name) not in doc.TextFrames:
            continue
        firstFrame = doc.TextFrames["Перв.1: " + name]
        for prefix in ("Титул2: ", "Прочие: ", "ПрочиеА3: ", "РегИзм: "):
            if (prefix + name) not in doc.TextFrames:
                continue
            otherFrame = doc.TextFrames[prefix + name]
            otherFrame.String = firstFrame.String

            firstCursor = firstFrame.createTextCursor()
            otherCursor = otherFrame.createTextCursor()
            otherCursor.gotoEnd(True)
            otherCursor.CharHeight = firstCursor.CharHeight
            if name == "2 Обозначение документа":
                # На первом листе ширина графы 120 мм, а на последующих -- 110.
                otherCursor.CharScaleWidth = textwidth.getWidthFactor(
                    otherFrame.String,
                    otherCursor.CharHeight,
                    109
                )
            else:
                otherCursor.CharScaleWidth = firstCursor.CharScaleWidth
    doc.unlockControllers()
    doc.UndoManager.unlock()

def addPageRevTable():
    """Добавить таблицу изменений на текущем листе."""
    doc = XSCRIPTCONTEXT.getDocument()
    pageNum = doc.CurrentController.ViewCursor.Page
    pageStyle = doc.CurrentController.ViewCursor.PageStyleName
    if not pageStyle.startswith("Первый лист") \
            and pageStyle not in ("Последующие листы", "Последующие листы А3", "Лист регистрации изменений"):
        return
    global SKIP_MODIFY_EVENTS
    SKIP_MODIFY_EVENTS = True
    doc.lockControllers()
    doc.UndoManager.lock()

    # Врезка
    frame = doc.createInstance("com.sun.star.text.TextFrame")
    frame.Name = "Изм_стр_%d" % pageNum
    doc.Text.insertTextContent(doc.Text.End, frame, False)
    frame.AnchorType = uno.Enum("com.sun.star.text.TextContentAnchorType", "AT_PAGE")
    frame.AnchorPageNo = pageNum
    # Обрамление
    noLine = uno.createUnoStruct("com.sun.star.table.BorderLine")
    frame.LeftBorder = noLine
    frame.RightBorder = noLine
    frame.TopBorder = noLine
    frame.BottomBorder = noLine
    frame.BorderDistance = 0
    frame.LeftMargin = 0
    frame.RightMargin = 0
    frame.TopMargin = 0
    frame.BottomMargin = 0
    # Размер и расположение
    frame.HoriOrient = uno.getConstantByName("com.sun.star.text.HoriOrientation.NONE")
    frame.VertOrient = uno.getConstantByName("com.sun.star.text.VertOrientation.NONE")
    frame.HoriOrientRelation = 0 # Frame
    frame.VertOrientRelation = 0 # Frame
    frame.HoriOrientPosition = 23000 if pageStyle == "Последующие листы А3" else 2000
    frame.VertOrientPosition = 25200 if pageStyle.startswith("Первый лист") else 27700
    frame.Height = 1000
    frame.Width = 6500
    frame.PositionProtected = True
    frame.SizeProtected = True

    # Таблица
    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(2, 5)
    frame.Text.insertTextContent(frame.Text.End, table, False)
    table.Name = "Изм_таб_%d" % pageNum
    table.HoriOrient = uno.getConstantByName("com.sun.star.text.HoriOrientation.LEFT_AND_WIDTH")
    table.Width = 6500
    table.LeftMargin = 0
    table.Rows[0].Height = 500
    table.Rows[0].IsAutoHeight = False
    table.Rows[1].Height = 500
    table.Rows[1].IsAutoHeight = False
    columnSeparators = table.TableColumnSeparators
    columnSeparators[0].Position = 1076  # int((7)/65*10000)
    columnSeparators[1].Position = 2615  # int((7+10)/65*10000)
    columnSeparators[2].Position = 6153  # int((7+10+23)/65*10000)
    columnSeparators[3].Position = 8461  # int((7+10+23+15)/65*10000)
    table.TableColumnSeparators = columnSeparators
    for j in range(2):
        for i in range(5):
            cell = table.getCellByPosition(i, j)
            cursor = cell.createTextCursor()
            cursor.ParaStyleName = "Значение графы форматной рамки"
            cell.TopBorderDistance = 0
            cell.BottomBorderDistance = 0
            cell.LeftBorderDistance = 0
            cell.RightBorderDistance = 0
            cell.VertOrient = uno.getConstantByName(
                "com.sun.star.text.VertOrientation.CENTER"
            )
    # Обрамление
    border = table.TableBorder
    noLine = uno.createUnoStruct("com.sun.star.table.BorderLine")
    border.TopLine = noLine
    border.LeftLine = noLine
    border.RightLine = noLine
    border.BottomLine = noLine
    border.HorizontalLine = noLine
    border.VerticalLine = noLine
    table.TableBorder = border
    doc.refresh()
    viewCursor = doc.CurrentController.ViewCursor
    viewCursor.gotoRange(table.getCellByName("A2").Start, False)
    doc.UndoManager.unlock()
    doc.UndoManager.clear()
    doc.unlockControllers()
    SKIP_MODIFY_EVENTS = False
    return

def removePageRevTable():
    """Удалить таблицу изменений на текущем листе."""
    doc = XSCRIPTCONTEXT.getDocument()
    frame = doc.CurrentController.Frame
    pageNum = doc.CurrentController.ViewCursor.Page
    if ("Изм_стр_%d" % pageNum) not in doc.TextFrames:
        return
    global SKIP_MODIFY_EVENTS
    SKIP_MODIFY_EVENTS = True
    isEmpty = True
    for j in range(2):
        for i in range(5):
            cell = doc.TextTables["Изм_таб_%d" % pageNum].getCellByPosition(i, j)
            if cell.String != "":
                isEmpty = False
                viewCursor = doc.CurrentController.ViewCursor
                viewCursor.gotoRange(cell.Start, False)
                doc.refresh()
                break
    doRemove = True
    if not isEmpty:
        msgbox = frame.ContainerWindow.Toolkit.createMessageBox(
            frame.ContainerWindow,
            uno.Enum("com.sun.star.awt.MessageBoxType", "MESSAGEBOX"),
            uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_YES_NO"),
            "Внимание!",
            "Таблица с изменениями не пуста.\n"
            "Удалить?"
        )
        yes = uno.getConstantByName("com.sun.star.awt.MessageBoxResults.YES")
        result = msgbox.execute()
        if result != yes:
            doRemove = False
    if doRemove:
        if not isEmpty:
            viewCursor = doc.CurrentController.ViewCursor
            viewCursor.jumpToEndOfPage()
            doc.refresh()
        doc.lockControllers()
        doc.UndoManager.lock()
        doc.TextTables["Изм_таб_%d" % pageNum].dispose()
        doc.TextFrames["Изм_стр_%d" % pageNum].dispose()
        doc.UndoManager.unlock()
        doc.UndoManager.clear()
        doc.unlockControllers()
    SKIP_MODIFY_EVENTS = False
    return
