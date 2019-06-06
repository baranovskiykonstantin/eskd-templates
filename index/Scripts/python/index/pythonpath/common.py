"""Вспомогательные функции и общие данные.

Модуль содержит вспомогательные функции и данные, которые могут использоваться
различными макросами.

Внимание!
Глобальную переменную XSCRIPTCONTEXT обязательно нужно установить
после импорта модуля.

"""

import os
import traceback
import uno
import kicadnet
import schematic
import config

XSCRIPTCONTEXT = None

STAMP_COMMON_FIELDS = (
    "2 Обозначение документа",
    "19 Инв. № подл.",
    "21 Взам. инв. №",
    "22 Инв. № дубл."
)

def showMessage(text, title="Сообщение"):
    """Показать текстовое сообщение.

    Аргументы:
    text -- текст сообщения;
    title -- заголовок окна сообщения.

    """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    parent = model.CurrentController.Frame.ContainerWindow
    msgbox = parent.getToolkit().createMessageBox(
        parent,
        uno.Enum("com.sun.star.awt.MessageBoxType", "MESSAGEBOX"),
        uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK"),
        title,
        text
    )
    msgbox.execute()

def showFilePicker(filePath=""):
    """Показать диалоговое окно выбора файла.

    Аргументы:

    filePath -- имя файла по умолчанию.

    Возвращаемое значение -- полное имя файла или None, если файл не выбран.

    """
    context = XSCRIPTCONTEXT.getComponentContext()
    if os.path.isfile(filePath):
        directory, file = os.path.split(filePath)
    else:
        docUrl = XSCRIPTCONTEXT.getDocument().getURL()
        if docUrl:
            directory = os.path.dirname(uno.fileUrlToSystemPath(docUrl))
        else:
            directory = os.path.expanduser('~')
        file = ""
    filePicker = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.ui.dialogs.OfficeFilePicker",
        context
    )
    filePicker.setTitle("Выбор файла с данными о схеме")
    pickerType = uno.getConstantByName(
        "com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE"
    )
    filePicker.initialize((pickerType,))
    filePicker.setDisplayDirectory(
        uno.systemPathToFileUrl(directory)
    )
    filePicker.setDefaultName(file)
    result = filePicker.execute()
    resultOK = uno.getConstantByName(
        "com.sun.star.ui.dialogs.ExecutableDialogResults.OK"
    )
    if result == resultOK:
        sourcePath = uno.fileUrlToSystemPath(filePicker.getFiles()[0])
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
    settings = config.load()
    sourcePath = settings.get("index", "source")
    if os.path.exists(sourcePath):
        return sourcePath
    sourceDir = ""
    sourceName = ""
    docUrl = XSCRIPTCONTEXT.getDocument().getURL()
    if docUrl:
        docPath = uno.fileUrlToSystemPath(docUrl)
        sourceDir = os.path.dirname(docPath)
        for fileName in os.listdir(sourceDir):
            if fileName.endswith(".pro"):
                sourceName = fileName.replace(".pro", ".net")
        if sourceName:
            sourcePath = os.path.join(sourceDir, sourceName)
            if os.path.exists(sourcePath):
                settings.set("index", "source", sourcePath)
                config.save(settings)
                return sourcePath
    sourcePath = showFilePicker(
        os.path.join(sourceDir, sourceName)
    )
    if sourcePath is not None:
        settings.set("index", "source", sourcePath)
        config.save(settings)
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
        showMessage(
            "Не удалось получить данные о схеме.",
            "Перечень элементов"
        )
        return None
    settings = config.load()
    try:
        netlist = kicadnet.Netlist(sourceFileName)
        schematicData = schematic.Schematic(settings)
        for sheet in netlist.items("sheet"):
            if sheet.attributes["name"] == "/":
                title_block = netlist.find("title_block", sheet)
                for item in title_block.items:
                    if item.name == "title":
                        schematicData.title = item.text
                    elif item.name == "company":
                        schematicData.company = item.text
                    elif item.name == "comment":
                        if item.attributes["number"] == "1":
                            schematicData.number = item.attributes["value"]
                        elif item.attributes["number"] == "2":
                            schematicData.developer = item.attributes["value"]
                        elif item.attributes["number"] == "3":
                            schematicData.verifier = item.attributes["value"]
                        elif item.attributes["number"] == "4":
                            schematicData.approver = item.attributes["value"]
                break
        for comp in netlist.items("comp"):
            component = schematic.Component(schematicData)
            component.reference = comp.attributes["ref"]
            for item in comp.items:
                if item.name == "value":
                    component.value = "" if item.text == "~" else item.text
                elif item.name == "footprint":
                    component.footprint = "" if item.text == "~" else item.text
                elif item.name == "datasheet":
                    component.datasheet = "" if item.text == "~" else item.text
                elif item.name == "fields":
                    for field in item.items:
                        fieldName = field.attributes["name"]
                        component.fields[fieldName] = field.text
            schematicData.components.append(component)
        return schematicData
    except kicadnet.ParseException as error:
        showMessage(
            "Не удалось получить данные о схеме.\n\n" \
            "При разборе файла обнаружена ошибка:\n" \
            + str(error),
            "Перечень элементов"
        )
    except:
        showMessage(
            "Не удалось получить данные о схеме.\n\n" \
            + traceback.format_exc(),
            "Перечень элементов"
        )
    return None

def appendRevTable():
    """Добавить таблицу регистрации изменений."""
    doc = XSCRIPTCONTEXT.getDocument()
    if doc.getTextTables().hasByName("Лист_регистрации_изменений"):
        return False
    text = doc.getText()
    text.insertControlCharacter(
        text.getEnd(),
        uno.getConstantByName(
            "com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"
        ),
        False
    )
    # Таблица
    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(4, 10)
    text.insertTextContent(text.getEnd(), table, False)
    table.setName("Лист_регистрации_изменений")
    table.BreakType = uno.Enum("com.sun.star.style.BreakType", "PAGE_BEFORE")
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
    table.getRows().getByIndex(0).Height = 1030
    table.getRows().getByIndex(0).IsAutoHeight = False
    table.getRows().getByIndex(1).Height = 600
    table.getRows().getByIndex(1).IsAutoHeight = False
    table.getRows().getByIndex(2).Height = 1900
    table.getRows().getByIndex(2).IsAutoHeight = False
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
        cell.setString(headerName)
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
    table.getRows().getByIndex(3).Height = 815
    table.getRows().getByIndex(3).IsAutoHeight = False
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
    table.getRows().insertByIndex(3, 28)
    # Дабы избежать образования пустой страницы после листа рег.изм.
    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.ParaStyleName = "Пустой"
    return True

def removeRevTable():
    """Удалить таблицу регистрации изменений."""
    doc = XSCRIPTCONTEXT.getDocument()
    if not doc.getTextTables().hasByName("Лист_регистрации_изменений"):
        return False
    doc.getTextTables().getByName("Лист_регистрации_изменений").dispose()
    cursor = doc.getText().createTextCursor()
    cursor.gotoEnd(False)
    cursor.goLeft(1, True)
    cursor.setString("")
    return True
