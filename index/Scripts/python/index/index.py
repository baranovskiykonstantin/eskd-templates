import tempfile
import uno
import common
import config
import textwidth

common.XSCRIPTCONTEXT = XSCRIPTCONTEXT
config.XSCRIPTCONTEXT = XSCRIPTCONTEXT

def clear(*args):
    """Очистить перечень элементов.

    Удалить всё содержимое из таблицы перечня элементов, оставив только
    заголовок и одну пустую строку.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.getText()
    cursor = text.createTextCursor()
    firstPageStyleName = cursor.PageDescName
    text.setString("")
    cursor.ParaStyleName = "Пустой"
    cursor.PageDescName = firstPageStyleName
    # Если не оставить параграф перед таблицей, то при изменении форматирования
    # в ячейках с автоматическими стилями будет сбрасываться стиль страницы на
    # стиль по умолчанию.
    text.insertControlCharacter(
        cursor,
        uno.getConstantByName(
            "com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"
        ),
        False
    )
    # Таблица
    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(2, 4)
    text.insertTextContent(text.getEnd(), table, False)
    table.setName("Перечень_элементов")
    table.HoriOrient = uno.getConstantByName("com.sun.star.text.HoriOrientation.LEFT_AND_WIDTH")
    table.Width = 18500
    table.LeftMargin = 2000
    columnSeparators = table.TableColumnSeparators
    columnSeparators[0].Position = 1081  # int((20)/185*10000)
    columnSeparators[1].Position = 7027  # int((20+110)/185*10000)
    columnSeparators[2].Position = 7567  # int((20+110+10)/185*10000)
    table.TableColumnSeparators = columnSeparators
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
    # Заголовок
    table.RepeatHeadline = True
    table.HeaderRowCount = 1
    table.getRows().getByIndex(0).Height = 1500
    table.getRows().getByIndex(0).IsAutoHeight = False
    headerNames = (
        ("A1", "Поз.\nобозна-\nчение"),
        ("B1", "Наименование"),
        ("C1", "Кол."),
        ("D1", "Примечание")
    )
    for cellName, headerName in headerNames:
        cell = table.getCellByName(cellName)
        cellCursor = cell.createTextCursor()
        cellCursor.ParaStyleName = "Заголовок графы таблицы"
        cell.TopBorderDistance = 0
        cell.BottomBorderDistance = 0
        cell.LeftBorderDistance = 50
        cell.RightBorderDistance = 50
        cell.VertOrient = uno.getConstantByName(
            "com.sun.star.text.VertOrientation.CENTER"
        )
        cell.setString(headerName)
    # Строки
    table.getRows().getByIndex(1).Height = 800
    table.getRows().getByIndex(1).IsAutoHeight = False
    cellStyles = (
        ("A2", "Поз. обозначение"),
        ("B2", "Наименование"),
        ("C2", "Кол."),
        ("D2", "Примечание")
    )
    for cellName, cellStyle in cellStyles:
        cell = table.getCellByName(cellName)
        cursor = cell.createTextCursor()
        cursor.ParaStyleName = cellStyle
        cell.TopBorderDistance = 0
        cell.BottomBorderDistance = 0
        cell.LeftBorderDistance = 50
        cell.RightBorderDistance = 50
        cell.VertOrient = uno.getConstantByName(
            "com.sun.star.text.VertOrientation.CENTER"
        )

def build(*args):
    """Построить перечень элементов.

    Построить перечень элементов на основе данных из файла списка цепей.

    """
    clear()
    doc = XSCRIPTCONTEXT.getDocument()
    table = doc.getTextTables().getByName("Перечень_элементов")
    settings = config.load()
    schematic = common.getSchematicData()
    if schematic is None:
        return
    compGroups = schematic.getGroupedComponents()
    prevGroup = None
    emptyRowsRef = settings.getint("index", "empty rows between diff ref")
    emptyRowsType = settings.getint("index", "empty rows between diff type")
    lastRow = table.getRows().getCount() - 1
    # В процессе заполнения перечня в конце таблицы всегда должна оставаться
    # пустая строка с ненарушенным форматированием. На её основе будут
    # создаваться новые строки. По окончанию, последняя строка будет удалена.
    table.getRows().insertByIndex(lastRow, 1)

    def nextRow():
        nonlocal lastRow
        lastRow += 1
        table.getRows().insertByIndex(lastRow, 1)

    def fillRow(values, isTitle=False):
        normValues = list(values)
        extraRef = None
        extraName = None
        extraComment = None
        extraRow = ["", "", "", ""]
        widthFactors = [100, 100, 100, 100]
        colNames = (
            "Поз. обозначение",
            "Наименование",
            "Кол.",
            "Примечание"
        )
        extremeWidthFactor = settings.getint("index", "extreme width factor")
        for index, value in enumerate(values):
            widthFactors[index] = textwidth.getWidthFactor(
                colNames[index],
                value
            )
        # Поз. обозначение
        if widthFactors[0] < extremeWidthFactor:
            ref = values[0]
            extremePos = int(len(ref) * widthFactors[0] / extremeWidthFactor)
            # Первая попытка: определить длину не превышающую критическое
            # сжатие шрифта.
            pos1 = ref.rfind(", ", 0, extremePos)
            pos2 = ref.rfind("-", 0, extremePos)
            pos = max(pos1, pos2)
            if pos == -1:
                # Вторая попытка: определить длину, которая хоть и превышает
                # критическое, но всё же меньше максимального.
                pos1 = ref.find(", ", extremePos)
                pos2 = ref.find("-", extremePos)
                pos = max(pos1, pos2)
            if pos != -1:
                separator = ref[pos]
                if separator == ",":
                    normValues[0] = ref[:(pos + 1)]
                    extraRef = ref[(pos + 2):]
                elif separator == "-":
                    normValues[0] = ref[:(pos + 1)]
                    extraRef = ref[pos:]
            widthFactors[0] = textwidth.getWidthFactor(
                colNames[0],
                normValues[0]
            )
        # Наименование
        if widthFactors[1] < extremeWidthFactor:
            name = values[1]
            extremePos = int(len(name) * widthFactors[1] / extremeWidthFactor)
            # Первая попытка: определить длину не превышающую критическое
            # сжатие шрифта.
            pos = name.rfind(" ", 0, extremePos)
            if pos == -1:
                # Вторая попытка: определить длину, которая хоть и превышает
                # критическое, но всё же меньше максимального.
                pos = name.find(" ", extremePos)
            if pos != -1:
                normValues[1] = name[:pos]
                extraName = name[(pos + 1):]
            widthFactors[1] = textwidth.getWidthFactor(
                colNames[1],
                normValues[1]
            )
        # Примечание
        if widthFactors[3] < extremeWidthFactor:
            comment = values[3]
            extremePos = int(len(comment) * widthFactors[3] / extremeWidthFactor)
            # Первая попытка: определить длину не превышающую критическое
            # сжатие шрифта.
            pos = comment.rfind(" ", 0, extremePos)
            if pos == -1:
                # Вторая попытка: определить длину, которая хоть и превышает
                # критическое, но всё же меньше максимального.
                pos = comment.find(" ", extremePos)
            if pos != -1:
                normValues[3] = comment[:pos]
                extraComment = comment[(pos + 1):]
            widthFactors[3] = textwidth.getWidthFactor(
                colNames[3],
                normValues[3]
            )

        if extraRef is not None:
            extraRow[0] = extraRef
        if extraName is not None:
            extraRow[1] = extraName
        if extraComment is not None:
            extraRow[3] = extraComment

        for i in range(4):
            cell = table.getCellByPosition(i, lastRow)
            cellCursor = cell.createTextCursor()
            if isTitle and i == 1:
                cellCursor.ParaStyleName = "Наименование (заголовок)"
            cellCursor.CharScaleWidth = widthFactors[i]
            cell.setString(normValues[i])
        nextRow()

        if any(extraRow):
            fillRow(extraRow, isTitle)

    for group in compGroups:
        if prevGroup is not None:
            emptyRows = 0
            if group[0].getRefType() != prevGroup[-1].getRefType():
                emptyRows = emptyRowsRef
            else:
                emptyRows = emptyRowsType
            for _ in range(emptyRows):
                nextRow()
        if len(group) == 1 \
            and not settings.getboolean("index", "every group has title"):
                compRef = group[0].getRefRangeString()
                compType = group[0].getTypeSingular()
                compName = group[0].getIndexValue("name")
                compDoc = group[0].getIndexValue("doc")
                name = ""
                if compType:
                    name += compType + ' '
                name += compName
                if compDoc:
                    name += ' ' + compDoc
                compComment = group[0].getIndexValue("comment")
                fillRow(
                    [compRef, name, "1", compComment]
                )
                prevGroup = group
                continue
        titleLines = group.getTitle()
        for title in titleLines:
            fillRow(
                ["", title, "", ""],
                isTitle=True
            )
        if settings.getboolean("index", "empty row after group title"):
            nextRow()
        for compRange in group:
            compRef = compRange.getRefRangeString()
            compName = compRange.getIndexValue("name")
            compDoc = compRange.getIndexValue("doc")
            name = compName
            if compDoc:
                for title in titleLines:
                    if title.endswith(compDoc):
                        break
                else:
                    name += ' ' + compDoc
            compComment = compRange.getIndexValue("comment")
            fillRow(
                [compRef, name, str(len(compRange)), compComment]
            )
        prevGroup = group

    lastRow += 1
    table.getRows().removeByIndex(lastRow, 1)

def toggleRevTable(*args):
    """Добавить/удалить таблицу регистрации изменений"""
    doc = XSCRIPTCONTEXT.getDocument()
    settings = config.load()
    settings.set("index", "append rev table", "no")
    config.save(settings)
    if doc.getTextTables().hasByName("Лист_регистрации_изменений"):
        common.removeRevTable()
    else:
        common.appendRevTable()

def help(*args):
    """Показать справочное руководство."""
    context = XSCRIPTCONTEXT.getComponentContext()
    shell = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.system.SystemShellExecute",
        context
    )
    fileAccess = context.ServiceManager.createInstance(
        "com.sun.star.ucb.SimpleFileAccess"
    )
    tempFile = tempfile.NamedTemporaryFile(
        delete=False,
        prefix="help-",
        suffix=".html"
    )
    tempFileUrl = uno.systemPathToFileUrl(tempFile.name)
    helpFileUrl = "vnd.sun.star.tdoc:/{}/Scripts/python/index/doc/help.html".format(
        XSCRIPTCONTEXT.getDocument().RuntimeUID
    )
    tempFile.close()
    fileAccess.copy(helpFileUrl, tempFileUrl)
    shell.execute(
        "file://" + tempFile.name,
        "",
        0
    )
