import sys
import threading
import tempfile
import uno
import unohelper
from com.sun.star.awt import XActionListener

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
config = sys.modules["config" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
textwidth = sys.modules["textwidth" + XSCRIPTCONTEXT.getDocument().RuntimeUID]


class ButtonStopActionListener(unohelper.Base, XActionListener):
    def __init__(self, stopEvent):
        self.stopEvent = stopEvent

    def actionPerformed(self, event):
        self.stopEvent.set()


class IndexBuildingThread(threading.Thread):
    """Перечень заполняется из отдельного вычислительного потока.

    Из-за особенностей реализации uno-интерфейса, процесс построения перечня
    занимает значительное время. Чтобы избежать продолжительного зависания
    графического интерфейса LibreOffice, перечень заполняется из отдельного
    вычислительного потока и внесённые изменения сразу же отображаются в окне
    текстового редактора.

    """
    def __init__(self):
        threading.Thread.__init__(self)
        self.name = "BuildingThread"
        self.stopEvent = threading.Event()

    def run(self):
        doc = XSCRIPTCONTEXT.getDocument()
        # --------------------------------------------------------------------
        # Диалоговое окно прогресса
        # --------------------------------------------------------------------
        context = XSCRIPTCONTEXT.getComponentContext()

        dialogModel = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel",
            context
        )
        dialogModel.Width = 200
        dialogModel.Height = 70
        dialogModel.PositionX = 0
        dialogModel.PositionY = 0
        dialogModel.Title = "Прогресс: 0%"

        labelModel = dialogModel.createInstance(
            "com.sun.star.awt.UnoControlFixedTextModel"
        )
        labelModel.PositionX = 0
        labelModel.PositionY = 0
        labelModel.Width = dialogModel.Width
        labelModel.Height = 30
        labelModel.Align = 1
        labelModel.VerticalAlign = uno.Enum(
            "com.sun.star.style.VerticalAlignment",
            "MIDDLE"
        )
        labelModel.Name = "Label"
        labelModel.Label = "Выполняется построение перечня элементов"
        dialogModel.insertByName("Label", labelModel)

        progressBarModel = dialogModel.createInstance(
            "com.sun.star.awt.UnoControlProgressBarModel"
        )
        progressBarModel.PositionX = 4
        progressBarModel.PositionY = labelModel.Height
        progressBarModel.Width = dialogModel.Width - 8
        progressBarModel.Height = 12
        progressBarModel.Name = "ProgressBar"
        progressBarModel.ProgressValue = 0
        progressBarModel.ProgressValueMin = 0
        progressBarModel.ProgressValueMax = 1
        dialogModel.insertByName("ProgressBar", progressBarModel)

        bottonModelStop = dialogModel.createInstance(
            "com.sun.star.awt.UnoControlButtonModel"
        )
        bottonModelStop.Width = 45
        bottonModelStop.Height = 16
        bottonModelStop.PositionX = (dialogModel.Width - bottonModelStop.Width) / 2
        bottonModelStop.PositionY = dialogModel.Height - bottonModelStop.Height - 5
        bottonModelStop.Name = "ButtonStop"
        bottonModelStop.Label = "Прервать"
        dialogModel.insertByName("ButtonStop", bottonModelStop)

        dialog = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog",
            context
        )
        dialog.setVisible(False)
        dialog.setModel(dialogModel)
        dialog.getControl("ButtonStop").addActionListener(
            ButtonStopActionListener(self.stopEvent)
        )
        toolkit = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.Toolkit",
            context
        )
        dialog.createPeer(toolkit, None)
        # Установить диалоговое окно по центру
        windowPosSize = doc.CurrentController.Frame.ContainerWindow.getPosSize()
        dialogPosSize = dialog.getPosSize()
        dialog.setPosSize(
            (windowPosSize.Width - dialogPosSize.Width) / 2,
            (windowPosSize.Height - dialogPosSize.Height) / 2,
            dialogPosSize.Width,
            dialogPosSize.Height,
            uno.getConstantByName("com.sun.star.awt.PosSize.POS")
        )

        # --------------------------------------------------------------------
        # Методы для построения таблицы
        # --------------------------------------------------------------------

        def kickProgress():
            nonlocal progress
            progress += 1
            dialog.getControl("ProgressBar").setValue(progress)
            dialog.setTitle("Прогресс: {:.0f}%".format(
                100*progress/progressTotal
            ))
            if self.stopEvent.is_set():
                dialog.dispose()
                return False
            return True

        def nextRow():
            nonlocal lastRow
            lastRow += 1
            table.Rows.insertByIndex(lastRow, 1)

        def getFontSize(col):
            nonlocal lastRow
            cell = table.getCellByPosition(col, lastRow)
            cellCursor = cell.createTextCursor()
            return cellCursor.CharHeight

        def fillRow(values, isTitle=False):
            colWidth = [19, 109, 9, 44]
            extraRow = [""] * len(values)
            extremeWidthFactor = config.getint("index", "extreme width factor")
            doc.lockControllers()
            for col in range(len(values)):
                if '\n' in values[col]:
                    text = values[col]
                    lfPos = text.find('\n')
                    values[col] = text[:lfPos]
                    extraRow[col] = text[(lfPos + 1):]
                widthFactor = textwidth.getWidthFactor(
                    values[col],
                    getFontSize(col),
                    colWidth[col]
                )
                if widthFactor < extremeWidthFactor:
                    text = values[col]
                    extremePos = int(len(text) * widthFactor / extremeWidthFactor)
                    # Первая попытка: определить длину не превышающую критическое
                    # сжатие шрифта.
                    pos = text.rfind(" ", 0, extremePos)
                    if pos == -1:
                        # Вторая попытка: определить длину, которая хоть и
                        # превышает критическое значение, но всё же меньше
                        # максимального.
                        pos = text.find(" ", extremePos)
                    if pos != -1:
                        values[col] = text[:pos]
                        extraRow[col] = text[(pos + 1):] + extraRow[col]
                        widthFactor = textwidth.getWidthFactor(
                            values[col],
                            getFontSize(col),
                            colWidth[col]
                        )
                cell = table.getCellByPosition(col, lastRow)
                cellCursor = cell.createTextCursor()
                if col == 1 and isTitle:
                    cellCursor.ParaStyleName = "Наименование (заголовок)"
                # Параметры символов необходимо устанавливать после
                # параметров абзаца!
                cellCursor.CharScaleWidth = widthFactor
                cell.String = values[col]
            doc.unlockControllers()

            nextRow()
            if any(extraRow):
                fillRow(extraRow, isTitle)

        # --------------------------------------------------------------------
        # Начало построения таблицы
        # --------------------------------------------------------------------

        schematic = common.getSchematicData()
        if schematic is None:
            return
        dialog.setVisible(True)
        clean(force=True)
        table = doc.TextTables["Перечень_элементов"]
        compGroups = schematic.getGroupedComponents()
        prevGroup = None
        emptyRowsRef = config.getint("index", "empty rows between diff ref")
        emptyRowsType = config.getint("index", "empty rows between diff type")
        lastRow = table.Rows.Count - 1
        # В процессе заполнения перечня, в конце таблицы всегда должна
        # оставаться пустая строка с ненарушенным форматированием.
        # На её основе будут создаваться новые строки.
        # По окончанию, последняя строка будет удалена.
        table.Rows.insertByIndex(lastRow, 1)

        progress = 0
        progressTotal = 3
        for group in compGroups:
            progressTotal += len(group)
        dialog.getControl("ProgressBar").setRange(0, progressTotal)

        for group in compGroups:
            if prevGroup is not None:
                emptyRows = 0
                if group[0].getRefType() != prevGroup[-1].getRefType():
                    emptyRows = emptyRowsRef
                else:
                    emptyRows = emptyRowsType
                for _ in range(emptyRows):
                    doc.lockControllers()
                    nextRow()
                    doc.unlockControllers()
            if len(group) == 1 \
                and not config.getboolean("index", "every group has title"):
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
                        [compRef, name, str(len(group[0])), compComment]
                    )
                    if not kickProgress():
                        return
            else:
                titleLines = group.getTitle()
                for title in titleLines:
                    if title:
                        fillRow(
                            ["", title, "", ""],
                            isTitle=True
                        )
                if config.getboolean("index", "empty row after group title"):
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
                    if not kickProgress():
                        return
            prevGroup = group

        table.Rows.removeByIndex(lastRow, 2)

        if not kickProgress():
            return

        if config.getboolean("index", "prohibit titles at bottom"):
            firstPageStyleName = doc.Text.createTextCursor().PageDescName
            firstRowCount = 28
            otherRowCount = 32
            if firstPageStyleName.endswith("3") \
                or firstPageStyleName.endswith("4"):
                    firstRowCount = 26
            pos = firstRowCount
            while pos < table.Rows.Count:
                cell = table.getCellByPosition(1, pos)
                cellCursor = cell.createTextCursor()
                if cellCursor.ParaStyleName == "Наименование (заголовок)" \
                    and cell.String != "":
                        offset = 1
                        while pos > offset:
                            cell = table.getCellByPosition(1, pos - offset)
                            cellCursor = cell.createTextCursor()
                            if cellCursor.ParaStyleName != "Наименование (заголовок)" \
                                or cell.String == "":
                                    doc.lockControllers()
                                    table.Rows.insertByIndex(pos - offset, offset)
                                    doc.unlockControllers()
                                    break
                            offset += 1
                pos += otherRowCount

        if not kickProgress():
            return

        if config.getboolean("index", "prohibit empty rows at top"):
            firstPageStyleName = doc.Text.createTextCursor().PageDescName
            firstRowCount = 29
            otherRowCount = 32
            if firstPageStyleName.endswith("3") \
                or firstPageStyleName.endswith("4"):
                    firstRowCount = 27
            pos = firstRowCount
            while pos < table.Rows.Count:
                doc.lockControllers()
                while True:
                    rowIsEmpty = False
                    for i in range(4):
                        cell = table.getCellByPosition(i, pos)
                        cellCursor = cell.createTextCursor()
                        if cell.String != "":
                            break
                    else:
                        rowIsEmpty = True
                    if not rowIsEmpty:
                        break
                    table.Rows.removeByIndex(pos, 1)
                pos += otherRowCount
                doc.unlockControllers()

        if not kickProgress():
            return

        doc.lockControllers()
        for rowIndex in range(1, table.Rows.Count):
            table.Rows[rowIndex].Height = common.getIndexRowHeight(rowIndex)
        doc.unlockControllers()

        if not kickProgress():
            return

        if config.getboolean("index", "append rev table"):
            pageCount = doc.CurrentController.PageCount
            if pageCount > config.getint("index", "pages rev table"):
                common.appendRevTable()

        dialog.dispose()


def clean(*args, force=False):
    """Очистить перечень элементов.

    Удалить всё содержимое из таблицы перечня элементов, оставив только
    заголовок и одну пустую строку.

    """
    if not force and common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    doc.lockControllers()
    text = doc.Text
    cursor = text.createTextCursor()
    firstPageStyleName = cursor.PageDescName
    text.String = ""
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
    text.insertTextContent(text.End, table, False)
    table.Name = "Перечень_элементов"
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
    table.Rows[0].Height = 1500
    table.Rows[0].IsAutoHeight = False
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
        if cellName == "A1":
            lineSpacing = uno.createUnoStruct("com.sun.star.style.LineSpacing")
            lineSpacing.Height = 87
            cellCursor.ParaLineSpacing = lineSpacing
        cell.TopBorderDistance = 50
        cell.BottomBorderDistance = 50
        cell.LeftBorderDistance = 50
        cell.RightBorderDistance = 50
        cell.VertOrient = uno.getConstantByName(
            "com.sun.star.text.VertOrientation.CENTER"
        )
        cell.String = headerName
    # Строки
    table.Rows[1].Height = 800
    table.Rows[1].IsAutoHeight = False
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
    doc.unlockControllers()

def build(*args):
    """Построить перечень элементов.

    Построить перечень элементов на основе данных из файла списка цепей.

    """
    if common.isThreadWorking():
        return
    indexBuilder = IndexBuildingThread()
    indexBuilder.start()

def toggleRevTable(*args):
    """Добавить/удалить таблицу регистрации изменений"""
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    config.set("index", "append rev table", "no")
    config.save()
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()
