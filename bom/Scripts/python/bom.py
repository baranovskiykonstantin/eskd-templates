import re
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


class BomBuildingThread(threading.Thread):
    """Ведомость заполняется из отдельного вычислительного потока.

    Из-за особенностей реализации uno-интерфейса, процесс построения ведомость
    занимает значительное время. Чтобы избежать продолжительного зависания
    графического интерфейса LibreOffice, ведомость заполняется из отдельного
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
        labelModel.Label = "Выполняется построение ведомости\nпокупных изделий"
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

        def fillRow(values, isTitle=False, posIncrement=0):
            nonlocal posValue
            colWidth = [6, 59, 44, 69, 54, 69, 15, 15, 15, 15, 23]
            extraRow = [""] * len(values)
            extremeWidthFactor = config.getint("bom", "extreme width factor")
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
                if col == 0 and posIncrement \
                    and config.getboolean("bom", "only components have position numbers"):
                        if "com.sun.star.text.fieldmaster.SetExpression.Позиция" in doc.TextFieldMasters:
                            posFieldMaster = doc.TextFieldMasters["com.sun.star.text.fieldmaster.SetExpression.Позиция"]
                        else:
                            posFieldMaster = doc.createInstance("com.sun.star.text.fieldmaster.SetExpression")
                            posFieldMaster.SubType = 0
                            posFieldMaster.Name = "Позиция"
                        posField = doc.createInstance("com.sun.star.text.textfield.SetExpression")
                        posField.Content = "Позиция+" + str(posIncrement)
                        posField.attachTextFieldMaster(posFieldMaster)
                        cell.Text.insertTextContent(cellCursor, posField, False)

                        posValue += posIncrement
                        widthFactor = textwidth.getWidthFactor(
                            str(posValue),
                            getFontSize(col),
                            colWidth[col]
                        )
                        cellCursor = cell.createTextCursor()
                        cellCursor.gotoEnd(True)
                        cellCursor.CharScaleWidth = widthFactor
                else:
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
        doc = XSCRIPTCONTEXT.getDocument()
        clean(force=True)
        table = doc.TextTables["Ведомость_покупных_изделий"]
        lastRow = table.Rows.Count - 1
        posValue = 0
        compGroups = schematic.getGroupedComponents()
        prevGroup = None
        emptyRowsType = config.getint("bom", "empty rows between diff type")

        progress = 0
        progressTotal = 6
        for group in compGroups:
            progressTotal += len(group)
        dialog.getControl("ProgressBar").setRange(0, progressTotal)
        dialog.setVisible(True)

        # В процессе заполнения ведомость, после текущей строки всегда должна
        # оставаться пустая строка с ненарушенным форматированием.
        # На её основе будут создаваться новые строки.
        # По окончанию, эта строка будет удалена.
        table.Rows.insertByIndex(lastRow, 1)

        for group in compGroups:
            increment = 1
            if prevGroup is not None:
                for _ in range(emptyRowsType):
                    doc.lockControllers()
                    nextRow()
                    doc.unlockControllers()
                    if config.getboolean("bom", "reserve position numbers"):
                        increment += 1
            if len(group) == 1 \
                and not config.getboolean("bom", "every group has title"):
                    compType = group[0].getBomValue("type", singular=True)
                    compName = group[0].getBomValue("name")
                    compCode = group[0].getBomValue("code")
                    compDoc = group[0].getBomValue("doc")
                    compDealer = group[0].getBomValue("dealer")
                    compForWhat = group[0].getBomValue("for what")
                    compComment = group[0].getBomValue("comment")
                    name = ""
                    if compType:
                        name += compType + ' '
                    name += compName
                    compCount = str(len(group[0]))
                    fillRow(
                        ["", name, compCode, compDoc, compDealer, compForWhat, compCount, "", "", compCount, compComment],
                        posIncrement=increment
                    )
                    if not kickProgress():
                        return
            else:
                title = group[0].getBomValue("type", plural=True)
                if title:
                    fillRow(
                        ["", title, "", "", "", "", "", "", "", "", ""],
                        isTitle=True
                    )
                if config.getboolean("bom", "empty row after group title"):
                    nextRow()
                    if config.getboolean("bom", "reserve position numbers"):
                        increment += 1
                for compRange in group:
                    compName = compRange.getBomValue("name")
                    compCode = compRange.getBomValue("code")
                    compDoc = compRange.getBomValue("doc")
                    compDealer = compRange.getBomValue("dealer")
                    compForWhat = compRange.getBomValue("for what")
                    compComment = compRange.getBomValue("comment")
                    compCount = str(len(compRange))
                    fillRow(
                        ["", compName, compCode, compDoc, compDealer, compForWhat, compCount, "", "", compCount, compComment],
                        posIncrement=increment
                    )
                    increment = 1
                    if not kickProgress():
                        return
            prevGroup = group

        table.getRows().removeByIndex(lastRow, 2)

        if not kickProgress():
            return

        if config.getboolean("bom", "prohibit titles at bottom"):
            firstPageStyleName = doc.Text.createTextCursor().PageDescName
            firstRowCount = 28
            otherRowCount = 30
            if firstPageStyleName.endswith("3") \
                or firstPageStyleName.endswith("4"):
                    firstRowCount = 25
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
                            if not cellCursor.ParaStyleName == "Наименование (заголовок)" \
                                or cell.String == "":
                                    doc.lockControllers()
                                    table.Rows.insertByIndex(pos - offset, offset)
                                    doc.unlockControllers()
                                    break
                            offset += 1
                pos += otherRowCount

        if not kickProgress():
            return

        if config.getboolean("bom", "prohibit empty rows at top"):
            firstPageStyleName = doc.Text.createTextCursor().PageDescName
            firstRowCount = 29
            otherRowCount = 30
            if firstPageStyleName.endswith("3") \
                or firstPageStyleName.endswith("4"):
                    firstRowCount = 26
            pos = firstRowCount
            while pos < table.Rows.Count:
                doc.lockControllers()
                while True:
                    rowIsEmpty = False
                    for i in range(11):
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

        if not config.getboolean("bom", "only components have position numbers"):
            doc.lockControllers()
            if "com.sun.star.text.fieldmaster.SetExpression.Позиция" in doc.TextFieldMasters:
                posFieldMaster = doc.TextFieldMasters["com.sun.star.text.fieldmaster.SetExpression.Позиция"]
            else:
                posFieldMaster = doc.createInstance("com.sun.star.text.fieldmaster.SetExpression")
                posFieldMaster.SubType = 0
                posFieldMaster.Name = "Позиция"
            for lastRow in range(2, table.Rows.Count):
                posField = doc.createInstance("com.sun.star.text.textfield.SetExpression")
                posField.Content = "Позиция+1"
                posField.attachTextFieldMaster(posFieldMaster)
                cell = table.getCellByPosition(0, lastRow)
                cellCursor = cell.createTextCursor()
                cell.Text.insertTextContent(cellCursor, posField, False)

                widthFactor = textwidth.getWidthFactor(
                    str(lastRow - 1),
                    getFontSize(0),
                    6
                )
                cellCursor = cell.createTextCursor()
                cellCursor.gotoEnd(True)
                cellCursor.CharScaleWidth = widthFactor
            doc.unlockControllers()

        if not kickProgress():
            return

        doc.lockControllers()
        for rowIndex in range(2, table.Rows.Count):
            table.Rows[rowIndex].Height = common.getBomRowHeight(rowIndex)
        doc.unlockControllers()

        if not kickProgress():
            return

        if config.getboolean("bom", "process repeated values"):
            doc.lockControllers()
            prevValues = [""] * 11
            repeatCount  = [0] * 11
            for rowIndex in range(2, table.Rows.Count):
                for colIndex in (2, 3, 4, 5, 10):
                    cell = table.getCellByPosition(colIndex, rowIndex)
                    if cell.String and cell.String == prevValues[colIndex]:
                        repeatCount[colIndex] += 1
                        if repeatCount[colIndex] == 1:
                            cell.String = "То же"
                        elif repeatCount[colIndex] > 1:
                            cell.String = '»'
                    else:
                        prevValues[colIndex] = cell.String
                        repeatCount[colIndex] = 0
            doc.unlockControllers()

        if not kickProgress():
            return

        if config.getboolean("bom", "append rev table"):
            pageCount = doc.CurrentController.PageCount
            if pageCount > config.getint("bom", "pages rev table"):
                common.appendRevTable()

        dialog.dispose()


def clean(*args, force=False):
    """Очистить ведомость.

    Удалить всё содержимое из таблицы ведомости, оставив только
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
    table.initialize(3, 11)
    text.insertTextContent(text.End, table, False)
    table.Name = "Ведомость_покупных_изделий"
    table.HoriOrient = uno.getConstantByName("com.sun.star.text.HoriOrientation.LEFT_AND_WIDTH")
    table.Width = 39500
    table.LeftMargin = 2000
    columnSeparators = table.TableColumnSeparators
    columnSeparators[0].Position = 177   # int((7)/395*10000)
    columnSeparators[1].Position = 1695  # int((7+60)/395*10000)
    columnSeparators[2].Position = 2834  # int((7+60+45)/395*10000)
    columnSeparators[3].Position = 4606  # int((7+60+45+70)/395*10000)
    columnSeparators[4].Position = 5998  # int((7+60+45+70+55)/395*10000)
    columnSeparators[5].Position = 7770  # int((7+60+45+70+55+70)/395*10000)
    columnSeparators[6].Position = 8175  # int((7+60+45+70+55+70+16)/395*10000)
    columnSeparators[7].Position = 8580  # int((7+60+45+70+55+70+16+16)/395*10000)
    columnSeparators[8].Position = 8985  # int((7+60+45+70+55+70+16+16+16)/395*10000)
    columnSeparators[9].Position = 9390  # int((7+60+45+70+55+70+16+16+16+16)/395*10000)
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
    table.HeaderRowCount = 2
    table.Rows[0].Height = 900
    table.Rows[0].IsAutoHeight = False
    table.Rows[1].Height = 1800
    table.Rows[1].IsAutoHeight = False
    cellCursor = table.createCursorByCellName("A1")
    cellCursor.gotoCellByName("A2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("B1", False)
    cellCursor.gotoCellByName("B2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("C1", False)
    cellCursor.gotoCellByName("C2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("D1", False)
    cellCursor.gotoCellByName("D2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("E1", False)
    cellCursor.gotoCellByName("E2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("F1", False)
    cellCursor.gotoCellByName("F2", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("G1", False)
    cellCursor.gotoCellByName("J1", True)
    cellCursor.mergeRange()
    cellCursor.gotoCellByName("H1", False)
    cellCursor.gotoCellByName("K2", True)
    cellCursor.mergeRange()
    headerNames = (
        ("A1", "№ строки"),
        ("B1", "Наименование"),
        ("C1", "Код\nпродукции"),
        ("D1", "Обозначение\nдокумента на\nпоставку"),
        ("E1", "Поставщик"),
        ("F1", "Куда входит\n(обозначение)"),
        ("G1", "Количество"),
        ("G2", "на из-\nделие"),
        ("H2", "в ком-\nплекты"),
        ("I2", "на ре-\nгулир."),
        ("J2", "всего"),
        ("H1", "Приме-\nчание")
    )
    for cellName, headerName in headerNames:
        cell = table.getCellByName(cellName)
        cellCursor = cell.createTextCursor()
        cellCursor.ParaStyleName = "Заголовок графы таблицы"
        cell.TopBorderDistance = 50
        cell.BottomBorderDistance = 50
        cell.LeftBorderDistance = 50
        cell.RightBorderDistance = 50
        cell.VertOrient = uno.getConstantByName(
            "com.sun.star.text.VertOrientation.CENTER"
        )
        if cellName.endswith("1"):
            cellCursor.CharHeight = 18
        if cellName == "A1":
            cellCursor.CharRotation = 900
            cell.LeftBorderDistance = 0
            cell.RightBorderDistance = 100
        cell.String = headerName
    # Строки
    table.Rows[2].Height = 800
    table.Rows[2].IsAutoHeight = False
    cellStyles = (
        ("A3", "№ строки"),
        ("B3", "Наименование"),
        ("C3", "Код продукции"),
        ("D3", "Обозначение документа на поставку"),
        ("E3", "Поставщик"),
        ("F3", "Куда входит (обозначение)"),
        ("G3", "Кол. на изделие"),
        ("H3", "Кол. в комплекты"),
        ("I3", "Кол. на регулир."),
        ("J3", "Кол. всего"),
        ("K3", "Примечание")
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
    """Построить ведомость покупных изделий.

    Построить ведомость на основе данных из файла списка цепей.

    """
    if common.isThreadWorking():
        return
    bomBuilder = BomBuildingThread()
    bomBuilder.start()

def toggleRevTable(*args):
    """Добавить/удалить таблицу регистрации изменений"""
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    config.set("bom", "append rev table", "no")
    config.save()
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()
