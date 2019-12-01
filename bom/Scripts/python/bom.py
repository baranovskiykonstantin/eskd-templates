import re
import sys
import traceback
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


class StopException(Exception):
    pass


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
        try:
            doc = XSCRIPTCONTEXT.getDocument()
            # ----------------------------------------------------------------
            # Диалоговое окно прогресса
            # ----------------------------------------------------------------
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

            # ----------------------------------------------------------------
            # Методы для построения таблицы
            # ----------------------------------------------------------------

            def kickProgress():
                if self.stopEvent.is_set():
                    raise StopException
                nonlocal progress
                progress += 1
                dialog.getControl("ProgressBar").setValue(progress)
                dialog.setTitle("Прогресс: {:.0f}%".format(
                    100*progress/progressTotal
                ))

            def nextRow():
                nonlocal lastRow
                lastRow += 1
                table.Rows.insertByIndex(lastRow, 1)

            def getFontSize(col):
                nonlocal lastRow
                cell = table.getCellByPosition(col, lastRow)
                cellCursor = cell.createTextCursor()
                return cellCursor.CharHeight

            def isRowEmpty(row):
                rowCells = table.getCellRangeByPosition(
                    0, # left
                    row, # top
                    table.Columns.Count - 1, # right
                    row # bottom
                )
                dataIsPresent = any(rowCells.DataArray[0])
                return not dataIsPresent

            def fillRow(values, isTitle=False, posIncrement=0):
                nonlocal posValue
                colWidth = (6, 59, 44, 69, 54, 69, 15, 15, 15, 15, 23)
                extraRow = [""] * len(values)
                extremeWidthFactor = config.getint("bom", "extreme width factor")
                doc.lockControllers()
                for col in range(len(values)):
                    if values[col] == "" and not (col == 0 and posIncrement != 0):
                        continue
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
                        # Первая попытка: определить длину не превышающую
                        # критическое сжатие шрифта.
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
                            cellCursor.gotoStart(False)
                            cellCursor.gotoEnd(True)
                            cellCursor.CharScaleWidth = widthFactor
                    elif values[col]:
                        cell.String = values[col]
                doc.unlockControllers()

                nextRow()
                if any(extraRow):
                    fillRow(extraRow, isTitle)

            # ----------------------------------------------------------------
            # Начало построения таблицы
            # ----------------------------------------------------------------

            schematic = common.getSchematicData()
            if schematic is None:
                return
            doc = XSCRIPTCONTEXT.getDocument()
            doc.UndoManager.lock()
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

            # В процессе заполнения ведомости, после текущей строки всегда
            # должна оставаться пустая строка с ненарушенным форматированием.
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
                        increment += emptyRowsType
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
                        kickProgress()
                else:
                    title = group[0].getBomValue("type", plural=True)
                    if title:
                        fillRow(
                            ["", title],
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
                        kickProgress()
                prevGroup = group

            table.Rows.removeByIndex(lastRow, 2)

            kickProgress()

            if config.getboolean("bom", "prohibit titles at bottom"):
                _, firstRowCount, otherRowCount = common.getFirstPageInfo()
                pos = firstRowCount
                while pos < table.Rows.Count:
                    offset = 0
                    # Если внизу страницы пустая строка -
                    # подняться вверх к строке с данными.
                    while isRowEmpty(pos - offset) and pos > (offset + 1):
                        offset += 1
                    cell = table.getCellByPosition(1, pos - offset)
                    cellCursor = cell.createTextCursor()
                    if cellCursor.ParaStyleName == "Наименование (заголовок)" \
                        and cell.String != "":
                            offset += 1
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

            kickProgress()

            if config.getboolean("bom", "prohibit empty rows at top"):
                _, firstRowCount, otherRowCount = common.getFirstPageInfo()
                pos = firstRowCount + 1
                while pos < table.Rows.Count:
                    doc.lockControllers()
                    while pos < table.Rows.Count and isRowEmpty(pos):
                        table.Rows.removeByIndex(pos, 1)
                    pos += otherRowCount
                    doc.unlockControllers()

            kickProgress()

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

            kickProgress()

            common.updateTableRowsHeight()

            kickProgress()

            if config.getboolean("bom", "process repeated values"):
                doc.lockControllers()
                colCount = 11
                prevValues = [""] * colCount
                repeatCount  = [0] * colCount
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

            kickProgress()

            if config.getboolean("bom", "append rev table"):
                pageCount = doc.CurrentController.PageCount
                if pageCount > config.getint("bom", "pages rev table"):
                    common.appendRevTable()

            doc.UndoManager.clear()

        except StopException:
            # Прервано пользователем
            pass

        except:
            # Ошибка!
            common.showMessage(
                "При построении возникла ошибка:\n\n" \
                + traceback.format_exc(),
                "Ведомость покупных изделий"
            )
        finally:
            if "dialog" in locals():
                dialog.dispose()
            if doc.UndoManager.isLocked():
                doc.UndoManager.unlock()
            if doc.hasControllersLocked():
                doc.unlockControllers()


def clean(*args, force=False):
    """Очистить ведомость.

    Удалить всё содержимое из таблицы ведомости, оставив только
    заголовок и одну пустую строку.

    """
    if not force and common.isThreadWorking():
        return
    common.rebuildTable()

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
