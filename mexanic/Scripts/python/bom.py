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


class StopException(Exception):
    pass


class ProgressDialog:
    """Диалоговое окно прогресса.

    Диалоговое окно отображает текущий прогресс построения таблицы
    и позволяет пользователю прервать операцию досрочно.

    """

    def __init__(self, message, target):
        self.stopEvent = threading.Event()
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
        dialogModel.Closeable = False

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
        labelModel.Label = message
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
            self.ButtonStopActionListener(self.stopEvent)
        )
        toolkit = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.Toolkit",
            context
        )
        dialog.createPeer(toolkit, None)
        # Установить диалоговое окно по центру
        doc = XSCRIPTCONTEXT.getDocument()
        windowPosSize = doc.CurrentController.Frame.ContainerWindow.getPosSize()
        dialogPosSize = dialog.getPosSize()
        dialog.setPosSize(
            (windowPosSize.Width - dialogPosSize.Width) / 2,
            (windowPosSize.Height - dialogPosSize.Height) / 2,
            dialogPosSize.Width,
            dialogPosSize.Height,
            uno.getConstantByName("com.sun.star.awt.PosSize.POS")
        )
        dialog.getControl("ProgressBar").setRange(0, target)
        dialog.setVisible(True)

        self.dialog = dialog
        self.progress = 0
        self.progressTotal = target

    def stepUp(self):
        if self.stopEvent.is_set():
            raise StopException
        self.progress += 1
        self.dialog.getControl("ProgressBar").setValue(self.progress)
        self.dialog.setTitle("Прогресс: {:.0f}%".format(
            100 * self.progress / self.progressTotal
        ))

    def close(self):
        self.dialog.dispose()


    class ButtonStopActionListener(unohelper.Base, XActionListener):
        def __init__(self, stopEvent):
            self.stopEvent = stopEvent

        def actionPerformed(self, event):
            self.stopEvent.set()


class BomBuildingThread(threading.Thread):
    """Ведомость заполняется из отдельного вычислительного потока.

    Из-за особенностей реализации uno-интерфейса, процесс построения ведомости
    занимает значительное время. Чтобы избежать продолжительного зависания
    графического интерфейса LibreOffice, ведомость заполняется из отдельного
    вычислительного потока и внесённые изменения сразу же отображаются в окне
    текстового редактора.

    """
    def __init__(self):
        threading.Thread.__init__(self)
        self.name = "BuildingThread"

        self.currentRow = 0
        self.currentPosition = 0

    def run(self):
        # ----------------------------------------------------------------
        # Методы для построения таблицы
        # ----------------------------------------------------------------

        def gotoNextRow(count=1):
            table.Rows.insertByIndex(self.currentRow + 1, count)
            self.currentRow += count

        def getFontSize(col):
            cell = table.getCellByPosition(col, self.currentRow)
            cellCursor = cell.createTextCursor()
            return cellCursor.CharHeight

        def isRowEmpty(row):
            lastCol = len(table.Rows[row].TableColumnSeparators)
            rowCells = table.getCellRangeByPosition(
                0, # left
                row, # top
                lastCol, # right
                row # bottom
            )
            dataIsPresent = any(rowCells.DataArray[0])
            return not dataIsPresent

        def fillRow(values, isTitle=False, posIncrement=0):
            colWidth = (6, 54, 49, 29, 9, 9, 22)
            extraRow = [""] * len(values)
            extremeWidthFactor = config.getint("doc", "extreme width factor")
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
                        extraRow[col] = text[(pos + 1):] + '\n' + extraRow[col]
                        widthFactor = textwidth.getWidthFactor(
                            values[col],
                            getFontSize(col),
                            colWidth[col]
                        )
                cell = table.getCellByPosition(col, self.currentRow)
                cellCursor = cell.createTextCursor()
                if col == 1 and isTitle:
                    cellCursor.ParaStyleName = "Наименование (заголовок)"
                # Параметры символов необходимо устанавливать после
                # параметров абзаца!
                cellCursor.CharScaleWidth = widthFactor
                if col == 0 and posIncrement \
                    and config.getboolean("doc", "only components have position numbers"):
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

                        self.currentPosition += posIncrement
                        widthFactor = textwidth.getWidthFactor(
                            str(self.currentPosition),
                            getFontSize(col),
                            colWidth[col]
                        )
                        cellCursor.gotoStart(False)
                        cellCursor.gotoEnd(True)
                        cellCursor.CharScaleWidth = widthFactor
                elif values[col]:
                    cell.String = values[col]
            doc.unlockControllers()

            gotoNextRow()
            if any(extraRow):
                fillRow(extraRow, isTitle)

        # ----------------------------------------------------------------
        # Начало построения таблицы
        # ----------------------------------------------------------------
        try:
            schematic = common.getSchematicData()
            if schematic is None:
                return
            doc = XSCRIPTCONTEXT.getDocument()
            doc.UndoManager.lock()
            clean(force=True)
            table = doc.TextTables["Ведомость_покупных_изделий"]
            self.currentRow = table.Rows.Count - 1
            compGroups = schematic.getGroupedComponents()
            prevGroup = None
            emptyRowsType = config.getint("doc", "empty rows between diff type")

            progressTotal = 6
            for group in compGroups:
                progressTotal += len(group)
            progressDialog = ProgressDialog(
                "Выполняется построение ведомости\nпокупных изделий",
                progressTotal
            )

            # В процессе заполнения ведомости, после текущей строки всегда
            # должна оставаться пустая строка с ненарушенным форматированием.
            # На её основе будут создаваться новые строки.
            # По окончанию, эта строка будет удалена.
            table.Rows.insertByIndex(self.currentRow, 1)

            for group in compGroups:
                increment = 1
                if prevGroup is not None:
                    doc.lockControllers()
                    gotoNextRow(emptyRowsType)
                    doc.unlockControllers()
                    if config.getboolean("doc", "reserve position numbers"):
                        increment += emptyRowsType
                if len(group) == 1 \
                    and not config.getboolean("doc", "every group has title"):
                        compType = group[0].getBomValue("type", singular=True)
                        compName = group[0].getBomValue("name")
                        compDoc = group[0].getBomValue("doc")
                        compDealer = group[0].getBomValue("dealer")
                        compComment = group[0].getBomValue("comment")
                        name = (compType + ' ') if compType else ""
                        name += compName
                        compCount = str(len(group[0]))
                        compCountUnits = "шт."
                        fillRow(
                            ["", name, compDoc, compDealer, compCount, compCountUnits, compComment],
                            posIncrement=increment
                        )
                        progressDialog.stepUp()
                else:
                    title = group[0].getBomValue("type", plural=True)
                    if title:
                        fillRow(
                            ["", title],
                            isTitle=True
                        )
                    if config.getboolean("doc", "empty row after group title"):
                        gotoNextRow()
                        if config.getboolean("doc", "reserve position numbers"):
                            increment += 1
                    for compRange in group:
                        compName = compRange.getBomValue("name")
                        compDoc = compRange.getBomValue("doc")
                        compDealer = compRange.getBomValue("dealer")
                        compComment = compRange.getBomValue("comment")
                        compCount = str(len(compRange))
                        compCountUnits = "шт."
                        fillRow(
                            ["", compName, compDoc, compDealer, compCount, compCountUnits, compComment],
                            posIncrement=increment
                        )
                        increment = 1
                        progressDialog.stepUp()
                prevGroup = group

            table.Rows.removeByIndex(self.currentRow, 2)

            progressDialog.stepUp()

            if config.getboolean("doc", "prohibit titles at bottom"):
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

            progressDialog.stepUp()

            if config.getboolean("doc", "prohibit empty rows at top"):
                _, firstRowCount, otherRowCount = common.getFirstPageInfo()
                pos = firstRowCount + 1
                while pos < table.Rows.Count:
                    doc.lockControllers()
                    while pos < table.Rows.Count and isRowEmpty(pos):
                        table.Rows.removeByIndex(pos, 1)
                    pos += otherRowCount
                    doc.unlockControllers()

            progressDialog.stepUp()

            if not config.getboolean("doc", "only components have position numbers"):
                doc.lockControllers()
                if "com.sun.star.text.fieldmaster.SetExpression.Позиция" in doc.TextFieldMasters:
                    posFieldMaster = doc.TextFieldMasters["com.sun.star.text.fieldmaster.SetExpression.Позиция"]
                else:
                    posFieldMaster = doc.createInstance("com.sun.star.text.fieldmaster.SetExpression")
                    posFieldMaster.SubType = 0
                    posFieldMaster.Name = "Позиция"
                for self.currentRow in range(1, table.Rows.Count):
                    posField = doc.createInstance("com.sun.star.text.textfield.SetExpression")
                    posField.Content = "Позиция+1"
                    posField.attachTextFieldMaster(posFieldMaster)
                    cell = table.getCellByPosition(0, self.currentRow)
                    cellCursor = cell.createTextCursor()
                    cell.Text.insertTextContent(cellCursor, posField, False)

                    widthFactor = textwidth.getWidthFactor(
                        str(self.currentRow - 1),
                        getFontSize(0),
                        6
                    )
                    cellCursor = cell.createTextCursor()
                    cellCursor.gotoEnd(True)
                    cellCursor.CharScaleWidth = widthFactor
                doc.unlockControllers()

            progressDialog.stepUp()

            common.updateTableRowsHeight()

            progressDialog.stepUp()

            if config.getboolean("doc", "process repeated values"):
                doc.lockControllers()
                colCount = 11
                prevValues = [""] * colCount
                repeatCount  = [0] * colCount
                for rowIndex in range(1, table.Rows.Count):
                    for colIndex in (2, 3, 6):
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

            progressDialog.stepUp()

            if config.getboolean("doc", "append rev table"):
                pageCount = doc.CurrentController.PageCount
                if pageCount > config.getint("doc", "pages rev table"):
                    common.appendRevTable()

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
            if "progressDialog" in locals():
                progressDialog.close()
            if doc.UndoManager.isLocked():
                doc.UndoManager.unlock()
            doc.UndoManager.clear()
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
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()
