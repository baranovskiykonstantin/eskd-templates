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
        try:
            doc = XSCRIPTCONTEXT.getDocument()
            doc.UndoManager.lock()
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

            def fillRow(values, isTitle=False):
                colWidth = (19, 109, 9, 44)
                extraRow = [""] * len(values)
                extremeWidthFactor = config.getint("index", "extreme width factor")
                doc.lockControllers()
                for col in range(len(values)):
                    if values[col] == "":
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
                        compType = group[0].getIndexValue("type", singular=True)
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
                        kickProgress()
                else:
                    titleLines = group.getTitle()
                    for title in titleLines:
                        if title:
                            fillRow(
                                ["", title],
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
                        kickProgress()
                prevGroup = group

            table.Rows.removeByIndex(lastRow, 2)

            kickProgress()

            if config.getboolean("index", "prohibit titles at bottom"):
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
                                if cellCursor.ParaStyleName != "Наименование (заголовок)" \
                                    or cell.String == "":
                                        doc.lockControllers()
                                        table.Rows.insertByIndex(pos - offset, offset)
                                        doc.unlockControllers()
                                        break
                                offset += 1
                    pos += otherRowCount

            kickProgress()

            if config.getboolean("index", "prohibit empty rows at top"):
                _, firstRowCount, otherRowCount = common.getFirstPageInfo()
                pos = firstRowCount + 1
                while pos < table.Rows.Count:
                    doc.lockControllers()
                    while pos < table.Rows.Count and isRowEmpty(pos):
                        table.Rows.removeByIndex(pos, 1)
                    pos += otherRowCount
                    doc.unlockControllers()

            kickProgress()

            doc.lockControllers()
            for rowIndex in range(1, table.Rows.Count):
                table.Rows[rowIndex].Height = common.getTableRowHeight(rowIndex)
            doc.unlockControllers()

            kickProgress()

            if config.getboolean("index", "append rev table"):
                pageCount = doc.CurrentController.PageCount
                if pageCount > config.getint("index", "pages rev table"):
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
                "Перечень элементов"
            )
        finally:
            if "dialog" in locals():
                dialog.dispose()
            if doc.UndoManager.isLocked():
                doc.UndoManager.unlock()
            if doc.hasControllersLocked():
                doc.unlockControllers()


def clean(*args, force=False):
    """Очистить перечень элементов.

    Удалить всё содержимое из таблицы перечня элементов, оставив только
    заголовок и одну пустую строку.

    """
    if not force and common.isThreadWorking():
        return
    common.rebuildTable()

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
