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


class SpecBuildingThread(threading.Thread):
    """Спецификация заполняется из отдельного вычислительного потока.

    Из-за особенностей реализации uno-интерфейса, процесс построения специф.
    занимает значительное время. Чтобы избежать продолжительного зависания
    графического интерфейса LibreOffice, специф. заполняется из отдельного
    вычислительного потока и внесённые изменения сразу же отображаются в окне
    текстового редактора.

    """
    def __init__(self, update=False):
        threading.Thread.__init__(self)
        self.name = "BuildingThread"
        self.update = update
        self.stopEvent = threading.Event()

    def run(self):
        try:
            # ----------------------------------------------------------------
            # Методы для построения таблицы
            # ----------------------------------------------------------------

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

            def fillSectionTitle(section):
                doc.lockControllers()
                cell = table.getCellByPosition(4, lastRow)
                cellCursor = cell.createTextCursor()
                cellCursor.ParaStyleName = "Наименование (заголовок раздела)"
                cell.String = section
                nextRow()
                doc.unlockControllers()

            def fillRow(values, isTitle=False, posIncrement=0):
                nonlocal posValue
                colWidth = (5, 5, 7, 69, 62, 9, 21)
                extraRow = [""] * len(values)
                extremeWidthFactor = config.getint("spec", "extreme width factor")
                doc.lockControllers()
                for col in range(len(values)):
                    if values[col] == "" and not (col == 2 and posIncrement != 0):
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
                    if col == 4 and isTitle:
                        cellCursor.ParaStyleName = "Наименование (заголовок группы)"
                    # Параметры символов необходимо устанавливать после
                    # параметров абзаца!
                    cellCursor.CharScaleWidth = widthFactor
                    if col == 2 and posIncrement:
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
            if self.update:
                if "Спецификация" not in doc.TextTables:
                    common.showMessage(
                        "Таблица спецификации не найдена!",
                        "Ошибка"
                    )
                    return
            else:
                clean(force=True)
            table = doc.TextTables["Спецификация"]
            tableRowCount = table.Rows.Count
            lastRow = tableRowCount - 1
            posValue = 0
            if self.update:
                otherPartsFirstRow = 0
                otherPartsLastRow = 0
                for rowIndex in range(tableRowCount):
                    if otherPartsFirstRow == 0:
                        cellPos = table.getCellByPosition(2, rowIndex).String
                        if cellPos.isdecimal():
                            posValue = int(cellPos)
                    cell = table.getCellByPosition(4, rowIndex)
                    cellCursor = cell.createTextCursor()
                    if cellCursor.ParaStyleName == "Наименование (заголовок раздела)":
                        if cell.String == "Прочие изделия":
                            otherPartsFirstRow = rowIndex
                        elif otherPartsFirstRow != 0:
                            # Следующий раздел после Прочих изделий
                            otherPartsLastRow = rowIndex - 1
                            break
                else:
                    # Прочие изделия - последний раздел в спецификации
                    otherPartsLastRow = tableRowCount - 1
                if otherPartsFirstRow == 0:
                    common.showMessage(
                        "Раздел \"Прочие изделия\" не найден!",
                        "Ошибка"
                    )
                    return
            compGroups = schematic.getGroupedComponents()
            prevGroup = None
            emptyRowsType = config.getint("spec", "empty rows between diff type")

            progressTotal = 5 if self.update else 8
            if config.getboolean("sections", "other parts"):
                for group in compGroups:
                    progressTotal += len(group)
            if self.update:
                progressDialog = ProgressDialog(
                    "Выполняется обновление раздела \"Прочие изделия\"",
                    progressTotal
                )
            else:
                progressDialog = ProgressDialog(
                    "Выполняется построение спецификации",
                    progressTotal
                )

            if self.update:
                # Удалить содержимое раздела
                table.Rows.removeByIndex(
                    otherPartsFirstRow + 1,
                    otherPartsLastRow - otherPartsFirstRow
                )

                progressDialog.stepUp()

                # Очистить содержимое и форматирование для дальнейшего заполнения
                colStyles = (
                    "Формат",
                    "Зона",
                    "Поз.",
                    "Обозначение",
                    "Наименование",
                    "Кол.",
                    "Примечание"
                )
                for colIndex in range(len(colStyles)):
                    cell = table.getCellByPosition(colIndex, otherPartsFirstRow)
                    cell.String = ""
                    cellCursor = cell.createTextCursor()
                    cellCursor.ParaStyleName = colStyles[colIndex]
                # Если за прочими изделиями следует другой раздел,
                # необходимо добавить пустую разделительную строку.
                if otherPartsLastRow != tableRowCount - 1:
                    table.Rows.insertByIndex(otherPartsFirstRow, 1)
                lastRow = otherPartsFirstRow

                progressDialog.stepUp()

            # В процессе заполнения специф., после текущей строки всегда должна
            # оставаться пустая строка с ненарушенным форматированием.
            # На её основе будут создаваться новые строки.
            # По окончанию, эта строка будет удалена.
            table.Rows.insertByIndex(lastRow, 1)

            if not self.update:
                if config.getboolean("sections", "documentation"):
                    if not config.getboolean("spec", "prohibit empty rows at top"):
                        nextRow()
                    fillSectionTitle("Документация")

                    if config.getboolean("sections", "assembly drawing") \
                        or config.getboolean("sections", "schematic") \
                        or config.getboolean("sections", "index"):
                            nextRow()

                    if config.getboolean("sections", "assembly drawing"):
                        name = "Сборочный чертёж"
                        fillRow(
                            ["", "", "", "", name]
                        )

                    if config.getboolean("sections", "schematic"):
                        size, ref = common.getSchematicInfo()
                        name = "Схема электрическая принципиальная"
                        fillRow(
                            [size, "", "", ref, name]
                        )

                    if config.getboolean("sections", "index"):
                        size, ref = common.getSchematicInfo()
                        size = "A4"
                        refParts = re.match(
                            r"([А-ЯA-Z0-9]+(?:[\.\-]\d+)+\s?)(Э\d)",
                            ref
                        )
                        if refParts is not None:
                            ref = 'П'.join(refParts.groups())
                        name = "Перечень элементов"
                        fillRow(
                            [size, "", "", ref, name]
                        )

                progressDialog.stepUp()

                if config.getboolean("sections", "assembly units"):
                    nextRow()
                    fillSectionTitle("Сборочные единицы")

                progressDialog.stepUp()

                if config.getboolean("sections", "details"):
                    nextRow()
                    fillSectionTitle("Детали")

                    if config.getboolean("sections", "pcb"):
                        nextRow()
                        size, ref = common.getPcbInfo()
                        name = "Плата печатная"
                        fillRow(
                            [size, "", "", ref, name, "1"],
                            posIncrement=1
                        )

                progressDialog.stepUp()

                if config.getboolean("sections", "standard parts"):
                    nextRow()
                    fillSectionTitle("Стандартные изделия")

                progressDialog.stepUp()

            if config.getboolean("sections", "other parts"):
                if not self.update:
                    nextRow()
                fillSectionTitle("Прочие изделия")

                nextRow()
                for group in compGroups:
                    increment = 1
                    if prevGroup is not None:
                        for _ in range(emptyRowsType):
                            doc.lockControllers()
                            nextRow()
                            doc.unlockControllers()
                        if config.getboolean("spec", "reserve position numbers"):
                            increment += emptyRowsType
                    if len(group) == 1 \
                        and not config.getboolean("spec", "every group has title"):
                            compType = group[0].getSpecValue("type", singular=True)
                            compName = group[0].getSpecValue("name")
                            compDoc = group[0].getSpecValue("doc")
                            name = ""
                            if compType:
                                name += compType + ' '
                            name += compName
                            if compDoc:
                                name += ' ' + compDoc
                            compRef = group[0].getRefRangeString()
                            compComment = group[0].getSpecValue("comment")
                            comment = compRef
                            if comment:
                                if compComment:
                                    comment = comment + '\n' + compComment
                            else:
                                comment = compComment
                            fillRow(
                                ["", "", "", "", name, str(len(group[0])), comment],
                                posIncrement=increment
                            )
                            progressDialog.stepUp()
                    else:
                        titleLines = group.getTitle()
                        for title in titleLines:
                            if title:
                                fillRow(
                                    ["", "", "", "", title],
                                    isTitle=True
                                )
                        if config.getboolean("spec", "empty row after group title"):
                            nextRow()
                            if config.getboolean("spec", "reserve position numbers"):
                                increment += 1
                        for compRange in group:
                            compName = compRange.getSpecValue("name")
                            compDoc = compRange.getSpecValue("doc")
                            name = compName
                            if compDoc:
                                for title in titleLines:
                                    if title.endswith(compDoc):
                                        break
                                else:
                                    name += ' ' + compDoc
                            compRef = compRange.getRefRangeString()
                            compComment = compRange.getSpecValue("comment")
                            comment = compRef
                            if comment:
                                if compComment:
                                    comment = comment + '\n' + compComment
                            else:
                                comment = compComment
                            fillRow(
                                ["", "", "", "", name, str(len(compRange)), comment],
                                posIncrement=increment
                            )
                            increment = 1
                            progressDialog.stepUp()
                    prevGroup = group

            if not self.update:
                if config.getboolean("sections", "materials"):
                    nextRow()
                    fillSectionTitle("Материалы")
                    nextRow()

                progressDialog.stepUp()

            table.Rows.removeByIndex(lastRow, 2)

            progressDialog.stepUp()

            if config.getboolean("spec", "prohibit titles at bottom"):
                _, firstRowCount, otherRowCount = common.getFirstPageInfo()
                pos = firstRowCount
                while pos < table.Rows.Count:
                    offset = 0
                    # Если внизу страницы пустая строка -
                    # подняться вверх к строке с данными.
                    while isRowEmpty(pos - offset) and pos > (offset + 1):
                        offset += 1
                    cell = table.getCellByPosition(4, pos - offset)
                    cellCursor = cell.createTextCursor()
                    if cellCursor.ParaStyleName.startswith("Наименование (заголовок") \
                        and cell.String != "":
                            offset += 1
                            while pos > offset:
                                cell = table.getCellByPosition(4, pos - offset)
                                cellCursor = cell.createTextCursor()
                                if not cellCursor.ParaStyleName.startswith("Наименование (заголовок") \
                                    or cell.String == "":
                                        doc.lockControllers()
                                        table.Rows.insertByIndex(pos - offset, offset)
                                        doc.unlockControllers()
                                        break
                                offset += 1
                    pos += otherRowCount

            progressDialog.stepUp()

            if config.getboolean("spec", "prohibit empty rows at top"):
                _, firstRowCount, otherRowCount = common.getFirstPageInfo()
                pos = firstRowCount + 1
                while pos < table.Rows.Count:
                    doc.lockControllers()
                    while pos < table.Rows.Count and isRowEmpty(pos):
                        table.Rows.removeByIndex(pos, 1)
                    pos += otherRowCount
                    doc.unlockControllers()

            progressDialog.stepUp()

            common.updateTableRowsHeight()

            progressDialog.stepUp()

            if config.getboolean("spec", "append rev table"):
                pageCount = doc.CurrentController.PageCount
                if pageCount > config.getint("spec", "pages rev table"):
                    common.appendRevTable()

        except StopException:
            # Прервано пользователем
            pass
        except:
            # Ошибка!
            common.showMessage(
                "При построении возникла ошибка:\n\n" \
                + traceback.format_exc(),
                "Спецификация"
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
    """Очистить спецификацию.

    Удалить всё содержимое из таблицы спецификации, оставив только
    заголовок и одну пустую строку.

    """
    if not force and common.isThreadWorking():
        return
    common.rebuildTable()

def build(*args):
    """Построить спецификацию.

    Построить спецификацию на основе данных из файла списка цепей.

    """
    if common.isThreadWorking():
        return
    specBuilder = SpecBuildingThread()
    specBuilder.start()

def update(*args):
    """Обновить "Прочие изделия".

    Обновить раздел спецификации "Прочие изделия", не изменяя при этом
    содержимого других разделов.

    """
    if common.isThreadWorking():
        return
    specUpdater = SpecBuildingThread(update=True)
    specUpdater.start()

def toggleRevTable(*args):
    """Добавить/удалить таблицу регистрации изменений"""
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()
