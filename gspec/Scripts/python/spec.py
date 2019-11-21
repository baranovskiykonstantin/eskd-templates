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
    def __init__(self, dialog, stopEvent):
        self.dialog = dialog
        self.stopEvent = stopEvent

    def actionPerformed(self, event):
        self.stopEvent.set()
        self.dialog.dispose()


class StopException(Exception):
    pass


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
            if self.update:
                labelModel.Label = "Выполняется обновление раздела \"Прочие изделия\""
            else:
                labelModel.Label = "Выполняется построение спецификации"
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
                ButtonStopActionListener(dialog, self.stopEvent)
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
                colWidth = (5, 5, 7, 69, 62, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 32)
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
                        cellCursor = cell.createTextCursor()
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

            progress = 0
            progressTotal = 5 if self.update else 8
            if config.getboolean("sections", "other parts"):
                for group in compGroups:
                    progressTotal += len(group)
            dialog.getControl("ProgressBar").setRange(0, progressTotal)
            dialog.setVisible(True)

            if self.update:
                # Удалить содержимое раздела
                table.Rows.removeByIndex(
                    otherPartsFirstRow + 1,
                    otherPartsLastRow - otherPartsFirstRow
                )

                kickProgress()

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

                kickProgress()

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
                            ["", "", "", "", name, "X", "", "", "", "", "", "", "", "", "", ""]
                        )

                    if config.getboolean("sections", "schematic"):
                        size, ref = common.getSchematicInfo()
                        name = "Схема электрическая принципиальная"
                        fillRow(
                            [size, "", "", ref, name, "X", "", "", "", "", "", "", "", "", "", ""]
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
                            [size, "", "", ref, name, "X", "", "", "", "", "", "", "", "", "", ""]
                        )

                kickProgress()

                if config.getboolean("sections", "assembly units"):
                    nextRow()
                    fillSectionTitle("Сборочные единицы")

                kickProgress()

                if config.getboolean("sections", "details"):
                    nextRow()
                    fillSectionTitle("Детали")

                    if config.getboolean("sections", "pcb"):
                        nextRow()
                        size, ref = common.getPcbInfo()
                        name = "Плата печатная"
                        fillRow(
                            [size, "", "", ref, name, "1", "", "", "", "", "", "", "", "", "", ""],
                            posIncrement=1
                        )

                kickProgress()

                if config.getboolean("sections", "standard parts"):
                    nextRow()
                    fillSectionTitle("Стандартные изделия")

                kickProgress()

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
                                increment += 1
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
                                ["", "", "", "", name, str(len(group[0])), "", "", "", "", "", "", "", "", "", comment],
                                posIncrement=increment
                            )
                            kickProgress()
                    else:
                        titleLines = group.getTitle()
                        for title in titleLines:
                            if title:
                                fillRow(
                                    ["", "", "", "", title, "", "", "", "", "", "", "", "", "", "", ""],
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
                                ["", "", "", "", name, str(len(compRange)), "", "", "", "", "", "", "", "", "", comment],
                                posIncrement=increment
                            )
                            increment = 1
                            kickProgress()
                    prevGroup = group

            if not self.update:
                if config.getboolean("sections", "materials"):
                    nextRow()
                    fillSectionTitle("Материалы")
                    nextRow()

                kickProgress()

            table.getRows().removeByIndex(lastRow, 2)

            kickProgress()

            if config.getboolean("spec", "prohibit titles at bottom"):
                _, firstRowCount, otherRowCount, _ = getFirstPageInfo()
                pos = firstRowCount
                while pos < table.Rows.Count:
                    cell = table.getCellByPosition(4, pos)
                    cellCursor = cell.createTextCursor()
                    if cellCursor.ParaStyleName.startswith("Наименование (заголовок") \
                        and cell.String != "":
                            offset = 1
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

            kickProgress()

            if config.getboolean("spec", "prohibit empty rows at top"):
                _, firstRowCount, otherRowCount, _ = getFirstPageInfo()
                pos = firstRowCount + 1
                while pos < table.Rows.Count:
                    doc.lockControllers()
                    while True:
                        rowIsEmpty = False
                        for i in range(7):
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

            kickProgress()

            doc.lockControllers()
            for rowIndex in range(2, table.Rows.Count):
                table.Rows[rowIndex].Height = common.getSpecRowHeight(rowIndex)
            doc.unlockControllers()

            kickProgress()

            if config.getboolean("spec", "append rev table"):
                pageCount = doc.CurrentController.PageCount
                if pageCount > config.getint("spec", "pages rev table"):
                    common.appendRevTable()

            common.updateVarTablePosition()

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
            if "dialog" in locals():
                dialog.dispose()


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
    """Добавить/удалить таблицу регистрации изменений."""
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    config.set("spec", "append rev table", "no")
    config.save()
    if "Лист_регистрации_изменений" in doc.TextTables:
        common.removeRevTable()
    else:
        common.appendRevTable()

def toggleVarTable(*args):
    """Добавить/удалить таблицу наименований исполнений."""
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    if "Наименования_исполнений" in doc.TextFrames:
        common.removeVarTable()
    else:
        common.addVarTable()
