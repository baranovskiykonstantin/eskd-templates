import sys
import uno
import unohelper
from com.sun.star.util import XModifyListener

def prepareEmbeddedModulesImport(*args):
    """Подготовка к импорту встроенных в документ модулей.

    При создании нового документа из шаблона, его сразу же нужно сохранить,
    чтобы получить доступ к содержимому. Пока новый документ не сохранён его
    путь не определён и равен пустой строке.

    После сохранения нового документа или при открытии ранее созданного
    документа нужно добавить путь встроенных модулей к путям поиска модулей
    python (sys.path). Это требуется выполнить один раз, после чего импорт
    встроенных модулей станет возможным и в других файлах макросов.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    if not doc.URL:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper",
            ctx
        )
        dispatchHelper.executeDispatch(
            XSCRIPTCONTEXT.getDesktop().getCurrentFrame(),
            ".uno:SaveAs",
            "",
            0,
            ()
        )
        if not doc.URL:
            desktop = XSCRIPTCONTEXT.getDesktop()
            model = desktop.getCurrentComponent()
            parent = model.CurrentController.Frame.ContainerWindow
            msgbox = parent.getToolkit().createMessageBox(
                parent,
                uno.Enum("com.sun.star.awt.MessageBoxType", "MESSAGEBOX"),
                uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK"),
                "Внимание!",
                "Для работы макросов необходимо сначала сохранить документ!"
            )
            msgbox.execute()
            prepareEmbeddedModulesImport()
            return
    sys.path.append(uno.fileUrlToSystemPath(XSCRIPTCONTEXT.getDocument().URL) + "/Scripts/python/index/pythonpath/")

prepareEmbeddedModulesImport()

import common
import config

common.XSCRIPTCONTEXT = XSCRIPTCONTEXT
config.XSCRIPTCONTEXT = XSCRIPTCONTEXT


class DocModifyListener(unohelper.Base, XModifyListener):
    """Класс для прослушивания изменений в документе."""

    def __init__(self,):
        doc = XSCRIPTCONTEXT.getDocument()
        cursor = doc.getText().createTextCursor()
        self.prevFirstPageStyleName = cursor.PageDescName
        table = doc.getTextTables().getByName("Перечень_элементов")
        self.prevTableRowsCount = table.getRows().getCount()
        self.prevPageCount = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController.PageCount

    def modified(self, event):
        """Приём сообщения об изменении в документе."""
        doc = XSCRIPTCONTEXT.getDocument()
        currentController = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController
        # Чтобы избежать рекурсивного зацикливания,
        # необходимо сначала удалить, а после изменений,
        # снова создать обработчик сообщений об изменениях.
        doc.removeModifyListener(self)

        # Высота строк подстраивается автоматически так, чтобы нижнее обрамление
        # последней строки листа совпадало с верхней линией основной надписи.
        cursor = doc.getText().createTextCursor()
        firstPageStyleName = cursor.PageDescName
        if firstPageStyleName and doc.getTextTables().hasByName("Перечень_элементов"):
            table = doc.getTextTables().getByName("Перечень_элементов")
            tableRowsCount = table.getRows().getCount()
            if firstPageStyleName != self.prevFirstPageStyleName \
                or tableRowsCount != self.prevTableRowsCount:
                    firstRowsCount = 28
                    if firstPageStyleName.endswith("3") \
                        or firstPageStyleName.endswith("4"):
                            firstRowsCount = 26
                    otherRowsCount = 32
                    for index in range(1, tableRowsCount):
                        if index <= firstRowsCount:
                            if firstPageStyleName.endswith("1") \
                                or firstPageStyleName.endswith("2"):
                                # без граф заказчика:
                                    table.getRows().getByIndex(index).Height = 827
                            else:
                                # с графами заказчика:
                                if index == firstRowsCount:
                                    table.getRows().getByIndex(index).Height = 811
                                else:
                                    table.getRows().getByIndex(index).Height = 806
                        elif (index - firstRowsCount) % otherRowsCount == 0:
                            table.getRows().getByIndex(index).Height = 834
                        else:
                            table.getRows().getByIndex(index).Height = 801
                    self.prevFirstPageStyleName = firstPageStyleName
                    self.prevTableRowsCount = tableRowsCount

                    # Автоматическое добавление/удаление
                    # таблицы регистрации изменений.
                    # Данное действие выполняется только при редактировании
                    # таблицы перечня вручную. При автоматическом построении
                    # перечня таблица регистрации изменений добавляется
                    # отдельным образом (см. index.py).
                    pageCount = currentController.PageCount
                    if pageCount != self.prevPageCount:
                        self.prevPageCount = pageCount
                        if not common.isThreadWorking():
                            settings = config.load()
                            if settings.getboolean("index", "append rev table"):
                                if doc.getTextTables().hasByName("Лист_регистрации_изменений"):
                                    pageCount -= 1
                                if pageCount > settings.getint("index", "pages rev table"):
                                    if common.appendRevTable():
                                        self.prevPageCount +=1
                                else:
                                    if common.removeRevTable():
                                        self.prevPageCount -=1

        currentFrame = currentController.ViewCursor.TextFrame
        if currentFrame is not None \
            and currentFrame.Name.startswith("1."):
                # Обновить только текущую графу
                name = currentFrame.Name[4:]
                # Есть 4 варианта оформления первого листа
                # в виде 4-х стилей страницы.
                # Поля форматной рамки хранятся в нижнем колонтитуле
                # и для каждого стиля имеется свой набор полей.
                # При редактировании, значения полей нужно синхронизировать
                # между собой.
                for i in range(1, 5):
                    if currentFrame.Name[2] == str(i):
                        continue
                    otherName = "1.{}.{}".format(i, name)
                    if doc.getTextFrames().hasByName(otherName):
                        otherFrame = doc.getTextFrames().getByName(otherName)
                        if not otherFrame.Name.endswith(".7 Лист") \
                            and not otherFrame.Name.endswith(".8 Листов"):
                                otherFrame.setString(currentFrame.getString())
                # А также, обновить поля на последующих листах
                if name in common.STAMP_COMMON_FIELDS:
                    otherFrame = doc.getTextFrames().getByName("N." + name)
                    otherFrame.setString(currentFrame.getString())
        doc.addModifyListener(self)


def init(*args):
    """Начальная настройка при открытии документа."""
    doc = XSCRIPTCONTEXT.getDocument()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    settings = config.load()
    listener = DocModifyListener()
    doc.addModifyListener(listener)
    if settings.getboolean("settings", "set view options"):
        options = (
            {
                "path": "/org.openoffice.Office.Writer/Content/NonprintingCharacter",
                "prop": "HiddenParagraph",
                "value": False,
                "command": ".uno:ShowHiddenParagraphs"
            },
            {
                "path": "/org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['LibreOffice']/DocBoundaries",
                "prop": "IsVisible",
                "value": False,
                "command": ".uno:ViewBounds"
            },
            {
                "path": "/org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['LibreOffice']/TableBoundaries",
                "prop": "IsVisible",
                "value": False,
                "command": ".uno:TableBoundaries"
            },
            {
                "path": "/org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['LibreOffice']/WriterFieldShadings",
                "prop": "IsVisible",
                "value": False,
                "command": ".uno:Marks"
            },
            {
                "path": "/org.openoffice.Office.Common/Help",
                "prop": "ExtendedTip",
                "value": True,
                "command": ".uno:ActiveHelp"
            },
        )
        configProvider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider",
            ctx
        )
        dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper",
            ctx
        )
        nodePath = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        nodePath.Name = "nodepath"
        for op in options:
            nodePath.Value = op["path"]
            configAccess = configProvider.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess",
                (nodePath,)
            )
            value = configAccess.getPropertyValue(op["prop"])
            if value != op["value"]:
                dispatchHelper.executeDispatch(
                    XSCRIPTCONTEXT.getDesktop().getCurrentFrame(),
                    op["command"],
                    "",
                    0,
                    ()
                )
        layoutManager = doc.getCurrentController().getFrame().LayoutManager
        toolbarPos = layoutManager.getElementPos(
            "private:resource/toolbar/custom_index"
        )
        if toolbarPos.X == 0 and toolbarPos.Y == 0:
            toolbarPos.Y = 2147483647
            layoutManager.dockWindow(
                "private:resource/toolbar/custom_index",
                uno.Enum("com.sun.star.ui.DockingArea", "DOCKINGAREA_DEFAULT"),
                toolbarPos
            )
