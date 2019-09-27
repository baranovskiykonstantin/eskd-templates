import sys
import uno
import unohelper
from com.sun.star.util import XModifyListener

# Декларация встроенных модулей. Они будут импортированы позже.
common = None
config = None


class DocModifyListener(unohelper.Base, XModifyListener):
    """Класс для прослушивания изменений в документе."""

    def __init__(self,):
        doc = XSCRIPTCONTEXT.getDocument()
        self.prevFirstPageStyleName = doc.getText().createTextCursor().PageDescName
        self.prevTableRowCount = doc.getTextTables().getByName("Перечень_элементов").getRows().getCount()
        self.prevPageCount = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController.PageCount

    def modified(self, event):
        """Приём сообщения об изменении в документе."""
        doc = XSCRIPTCONTEXT.getDocument()
        currentController = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController
        # Чтобы избежать рекурсивного зацикливания,
        # необходимо сначала удалить, а после изменений,
        # снова добавить обработчик сообщений об изменениях.
        doc.removeModifyListener(self)

        # Высота строк подстраивается автоматически так, чтобы нижнее
        # обрамление последней строки листа совпадало с верхней линией
        # основной надписи.
        # Данное действие выполняется только при редактировании таблицы
        # перечня вручную.
        # При автоматическом построении перечня высота строк и таблица
        # регистрации изменений обрабатываются отдельным образом
        # (см. index.py).
        firstPageStyleName = doc.getText().createTextCursor().PageDescName
        if firstPageStyleName and doc.getTextTables().hasByName("Перечень_элементов"):
            table = doc.getTextTables().getByName("Перечень_элементов")
            tableRowCount = table.getRows().getCount()
            if firstPageStyleName != self.prevFirstPageStyleName \
                or tableRowCount != self.prevTableRowCount \
                and not common.isThreadWorking():
                    doc.lockControllers()
                    for rowIndex in range(1, tableRowCount):
                        table.getRows().getByIndex(rowIndex).Height = common.getIndexRowHeight(rowIndex)
                    doc.unlockControllers()
                    self.prevFirstPageStyleName = firstPageStyleName
                    self.prevTableRowCount = tableRowCount

                    # Автоматическое добавление/удаление
                    # таблицы регистрации изменений.
                    pageCount = currentController.PageCount
                    if pageCount != self.prevPageCount:
                        self.prevPageCount = pageCount
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


def importEmbeddedModules(*args):
    """Импорт встроенных в документ модулей.

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
            parent = desktop.getCurrentComponent().CurrentController.Frame.ContainerWindow
            msgbox = parent.getToolkit().createMessageBox(
                parent,
                uno.Enum("com.sun.star.awt.MessageBoxType", "MESSAGEBOX"),
                uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_YES_NO"),
                "Внимание!",
                "Для работы макросов необходимо сначала сохранить документ.\n"
                "Продолжить?"
            )
            yes = uno.getConstantByName("com.sun.star.awt.MessageBoxResults.YES")
            result = msgbox.execute()
            if result == yes:
                return importEmbeddedModules()
            return False
    sys.path.append(uno.fileUrlToSystemPath(XSCRIPTCONTEXT.getDocument().URL) + "/Scripts/python/index/pythonpath/")
    global common
    import common
    common.XSCRIPTCONTEXT = XSCRIPTCONTEXT
    global config
    import config
    config.XSCRIPTCONTEXT = XSCRIPTCONTEXT
    return True


def init(*args):
    """Начальная настройка при открытии документа."""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper",
        ctx
    )
    if not importEmbeddedModules():
        dispatchHelper.executeDispatch(
            XSCRIPTCONTEXT.getDesktop().getCurrentFrame(),
            ".uno:CloseDoc",
            "",
            0,
            ()
        )
        return
    doc = XSCRIPTCONTEXT.getDocument()
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

g_exportedScripts = init,
