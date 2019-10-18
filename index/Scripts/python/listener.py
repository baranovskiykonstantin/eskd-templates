import sys
import uno
import unohelper
from com.sun.star.util import XModifyListener
import zipimport

EMBEDDED_MODULES = (
    "textwidth",
    "kicadnet",
    "config",
    "schematic",
    "common",
)

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

        firstPageStyleName = doc.getText().createTextCursor().PageDescName
        if firstPageStyleName and doc.getTextTables().hasByName("Перечень_элементов"):
            table = doc.getTextTables().getByName("Перечень_элементов")
            tableRowCount = table.getRows().getCount()
            if firstPageStyleName != self.prevFirstPageStyleName \
                or tableRowCount != self.prevTableRowCount:
                    self.prevFirstPageStyleName = firstPageStyleName
                    self.prevTableRowCount = tableRowCount
                    if not common.isThreadWorking():
                        # Высота строк подстраивается автоматически так, чтобы нижнее
                        # обрамление последней строки листа совпадало с верхней линией
                        # основной надписи.
                        # Данное действие выполняется только при редактировании таблицы
                        # перечня вручную.
                        # При автоматическом построении перечня высота строк и таблица
                        # регистрации изменений обрабатываются отдельным образом
                        # (см. index.py).
                        doc.lockControllers()
                        for rowIndex in range(1, tableRowCount):
                            table.getRows().getByIndex(rowIndex).Height = common.getIndexRowHeight(rowIndex)
                        doc.unlockControllers()

                        # Автоматическое добавление/удаление
                        # таблицы регистрации изменений.
                        pageCount = currentController.PageCount
                        if pageCount != self.prevPageCount:
                            self.prevPageCount = pageCount
                            if config.getboolean("index", "append rev table"):
                                if doc.getTextTables().hasByName("Лист_регистрации_изменений"):
                                    pageCount -= 1
                                if pageCount > config.getint("index", "pages rev table"):
                                    if common.appendRevTable():
                                        self.prevPageCount += 1
                                else:
                                    if common.removeRevTable():
                                        self.prevPageCount -= 1

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

        filePicker = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.ui.dialogs.OfficeFilePicker",
            ctx
        )
        filePicker.setTitle("Сохранение нового перечня элементов")
        pickerType = uno.getConstantByName(
            "com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_SIMPLE"
        )
        filePicker.initialize((pickerType,))
        path = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.util.PathSubstitution",
            ctx
        )
        homeDir = path.getSubstituteVariableValue("$(work)")
        filePicker.setDisplayDirectory(homeDir)
        filePicker.setDefaultName("Перечень элементов.odt")
        result = filePicker.execute()
        OK = uno.getConstantByName(
            "com.sun.star.ui.dialogs.ExecutableDialogResults.OK"
        )
        if result == OK:
            fileUrl = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            fileUrl.Name = "URL"
            fileUrl.Value = filePicker.getFiles()[0]

            dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper",
                ctx
            )
            dispatchHelper.executeDispatch(
                XSCRIPTCONTEXT.getDesktop().getCurrentFrame(),
                ".uno:SaveAs",
                "",
                0,
                (fileUrl,)
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
    docPath = uno.fileUrlToSystemPath(doc.URL)
    docId = doc.RuntimeUID
    modulePath = docPath + "/Scripts/python/modules/"
    importer = zipimport.zipimporter(modulePath)
    for moduleName in EMBEDDED_MODULES:
        if moduleName in sys.modules:
            # Если модуль с таким же именем был загружен ранее,
            # его необходимо удалить из списка системы импорта,
            # чтобы в последующем модуль был загружен строго из
            # указанного места.
            del sys.modules[moduleName]
        module = importer.load_module(moduleName)
        module.__name__ = moduleName + docId
        module.init(XSCRIPTCONTEXT)
        del sys.modules[moduleName]
        sys.modules[moduleName + docId] = module
    global common
    common = sys.modules["common" + docId]
    global config
    config = sys.modules["config" + docId]
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
    config.load()
    doc = XSCRIPTCONTEXT.getDocument()
    listener = DocModifyListener()
    doc.addModifyListener(listener)
    if config.getboolean("settings", "set view options"):
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

def cleanup(*args):
    """Удалить объекты встроенных модулей из системы импорта Python."""

    for moduleName in EMBEDDED_MODULES:
        moduleName += XSCRIPTCONTEXT.getDocument().RuntimeUID
        if moduleName in sys.modules:
            del sys.modules[moduleName]

g_exportedScripts = init, cleanup
