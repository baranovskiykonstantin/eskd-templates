import sys
import uno
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XWindowListener

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
config = sys.modules["config" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

def setup(*args):
    context = XSCRIPTCONTEXT.getComponentContext()

    editControlHeight = 14
    if sys.platform == "linux":
        # Элементы управления GTK3 требуют больше места
        editControlHeight = 20

    # ------------------------------------------------------------------------
    # Dialog Model
    # ------------------------------------------------------------------------

    dialogModel = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialogModel",
        context
    )
    dialogModel.Width = 300
    dialogModel.Height = 150
    dialogModel.PositionX = 0
    dialogModel.PositionY = 0
    dialogModel.Title = "Параметры пояснительной записки"

    # ------------------------------------------------------------------------
    # Tabs Model
    # ------------------------------------------------------------------------

    tabsModel = dialogModel.createInstance(
        "com.sun.star.awt.UnoMultiPageModel"
    )
    tabsModel.Width = dialogModel.Width
    tabsModel.Height = dialogModel.Height - 25
    tabsModel.PositionX = 0
    tabsModel.PositionY = 0
    tabsModel.Name = "Tabs"
    dialogModel.insertByName("Tabs", tabsModel)

    # ------------------------------------------------------------------------
    # Оптимальный вид
    # ------------------------------------------------------------------------

    checkModelSet = dialogModel.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModelSet.Width = 150
    checkModelSet.Height = 16
    checkModelSet.PositionX = 5
    checkModelSet.PositionY = dialogModel.Height - checkModelSet.Height - 4
    checkModelSet.Name = "CheckBoxSet"
    checkModelSet.State = {False: 0, True: 1}[
        config.getboolean("settings", "set view options")
    ]
    checkModelSet.Label = "Оптимальный вид документа"
    checkModelSet.HelpText = """\
Если отмечено, то при открытии документа
параметры отображения будут настроены для
обеспечения наилучшего вида содержимого:
"Границы текста" - скрыть
"Границы таблиц" - скрыть
"Затенение полей" - отключить
"Скрытые абзацы" - скрыть
"Подробные всплывающие подсказки" - вкл.
Панель инструментов - под стандартными."""
    dialogModel.insertByName("CheckBoxSet", checkModelSet)

    # ------------------------------------------------------------------------
    # Button Cancel
    # ------------------------------------------------------------------------

    buttonModelCancel = dialogModel.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModelCancel.Width = 45
    buttonModelCancel.Height = 16
    buttonModelCancel.PositionX = dialogModel.Width - buttonModelCancel.Width - 5
    buttonModelCancel.PositionY = dialogModel.Height - buttonModelCancel.Height - 5
    buttonModelCancel.Name = "ButtonCancel"
    buttonModelCancel.Label = "Отмена"
    dialogModel.insertByName("ButtonCancel", buttonModelCancel)

    # ------------------------------------------------------------------------
    # Button OK
    # ------------------------------------------------------------------------

    buttonModelOK = dialogModel.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModelOK.Width = buttonModelCancel.Width
    buttonModelOK.Height = buttonModelCancel.Height
    buttonModelOK.PositionX = buttonModelCancel.PositionX - buttonModelCancel.Width - 5
    buttonModelOK.PositionY = buttonModelCancel.PositionY
    buttonModelOK.Name = "ButtonOK"
    buttonModelOK.Label = "OK"
    dialogModel.insertByName("ButtonOK", buttonModelOK)

    # ------------------------------------------------------------------------
    # Button Import settings
    # ------------------------------------------------------------------------

    buttonModelImport = dialogModel.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModelImport.Width = buttonModelOK.Width
    buttonModelImport.Height = buttonModelOK.Height
    buttonModelImport.PositionX = buttonModelOK.PositionX - buttonModelOK.Width - 5
    buttonModelImport.PositionY = buttonModelOK.PositionY
    buttonModelImport.Name = "ButtonImport"
    buttonModelImport.Label = "Импорт..."
    dialogModel.insertByName("ButtonImport", buttonModelImport)

    # ------------------------------------------------------------------------
    # Dialog
    # ------------------------------------------------------------------------

    dialog = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialog",
        context
    )
    dialog.setModel(dialogModel)
    dialog.setPosSize(
        config.getint("settings", "pos x"),
        config.getint("settings", "pos y"),
        0,
        0,
        uno.getConstantByName("com.sun.star.awt.PosSize.POS")
    )

    # ------------------------------------------------------------------------
    # Manual Tab Model
    # ------------------------------------------------------------------------

    pageModel0 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page0", pageModel0)
    pageModel0.Title = " Пояснительная записка "

    labelModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel00.PositionX = 0
    labelModel00.PositionY = 0
    labelModel00.Width = tabsModel.Width
    labelModel00.Height = 16
    labelModel00.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel00.Name = "Label00"
    labelModel00.Label = "Файл с данными о схеме:"
    labelModel00.HelpText = """\
Источником данных о схеме является
файл списка цепей KiCad.
Поддерживаются файлы с расширением
*.net (Pcbnew) и с расширением
*.xml (вспомогательный)."""
    pageModel0.insertByName("Label00", labelModel00)

    buttonModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModel00.Width = 30
    buttonModel00.Height = editControlHeight
    buttonModel00.PositionX = tabsModel.Width - buttonModel00.Width - 3
    buttonModel00.PositionY = labelModel00.PositionY + labelModel00.Height
    buttonModel00.Name = "Button00"
    buttonModel00.Label = "Обзор"
    pageModel0.insertByName("Button00", buttonModel00)

    editControlModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel00.Width = buttonModel00.PositionX
    editControlModel00.Height = editControlHeight
    editControlModel00.PositionX = 0
    editControlModel00.PositionY = buttonModel00.PositionY
    editControlModel00.Name = "EditControl00"
    editControlModel00.Text = config.get("doc", "source")
    pageModel0.insertByName("EditControl00", editControlModel00)

    # ------------------------------------------------------------------------
    # Stamp Tab Model
    # ------------------------------------------------------------------------

    pageModel2 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page2", pageModel2)
    pageModel2.Title = " Основная надпись "

    checkModel20 = pageModel2.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel20.PositionX = 5
    checkModel20.PositionY = 5
    checkModel20.Width = tabsModel.Width - 10
    checkModel20.Height = 15
    checkModel20.Name = "CheckBox20"
    checkModel20.State = {False: 0, True: 1}[
        config.getboolean("stamp", "convert doc title")
    ]
    checkModel20.Label = "Преобразовать наименование документа"
    checkModel20.HelpText = """\
Если отмечено, тип схемы в наименовании
документа будет заменён надписью
"Пояснительная записка".
В противном случае, наименование
останется без изменений."""
    pageModel2.insertByName("CheckBox20", checkModel20)

    checkModel21 = pageModel2.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel21.PositionX = checkModel20.PositionX
    checkModel21.PositionY = checkModel20.PositionY + checkModel20.Height
    checkModel21.Width = checkModel20.Width
    checkModel21.Height = checkModel20.Height
    checkModel21.Name = "CheckBox21"
    checkModel21.State = {False: 0, True: 1}[
        config.getboolean("stamp", "convert doc id")
    ]
    checkModel21.Label = "Преобразовать обозначение документа"
    checkModel21.HelpText = """\
Если отмечено, вместо типа схемы
в обозначении документа будет
указан код "ПЗ".
В противном случае, обозначение
останется без изменений."""
    pageModel2.insertByName("CheckBox21", checkModel21)

    checkModel22 = pageModel2.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel22.PositionX = checkModel20.PositionX
    checkModel22.PositionY = checkModel20.PositionY + checkModel20.Height * 2
    checkModel22.Width = checkModel20.Width
    checkModel22.Height = checkModel20.Height
    checkModel22.Name = "CheckBox22"
    checkModel22.State = {False: 0, True: 1}[
        config.getboolean("stamp", "fill first usage")
    ]
    checkModel22.Label = "Автоматически заполнить графу \"Перв. примен.\""
    checkModel22.HelpText = """\
Если отмечено, в графу первичной
применяемости будет записано
обозначение документа без кода
документа.
В противном случае, графа
останется без изменений."""
    pageModel2.insertByName("CheckBox22", checkModel22)

    # ------------------------------------------------------------------------
    # Action Listeners
    # ------------------------------------------------------------------------

    dialog.addWindowListener(DialogWindowListener(dialog))
    buttonImport = dialog.getControl("ButtonImport")
    buttonImport.addActionListener(ButtonImportActionListener(dialog))
    buttonOK = dialog.getControl("ButtonOK")
    buttonOK.addActionListener(ButtonOKActionListener(dialog))
    ButtonCancel = dialog.getControl("ButtonCancel")
    ButtonCancel.addActionListener(ButtonCancelActionListener(dialog))

    Button00 = dialog.getControl("Tabs").getControl("Page0").getControl("Button00")
    Button00.addActionListener(Button00ActionListener(dialog))

    # ------------------------------------------------------------------------

    toolkit = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.Toolkit",
        context
    )
    dialog.createPeer(toolkit, None)
    dialog.execute()


class DialogWindowListener(unohelper.Base, XWindowListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def windowMoved(self, event):
        config.set("settings", "pos x", str(event.X))
        config.set("settings", "pos y", str(event.Y))

    def windowHidden(self, event):
        config.save()


class ButtonImportActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        self.dialog.endExecute()
        docName = common.showFilePicker(
            "",
            "Выбор документа для импорта параметров",
            **{"Текстовые документы": "*.odt", "Все файлы": "*.*"}
        )
        if docName:
            try:
                n = config.importFromDoc(docName)
                common.showMessage(
                    "Загружено {} параметров.".format(n),
                    "Импорт параметров"
                )

            except config.ImportIniNotExists:
                common.showMessage(
                    "Выбранный документ не содержит файл параметров.\n" + \
                    "В нём используются значения по умолчанию.",
                    "Ошибка импорта параметров"
                )
            except config.ImportBadDoc:
                common.showMessage(
                    "Выбранный документ не является zip архивом.",
                    "Ошибка импорта параметров"
                )
            except:
                common.showMessage(
                    "Не удалось загрузить параметры.",
                    "Ошибка импорта параметров"
                )


class ButtonOKActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        page0 = self.dialog.getControl("Tabs").getControl("Page0")
        page2 = self.dialog.getControl("Tabs").getControl("Page2")

        # --------------------------------------------------------------------
        # Оптимальный вид
        # --------------------------------------------------------------------
        config.set("settings", "set view options",
            {0: "no", 1: "yes"}[self.dialog.getControl("CheckBoxSet").State]
        )

        # --------------------------------------------------------------------
        # Пояснительная записка
        # --------------------------------------------------------------------

        config.set("doc", "source",
            page0.getControl("EditControl00").Text
        )

        # --------------------------------------------------------------------
        # Основная надпись
        # --------------------------------------------------------------------

        config.set("stamp", "convert doc title",
            {0: "no", 1: "yes"}[page2.getControl("CheckBox20").State]
        )
        config.set("stamp", "convert doc id",
            {0: "no", 1: "yes"}[page2.getControl("CheckBox21").State]
        )
        config.set("stamp", "fill first usage",
            {0: "no", 1: "yes"}[page2.getControl("CheckBox22").State]
        )

        self.dialog.endExecute()


class ButtonCancelActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        self.dialog.endExecute()


class Button00ActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        editControl = self.dialog.getControl("Tabs").getControl("Page0").getControl("EditControl00")
        source = common.showFilePicker(
            editControl.Text,
            **{"Список цепей KiCad": "*.net;*.xml", "Все файлы": "*.*"}
        )
        if source is not None:
            editControl.Text = source
