import sys
import uno
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XWindowListener

common = sys.modules["common" + XSCRIPTCONTEXT.getDocument().RuntimeUID]
config = sys.modules["config" + XSCRIPTCONTEXT.getDocument().RuntimeUID]

def setup(*args):
    if common.isThreadWorking():
        return
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
    dialogModel.Height = 335
    dialogModel.PositionX = 0
    dialogModel.PositionY = 0
    dialogModel.Title = "Параметры ведомости покупных изделий"

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
    # Bom Tab Model
    # ------------------------------------------------------------------------

    pageModel0 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page0", pageModel0)
    pageModel0.Title = " Ведомость "

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

    editControlModel02 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlNumericFieldModel"
    )
    editControlModel02.Width = 50
    editControlModel02.Height = editControlHeight
    editControlModel02.PositionX = 0
    editControlModel02.PositionY = editControlModel00.PositionY + editControlModel00.Height
    editControlModel02.Name = "EditControl02"
    editControlModel02.Value = config.getint("doc", "empty rows between diff type")
    editControlModel02.ValueMin = 0
    editControlModel02.ValueMax = 99
    editControlModel02.ValueStep = 1
    editControlModel02.Spin = True
    editControlModel02.DecimalAccuracy = 0
    pageModel0.insertByName("EditControl02", editControlModel02)

    labelModel02 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel02.PositionX = editControlModel02.PositionX + editControlModel02.Width
    labelModel02.PositionY = editControlModel02.PositionY
    labelModel02.Width = tabsModel.Width - labelModel02.PositionX
    labelModel02.Height = editControlModel02.Height
    labelModel02.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel02.Name = "Label02"
    labelModel02.Label = " пустых строк между компонентами разного типа"
    labelModel02.HelpText = """\
Указанное количество пустых строк будет
вставлено между компонентами различного
типа в разделе "Прочие изделия"."""
    pageModel0.insertByName("Label02", labelModel02)

    editControlModel03 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlNumericFieldModel"
    )
    editControlModel03.Width = editControlModel02.Width
    editControlModel03.Height = editControlModel02.Height
    editControlModel03.PositionX = editControlModel02.PositionX
    editControlModel03.PositionY = editControlModel02.PositionY + editControlModel02.Height
    editControlModel03.Name = "EditControl03"
    editControlModel03.Value = config.getint("doc", "extreme width factor")
    editControlModel03.ValueMin = 0
    editControlModel03.ValueMax = 99
    editControlModel03.ValueStep = 1
    editControlModel03.Spin = True
    editControlModel03.DecimalAccuracy = 0
    pageModel0.insertByName("EditControl03", editControlModel03)

    labelModel03 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel03.PositionX = editControlModel03.PositionX + editControlModel03.Width
    labelModel03.PositionY = editControlModel03.PositionY
    labelModel03.Width = tabsModel.Width - labelModel03.PositionX
    labelModel03.Height = editControlModel03.Height
    labelModel03.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel03.Name = "Label03"
    labelModel03.Label = " минимально допустимый масштаб по ширине (%)"
    labelModel03.HelpText = """\
Если текст не помещается в графе таблицы,
то уменьшается масштаб символов по ширине.
Когда масштаб становится меньше указанного
значения, текст разбивается на части и
размещается на последующих строках."""
    pageModel0.insertByName("Label03", labelModel03)

    checkModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel00.PositionX = 5
    checkModel00.PositionY = editControlModel03.PositionY + editControlModel03.Height + 5
    checkModel00.Width = tabsModel.Width - 10
    checkModel00.Height = 15
    checkModel00.Name = "CheckBox00"
    checkModel00.State = {False: 0, True: 1}[
        config.getboolean("doc", "add units")
    ]
    checkModel00.Label = "Добавить единицы измерения"
    checkModel00.HelpText = """\
Если для резисторов, конденсаторов или
индуктивностей указаны только значения
и данная опция включена, то к значениям
будут добавлены соответствующие единицы
измерения (Ом, Ф, Гн).
При этом, множители приводятся к
общему виду."""
    pageModel0.insertByName("CheckBox00", checkModel00)

    checkModel01 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel01.PositionX = 15
    checkModel01.PositionY = checkModel00.PositionY + checkModel00.Height
    checkModel01.Width = tabsModel.Width - 20
    checkModel01.Height = checkModel00.Height
    checkModel01.Name = "CheckBox01"
    checkModel01.State = \
        {False: 0, True: 1}[config.getboolean("doc", "space before units")]
    checkModel01.Label = "Вставить пробел перед единицами измерения"
    checkModel01.HelpText = """\
Если отмечено, то между цифровой
частью значения и единицами измерения
(включая множитель) будет вставлен пробел."""
    pageModel0.insertByName("CheckBox01", checkModel01)

    checkModel02 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel02.PositionX = 5
    checkModel02.PositionY = checkModel01.PositionY + checkModel00.Height
    checkModel02.Width = tabsModel.Width - 10
    checkModel02.Height = checkModel00.Height
    checkModel02.Name = "CheckBox02"
    checkModel02.State = \
        {False: 0, True: 1}[config.getboolean("doc", "separate group for each doc")]
    checkModel02.Label = "Формировать отдельную группу для каждого документа"
    checkModel02.HelpText = """\
По умолчанию, группы компонентов
формируются по их типу, например:
"Резисторы", "Конденсаторы" и т.д.
Если отмечено, то группы компонентов
будут разбиваться ещё и по документу,
например:
"Резисторы ГОСТ...", "Резисторы ТУ..."
и т.д."""
    pageModel0.insertByName("CheckBox02", checkModel02)

    checkModel04 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel04.PositionX = 5
    checkModel04.PositionY = checkModel02.PositionY + checkModel00.Height
    checkModel04.Width = tabsModel.Width - 10
    checkModel04.Height = checkModel00.Height
    checkModel04.Name = "CheckBox04"
    checkModel04.State = \
        {False: 0, True: 1}[config.getboolean("doc", "every group has title")]
    checkModel04.Label = "Формировать заголовок для каждой группы"
    checkModel04.HelpText = """\
По умолчанию, заголовок формируется
только если группа содержит более чем
один компонент.
Если же группа состоит из одного
компонента, заголовок не формируется,
а тип, в единственном числе,
указывается перед наименованием.
Если отмечено, то заголовок будет
сформирован для каждой группы, даже
если она состоит из одного компонента."""
    pageModel0.insertByName("CheckBox04", checkModel04)

    checkModel010 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel010.PositionX = 5
    checkModel010.PositionY = checkModel04.PositionY + checkModel00.Height
    checkModel010.Width = tabsModel.Width - 10
    checkModel010.Height = checkModel00.Height
    checkModel010.Name = "CheckBox010"
    checkModel010.State = \
        {False: 0, True: 1}[config.getboolean("doc", "only components have position numbers")]
    checkModel010.Label = "Нумеровать только позиции компонентов"
    checkModel010.HelpText = """\
По умолчанию, номера позиций
присваиваются каждой строке.
Если отмечено, то номера позиций
будут указываться только для
компонентов."""
    pageModel0.insertByName("CheckBox010", checkModel010)

    checkModel09 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel09.PositionX = 15
    checkModel09.PositionY = checkModel010.PositionY + checkModel00.Height
    checkModel09.Width = tabsModel.Width - 10
    checkModel09.Height = checkModel00.Height
    checkModel09.Name = "CheckBox09"
    checkModel09.State = \
        {False: 0, True: 1}[config.getboolean("doc", "reserve position numbers")]
    checkModel09.Label = "Резервировать номера позиций"
    checkModel09.HelpText = """\
По умолчанию, позиции в ведомости
увеличиваются на единицу.
Если отмечено, то для пустых строк,
вставляемых между группами компонентов,
будут зарезервированы номера позиций."""
    pageModel0.insertByName("CheckBox09", checkModel09)

    checkModel05 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel05.PositionX = 5
    checkModel05.PositionY = checkModel09.PositionY + checkModel00.Height
    checkModel05.Width = tabsModel.Width - 10
    checkModel05.Height = checkModel00.Height
    checkModel05.Name = "CheckBox05"
    checkModel05.State = \
        {False: 0, True: 1}[config.getboolean("doc", "empty row after group title")]
    checkModel05.Label = "Добавить пустую строку после заголовка группы"
    checkModel05.HelpText = """\
Если отмечено, то между заголовком
и первым компонентом группы будет
вставлена одна пустая строка."""
    pageModel0.insertByName("CheckBox05", checkModel05)

    editControlModel04 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlNumericFieldModel"
    )
    editControlModel04.Width = 50
    editControlModel04.Height = editControlHeight
    editControlModel04.PositionX = tabsModel.Width - editControlModel04.Width - 3
    editControlModel04.PositionY = checkModel05.PositionY + checkModel05.Height
    editControlModel04.Name = "EditControl04"
    editControlModel04.Value = config.getint("doc", "pages rev table")
    editControlModel04.ValueMin = 0
    editControlModel04.ValueMax = 99
    editControlModel04.ValueStep = 1
    editControlModel04.Spin = True
    editControlModel04.DecimalAccuracy = 0
    pageModel0.insertByName("EditControl04", editControlModel04)

    checkModel06 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel06.PositionX = 5
    checkModel06.PositionY = editControlModel04.PositionY
    checkModel06.Width = editControlModel04.PositionX - 5
    checkModel06.Height = editControlHeight
    checkModel06.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    checkModel06.Name = "CheckBox06"
    checkModel06.State = \
        {False: 0, True: 1}[config.getboolean("doc", "append rev table")]
    checkModel06.Label = "Добавить лист регистрации изменений, если количество листов больше:"
    checkModel06.HelpText = """\
Если отмечено и при автоматическом
построения таблицы количество листов
документа превысит указанное число,
то в конец документа будет добавлен
лист регистрации изменений."""
    pageModel0.insertByName("CheckBox06", checkModel06)

    checkModel07 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel07.PositionX = 5
    checkModel07.PositionY = editControlModel04.PositionY + editControlModel04.Height
    checkModel07.Width = tabsModel.Width - 10
    checkModel07.Height = checkModel00.Height
    checkModel07.Name = "CheckBox07"
    checkModel07.State = \
        {False: 0, True: 1}[config.getboolean("doc", "prohibit titles at bottom")]
    checkModel07.Label = "Запретить заголовки групп внизу страницы"
    checkModel07.HelpText = """\
Если отмечено, то заголовки групп,
находящиеся внизу страницы без единого
элемента, будут перемещены на следующую
страницу."""
    pageModel0.insertByName("CheckBox07", checkModel07)

    checkModel08 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel08.PositionX = 5
    checkModel08.PositionY = checkModel07.PositionY + checkModel07.Height
    checkModel08.Width = tabsModel.Width - 10
    checkModel08.Height = checkModel00.Height
    checkModel08.Name = "CheckBox08"
    checkModel08.State = \
        {False: 0, True: 1}[config.getboolean("doc", "prohibit empty rows at top")]
    checkModel08.Label = "Запретить пустые строки вверху страницы"
    checkModel08.HelpText = """\
Если отмечено, то пустые строки
вверху страницы будут удалены."""
    pageModel0.insertByName("CheckBox08", checkModel08)

    checkModel011 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel011.PositionX = 5
    checkModel011.PositionY = checkModel08.PositionY + checkModel08.Height
    checkModel011.Width = tabsModel.Width - 10
    checkModel011.Height = checkModel00.Height
    checkModel011.Name = "CheckBox011"
    checkModel011.State = \
        {False: 0, True: 1}[config.getboolean("doc", "process repeated values")]
    checkModel011.Label = "Обработать повторяющиеся значения в графах"
    checkModel011.HelpText = """\
Если отмечено, при первом повторении
значения в графе, оно будет заменено
фразой "То же", а далее кавычками."""
    pageModel0.insertByName("CheckBox011", checkModel011)

    checkModel012 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel012.PositionX = 5
    checkModel012.PositionY = checkModel011.PositionY + checkModel011.Height
    checkModel012.Width = tabsModel.Width - 10
    checkModel012.Height = checkModel00.Height
    checkModel012.Name = "CheckBox012"
    checkModel012.State = \
        {False: 0, True: 1}[config.getboolean("doc", "footprint only")]
    checkModel012.Label = "\"Посад.место\" без наименования библиотеки"
    checkModel012.HelpText = """\
Если отмечено, то посадочное место
будет указано без наименования библиотеки."""
    pageModel0.insertByName("CheckBox012", checkModel012)

    checkModel013 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel013.PositionX = 5
    checkModel013.PositionY = checkModel012.PositionY + checkModel012.Height
    checkModel013.Width = tabsModel.Width - 10
    checkModel013.Height = checkModel00.Height
    checkModel013.Name = "CheckBox013"
    checkModel013.State = \
        {False: 0, True: 1}[config.getboolean("doc", "split row by \\n")]
    checkModel013.Label = "Обрабатывать \"\\n\" как переход на новую строку"
    checkModel013.HelpText = """\
Если отмечено, то комбинация символов
"\\n" будет обрабатываться как переход
на следующую строку таблицы."""
    pageModel0.insertByName("CheckBox013", checkModel013)

    # ------------------------------------------------------------------------
    # Fields Tab Model
    # ------------------------------------------------------------------------

    pageModel1 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page1", pageModel1)
    pageModel1.Title = " Поля "

    labelModel10 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel10.PositionX = 0
    labelModel10.PositionY = 0
    labelModel10.Width = 100
    labelModel10.Height = editControlHeight
    labelModel10.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel10.Name = "Label10"
    labelModel10.Label = "Тип:"
    labelModel10.HelpText = """\
Значение поля с указанным именем
будет использовано для обозначения
типа компонента, например,
"Резистор (Резисторы)"."""
    pageModel1.insertByName("Label10", labelModel10)

    editControlModel10 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel10.Width = tabsModel.Width - labelModel10.Width - 3
    editControlModel10.Height = labelModel10.Height
    editControlModel10.PositionX = labelModel10.Width
    editControlModel10.PositionY = labelModel10.PositionY
    editControlModel10.Name = "EditControl10"
    editControlModel10.Text = config.get("fields", "type")
    pageModel1.insertByName("EditControl10", editControlModel10)

    labelModel11 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel11.PositionX = 0
    labelModel11.PositionY = labelModel10.Height
    labelModel11.Width = labelModel10.Width
    labelModel11.Height = labelModel10.Height
    labelModel11.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel11.Name = "Label11"
    labelModel11.Label = "Наименование:"
    labelModel11.HelpText = """\
Значение поля с указанным именем
будет помещено в графу "Наименование"."""
    pageModel1.insertByName("Label11", labelModel11)

    editControlModel11 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel11.Width = tabsModel.Width - labelModel11.Width - 3
    editControlModel11.Height = labelModel11.Height
    editControlModel11.PositionX = labelModel11.Width
    editControlModel11.PositionY = labelModel11.PositionY
    editControlModel11.Name = "EditControl11"
    editControlModel11.Text = config.get("fields", "name")
    pageModel1.insertByName("EditControl11", editControlModel11)

    labelModel12 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel12.PositionX = 0
    labelModel12.PositionY = labelModel10.Height * 2
    labelModel12.Width = labelModel10.Width
    labelModel12.Height = labelModel10.Height
    labelModel12.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel12.Name = "Label12"
    labelModel12.Label = "Документ на поставку:"
    labelModel12.HelpText = """\
Значение поля с указанным именем
будет помещено в графу
"Обозначение документа на поставку"."""
    pageModel1.insertByName("Label12", labelModel12)

    editControlModel12 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel12.Width = tabsModel.Width - labelModel12.Width - 3
    editControlModel12.Height = labelModel12.Height
    editControlModel12.PositionX = labelModel12.Width
    editControlModel12.PositionY = labelModel12.PositionY
    editControlModel12.Name = "EditControl12"
    editControlModel12.Text = config.get("fields", "doc")
    pageModel1.insertByName("EditControl12", editControlModel12)

    labelModel17 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel17.PositionX = 0
    labelModel17.PositionY = labelModel10.Height * 3
    labelModel17.Width = labelModel10.Width
    labelModel17.Height = labelModel10.Height
    labelModel17.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel17.Name = "Label17"
    labelModel17.Label = "Поставщик:"
    labelModel17.HelpText = """\
Значение поля с указанным именем
будет помещено в графу "Поставщик"."""
    pageModel1.insertByName("Label17", labelModel17)

    editControlModel17 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel17.Width = tabsModel.Width - labelModel17.Width - 3
    editControlModel17.Height = labelModel17.Height
    editControlModel17.PositionX = labelModel17.Width
    editControlModel17.PositionY = labelModel17.PositionY
    editControlModel17.Name = "EditControl17"
    editControlModel17.Text = config.get("fields", "dealer")
    pageModel1.insertByName("EditControl17", editControlModel17)

    labelModel13 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel13.PositionX = 0
    labelModel13.PositionY = labelModel10.Height * 4
    labelModel13.Width = labelModel10.Width
    labelModel13.Height = labelModel10.Height
    labelModel13.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel13.Name = "Label13"
    labelModel13.Label = "Примечание:"
    labelModel13.HelpText = """\
Значение поля с указанным именем
будет помещено в графу "Примечание"."""
    pageModel1.insertByName("Label13", labelModel13)

    editControlModel13 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel13.Width = tabsModel.Width - labelModel13.Width - 3
    editControlModel13.Height = labelModel13.Height
    editControlModel13.PositionX = labelModel13.Width
    editControlModel13.PositionY = labelModel13.PositionY
    editControlModel13.Name = "EditControl13"
    editControlModel13.Text = config.get("fields", "comment")
    pageModel1.insertByName("EditControl13", editControlModel13)

    labelModel15 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel15.PositionX = 0
    labelModel15.PositionY = labelModel10.Height * 5
    labelModel15.Width = labelModel10.Width
    labelModel15.Height = labelModel10.Height
    labelModel15.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel15.Name = "Label15"
    labelModel15.Label = "Исключить:"
    labelModel15.HelpText = """\
Если компонент содержит
поле с указанным именем,
то он будет исключён из
ведомости."""
    pageModel1.insertByName("Label15", labelModel15)

    editControlModel15 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel15.Width = tabsModel.Width - labelModel15.Width - 3
    editControlModel15.Height = labelModel15.Height
    editControlModel15.PositionX = labelModel15.Width
    editControlModel15.PositionY = labelModel15.PositionY
    editControlModel15.Name = "EditControl15"
    editControlModel15.Text = config.get("fields", "excluded")
    pageModel1.insertByName("EditControl15", editControlModel15)

    buttonModel10 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModel10.Width = tabsModel.Width - 7
    buttonModel10.Height = 16
    buttonModel10.PositionX = 2
    buttonModel10.PositionY = dialogModel.Height - 100
    buttonModel10.Name = "Button10"
    buttonModel10.Label = "Установить значения по умолчанию"
    pageModel1.insertByName("Button10", buttonModel10)

    buttonModel11 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModel11.Width = tabsModel.Width - 7
    buttonModel11.Height = 16
    buttonModel11.PositionX = 2
    buttonModel11.PositionY = buttonModel10.PositionY + buttonModel10.Height + 2
    buttonModel11.Name = "Button11"
    buttonModel11.Label = "Установить значения, совместимые с kicadbom2spec"
    pageModel1.insertByName("Button11", buttonModel11)

    checkModel10 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel10.Width = tabsModel.Width - 7
    checkModel10.Height = 15
    checkModel10.PositionX = 2
    checkModel10.PositionY = buttonModel11.PositionY + buttonModel11.Height + 2
    checkModel10.Name = "CheckBox10"
    checkModel10.State = \
        {False: 0, True: 1}[config.getboolean("settings", "compatibility mode")]
    checkModel10.Label = "Режим совместимости с kicadbom2spec"
    checkModel10.HelpText = """\
Если отмечено, то при формировании
ведомости из файла настроек
приложения kicadbom2spec будут
использованы данные о разделителях
и словарь наименований групп."""
    pageModel1.insertByName("CheckBox10", checkModel10)

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
"Ведомость покупных изделий".
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
указан код "ВП".
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

    checkModel23 = pageModel2.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel23.PositionX = checkModel20.PositionX
    checkModel23.PositionY = checkModel20.PositionY + checkModel20.Height * 3
    checkModel23.Width = checkModel20.Width
    checkModel23.Height = checkModel20.Height
    checkModel23.Name = "CheckBox23"
    checkModel23.State = {False: 0, True: 1}[
        config.getboolean("stamp", "doc type is file name")
    ]
    checkModel23.Label = "Использовать имя файла в качестве типа документа"
    checkModel23.HelpText = """\
Если отмечено, и активен параметр
"Преобразовать наименование документа",
то в качестве типа документа будет
указано имя файла."""
    pageModel2.insertByName("CheckBox23", checkModel23)

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

    Button10 = dialog.getControl("Tabs").getControl("Page1").getControl("Button10")
    Button11 = dialog.getControl("Tabs").getControl("Page1").getControl("Button11")
    Button10.addActionListener(Button10ActionListener(dialog))
    Button11.addActionListener(Button10ActionListener(dialog))

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
        page1 = self.dialog.getControl("Tabs").getControl("Page1")
        page2 = self.dialog.getControl("Tabs").getControl("Page2")
        page3 = self.dialog.getControl("Tabs").getControl("Page3")

        # --------------------------------------------------------------------
        # Оптимальный вид
        # --------------------------------------------------------------------
        config.set("settings", "set view options",
            {0: "no", 1: "yes"}[self.dialog.getControl("CheckBoxSet").State]
        )

        # --------------------------------------------------------------------
        # Ведомость
        # --------------------------------------------------------------------

        config.set("doc", "source",
            page0.getControl("EditControl00").Text
        )
        config.set("doc", "empty rows between diff type",
            str(int(page0.getControl("EditControl02").Value))
        )
        config.set("doc", "extreme width factor",
            str(int(page0.getControl("EditControl03").Value))
        )
        config.set("doc", "add units",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox00").State]
        )
        config.set("doc", "space before units",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox01").State]
        )
        config.set("doc", "separate group for each doc",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox02").State]
        )
        config.set("doc", "every group has title",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox04").State]
        )
        config.set("doc", "only components have position numbers",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox010").State]
        )
        config.set("doc", "reserve position numbers",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox09").State]
        )
        config.set("doc", "empty row after group title",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox05").State]
        )
        config.set("doc", "append rev table",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox06").State]
        )
        config.set("doc", "pages rev table",
            str(int(page0.getControl("EditControl04").Value))
        )
        config.set("doc", "prohibit titles at bottom",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox07").State]
        )
        config.set("doc", "prohibit empty rows at top",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox08").State]
        )
        config.set("doc", "process repeated values",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox011").State]
        )
        config.set("doc", "footprint only",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox012").State]
        )
        config.set("doc", "split row by \\n",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox013").State]
        )

        # --------------------------------------------------------------------
        # Поля
        # --------------------------------------------------------------------

        config.set("fields", "type",
            page1.getControl("EditControl10").Text
        )
        config.set("fields", "name",
            page1.getControl("EditControl11").Text
        )
        config.set("fields", "doc",
            page1.getControl("EditControl12").Text
        )
        config.set("fields", "dealer",
            page1.getControl("EditControl17").Text
        )
        config.set("fields", "comment",
            page1.getControl("EditControl13").Text
        )
        config.set("fields", "excluded",
            page1.getControl("EditControl15").Text
        )
        config.set("settings", "compatibility mode",
            {0: "no", 1: "yes"}[page1.getControl("CheckBox10").State]
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
        config.set("stamp", "doc type is file name",
            {0: "no", 1: "yes"}[page2.getControl("CheckBox23").State]
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


class Button10ActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        page1 = self.dialog.getControl("Tabs").getControl("Page1")
        defaultValues = (
            ("EditControl10", "Тип"),
            ("EditControl11", "Наименование"),
            ("EditControl12", "Документ"),
            ("EditControl13", "Примечание"),
            ("EditControl15", ""),
            ("EditControl17", ""),
        )
        if event.Source.Model.Name == "Button11":
            separators = {
                "Марка":["", "-"],
                "Значение":["", ""],
                "Класс точности":["", ""],
                "Тип":["-", ""],
            }
            # KB2S - kicadbom2spec
            settingsKB2S = config.loadFromKicadbom2spec()
            if settingsKB2S is None:
                common.showMessage(
                    "Не удалось найти или загрузить файл настроек kicadbom2spec.\n" \
                    "Будут использованы значения по умолчанию.",
                    "Режим совместимости"
                )
            else:
                for item in separators:
                    if settingsKB2S.has_option("prefixes", item.lower()):
                        separators[item][0] = settingsKB2S.get("prefixes", item.lower())[1:-1]
                    if settingsKB2S.has_option("suffixes", item.lower()):
                        separators[item][1] = settingsKB2S.get("suffixes", item.lower())[1:-1]
            name = ""
            for item in separators:
                if any(separators[item]):
                    name += "${{{}|{}|{}}}".format(
                        separators[item][0],
                        item,
                        separators[item][1]
                    )
                else:
                    name += "${{{}}}".format(item)
            defaultValues = (
                ("EditControl10", "Группа"),
                ("EditControl11", name),
                ("EditControl12", "Стандарт"),
                ("EditControl13", "Примечание"),
                ("EditControl15", "Исключён из ПЭ"),
                ("EditControl17", ""),
            )
            page1.getControl("CheckBox10").State = 1
        else:
            page1.getControl("CheckBox10").State = 0
        for control, value in defaultValues:
            page1.getControl(control).Text = value
