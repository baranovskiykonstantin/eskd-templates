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
    dialogModel.Height = 320
    dialogModel.PositionX = 0
    dialogModel.PositionY = 0
    dialogModel.Title = "Параметры перечня элементов"

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
    # Index Tab Model
    # ------------------------------------------------------------------------

    pageModel0 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page0", pageModel0)
    pageModel0.Title = " Перечень элементов "

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
В случае с KiCad, источником данных
о схеме является файл списка цепей.
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
    editControlModel00.Text = config.get("index", "source")
    pageModel0.insertByName("EditControl00", editControlModel00)

    editControlModel01 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlNumericFieldModel"
    )
    editControlModel01.Width = 50
    editControlModel01.Height = editControlHeight
    editControlModel01.PositionX = 0
    editControlModel01.PositionY = editControlModel00.PositionY + editControlModel00.Height
    editControlModel01.Name = "EditControl01"
    editControlModel01.Value = config.getint("index", "empty rows between diff ref")
    editControlModel01.ValueMin = 0
    editControlModel01.ValueMax = 99
    editControlModel01.ValueStep = 1
    editControlModel01.Spin = True
    editControlModel01.DecimalAccuracy = 0
    pageModel0.insertByName("EditControl01", editControlModel01)

    labelModel01 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel01.PositionX = editControlModel01.PositionX + editControlModel01.Width
    labelModel01.PositionY = editControlModel01.PositionY
    labelModel01.Width = tabsModel.Width - labelModel01.PositionX
    labelModel01.Height = editControlModel01.Height
    labelModel01.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel01.Name = "Label01"
    labelModel01.Label = " пустых строк между компонентами с разными обозначениями"
    labelModel01.HelpText = """\
Указанное количество пустых строк будет
вставлено между компонентами, которые
отличаются буквенной частью обозначения.
Но, если компоненты имеют одинаковый тип
и установлен параметр
"Объединить однотипные группы",
то пустые строки между ними вставлены
не будут."""
    pageModel0.insertByName("Label01", labelModel01)

    editControlModel02 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlNumericFieldModel"
    )
    editControlModel02.Width = editControlModel01.Width
    editControlModel02.Height = editControlModel01.Height
    editControlModel02.PositionX = editControlModel01.PositionX
    editControlModel02.PositionY = editControlModel01.PositionY + editControlModel01.Height
    editControlModel02.Name = "EditControl02"
    editControlModel02.Value = config.getint("index", "empty rows between diff type")
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
вставлено между компонентами, у которых
совпадает буквенная часть обозначения,
но отличается тип."""
    pageModel0.insertByName("Label02", labelModel02)

    editControlModel03 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlNumericFieldModel"
    )
    editControlModel03.Width = editControlModel02.Width
    editControlModel03.Height = editControlModel02.Height
    editControlModel03.PositionX = editControlModel02.PositionX
    editControlModel03.PositionY = editControlModel02.PositionY + editControlModel02.Height
    editControlModel03.Name = "EditControl03"
    editControlModel03.Value = config.getint("index", "extreme width factor")
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

    labelModel04 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel04.PositionX = 0
    labelModel04.PositionY = labelModel03.PositionY + labelModel03.Height
    labelModel04.Width = 150
    labelModel04.Height = 16
    labelModel04.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel04.Name = "Label04"
    labelModel04.Label = "Разделитель диапазона обозначений:"
    pageModel0.insertByName("Label04", labelModel04)

    radioButtonModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlRadioButtonModel"
    )
    radioButtonModel00.PositionX = 150
    radioButtonModel00.PositionY = labelModel04.PositionY
    radioButtonModel00.Width = 75
    radioButtonModel00.Height = labelModel04.Height
    radioButtonModel00.Name = "RadioButton00"
    radioButtonModel00.Label = "-"
    radioButtonModel00.State = 1 if config.get("index", "ref separator") == '-' else 0
    pageModel0.insertByName("RadioButton00", radioButtonModel00)

    radioButtonModel01 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlRadioButtonModel"
    )
    radioButtonModel01.PositionX = radioButtonModel00.PositionX + radioButtonModel00.Width
    radioButtonModel01.PositionY = radioButtonModel00.PositionY
    radioButtonModel01.Width = radioButtonModel00.Width
    radioButtonModel01.Height = radioButtonModel00.Height
    radioButtonModel01.Name = "RadioButton01"
    radioButtonModel01.Label = "…"
    radioButtonModel01.State = 1 if config.get("index", "ref separator") == '…' else 0
    pageModel0.insertByName("RadioButton01", radioButtonModel01)

    checkModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel00.PositionX = 5
    checkModel00.PositionY = labelModel04.PositionY + labelModel04.Height + 5
    checkModel00.Width = tabsModel.Width - 10
    checkModel00.Height = 15
    checkModel00.Name = "CheckBox00"
    checkModel00.State = {False: 0, True: 1}[
        config.getboolean("index", "add units")
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
        {False: 0, True: 1}[config.getboolean("index", "space before units")]
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
        {False: 0, True: 1}[config.getboolean("index", "concatenate same name groups")]
    checkModel02.Label = "Объединить однотипные группы"
    checkModel02.HelpText = """\
По умолчанию, группой считается
совокупность компонентов с одинаковой
буквенной частью обозначения.
Если отмечено, то идущие подряд компоненты
с одинаковым типом будут объединены в одну
группу, даже если буквенная часть их
обозначений отличается."""
    pageModel0.insertByName("CheckBox02", checkModel02)

    checkModel03 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel03.PositionX = 5
    checkModel03.PositionY = checkModel02.PositionY + checkModel00.Height
    checkModel03.Width = tabsModel.Width - 10
    checkModel03.Height = checkModel00.Height
    checkModel03.Name = "CheckBox03"
    checkModel03.State = \
        {False: 0, True: 1}[config.getboolean("index", "title with doc")]
    checkModel03.Label = "Указать документ в заголовке группы"
    checkModel03.HelpText = """\
По умолчанию, в качестве заголовка группы
компонентов выступает тип в множественном
числе.
Если отмечено, то в заголовке, после типа,
будет указан документ (ГОСТ, ТУ, ...).
Если в группе компоненты имеют разные
документы, то перед каждым документом в
заголовке будет указана часть наименования,
необходимая для идентификации
соответствующих компонентов."""
    pageModel0.insertByName("CheckBox03", checkModel03)

    checkModel04 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel04.PositionX = 5
    checkModel04.PositionY = checkModel03.PositionY + checkModel00.Height
    checkModel04.Width = tabsModel.Width - 10
    checkModel04.Height = checkModel00.Height
    checkModel04.Name = "CheckBox04"
    checkModel04.State = \
        {False: 0, True: 1}[config.getboolean("index", "every group has title")]
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

    checkModel05 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel05.PositionX = 5
    checkModel05.PositionY = checkModel04.PositionY + checkModel00.Height
    checkModel05.Width = tabsModel.Width - 10
    checkModel05.Height = checkModel00.Height
    checkModel05.Name = "CheckBox05"
    checkModel05.State = \
        {False: 0, True: 1}[config.getboolean("index", "empty row after group title")]
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
    editControlModel04.Value = config.getint("index", "pages rev table")
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
        {False: 0, True: 1}[config.getboolean("index", "append rev table")]
    checkModel06.Label = "Добавить лист регистрации изменений, если количество листов больше:"
    checkModel06.HelpText = """\
Если отмечено и количество листов
документа превышает указанное число,
то в конец документа будет добавлен
лист регистрации изменений.
Если в процессе редактирования
количество листов станет меньше
указанного значения, то лист
регистрации изменения будет удалён."""
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
        {False: 0, True: 1}[config.getboolean("index", "prohibit titles at bottom")]
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
        {False: 0, True: 1}[config.getboolean("index", "prohibit empty rows at top")]
    checkModel08.Label = "Запретить пустые строки вверху страницы"
    checkModel08.HelpText = """\
Если отмечено, то пустые строки
вверху страницы будут удалены."""
    pageModel0.insertByName("CheckBox08", checkModel08)

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
    labelModel12.Label = "Документ:"
    labelModel12.HelpText = """\
Значение поля с указанным именем
будет добавлено к "Наименованию",
указывая на ГОСТ, ТУ или прочий
документ."""
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

    labelModel13 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel13.PositionX = 0
    labelModel13.PositionY = labelModel10.Height * 3
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

    labelModel14 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel14.PositionX = 0
    labelModel14.PositionY = labelModel10.Height * 4
    labelModel14.Width = labelModel10.Width
    labelModel14.Height = labelModel10.Height
    labelModel14.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel14.Name = "Label14"
    labelModel14.Label = "Подбирают при регулировании:"
    labelModel14.HelpText = """\
Если компонент содержит
поле с указанным именем,
то возле его обозначения
в перечне, будет указан
символ "*"."""
    pageModel1.insertByName("Label14", labelModel14)

    editControlModel14 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel14.Width = tabsModel.Width - labelModel14.Width - 3
    editControlModel14.Height = labelModel14.Height
    editControlModel14.PositionX = labelModel14.Width
    editControlModel14.PositionY = labelModel14.PositionY
    editControlModel14.Name = "EditControl14"
    editControlModel14.Text = config.get("fields", "adjustable")
    pageModel1.insertByName("EditControl14", editControlModel14)

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
перечня элементов."""
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
    buttonModel10.PositionY = 196
    buttonModel10.Name = "Button10"
    buttonModel10.Label = "Установить значения по умолчанию"
    pageModel1.insertByName("Button10", buttonModel10)

    buttonModel11 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModel11.Width = tabsModel.Width - 7
    buttonModel11.Height = 16
    buttonModel11.PositionX = 2
    buttonModel11.PositionY = 214
    buttonModel11.Name = "Button11"
    buttonModel11.Label = "Установить значения, совместимые с kicadbom2spec"
    pageModel1.insertByName("Button11", buttonModel11)

    checkModel10 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel10.Width = tabsModel.Width - 7
    checkModel10.Height = 15
    checkModel10.PositionX = 2
    checkModel10.PositionY = 232
    checkModel10.Name = "CheckBox10"
    checkModel10.State = \
        {False: 0, True: 1}[config.getboolean("settings", "compatibility mode")]
    checkModel10.Label = "Режим совместимости с kicadbom2spec"
    checkModel10.HelpText = """\
Если отмечено, то при формировании
перечня элементов из файла настроек
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
"Перечень элементов".
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
Если отмечено, к типу схемы в
обозначении документа будет
добавлен префикс "П" (перечень).
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


class ButtonOKActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        page0 = self.dialog.getControl("Tabs").getControl("Page0")
        page1 = self.dialog.getControl("Tabs").getControl("Page1")
        page2 = self.dialog.getControl("Tabs").getControl("Page2")

        # --------------------------------------------------------------------
        # Оптимальный вид
        # --------------------------------------------------------------------
        config.set("settings", "set view options",
            {0: "no", 1: "yes"}[self.dialog.getControl("CheckBoxSet").State]
        )

        # --------------------------------------------------------------------
        # Перечень элементов
        # --------------------------------------------------------------------

        config.set("index", "source",
            page0.getControl("EditControl00").Text
        )
        config.set("index", "empty rows between diff ref",
            str(int(page0.getControl("EditControl01").Value))
        )
        config.set("index", "empty rows between diff type",
            str(int(page0.getControl("EditControl02").Value))
        )
        config.set("index", "extreme width factor",
            str(int(page0.getControl("EditControl03").Value))
        )
        if page0.getControl("RadioButton00").State:
            config.set("index", "ref separator", "-")
        if page0.getControl("RadioButton01").State:
            config.set("index", "ref separator", "…")
        config.set("index", "add units",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox00").State]
        )
        config.set("index", "space before units",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox01").State]
        )
        config.set("index", "concatenate same name groups",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox02").State]
        )
        config.set("index", "title with doc",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox03").State]
        )
        config.set("index", "every group has title",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox04").State]
        )
        config.set("index", "empty row after group title",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox05").State]
        )
        config.set("index", "append rev table",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox06").State]
        )
        config.set("index", "pages rev table",
            str(int(page0.getControl("EditControl04").Value))
        )
        config.set("index", "prohibit titles at bottom",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox07").State]
        )
        config.set("index", "prohibit empty rows at top",
            {0: "no", 1: "yes"}[page0.getControl("CheckBox08").State]
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
        config.set("fields", "comment",
            page1.getControl("EditControl13").Text
        )
        config.set("fields", "adjustable",
            page1.getControl("EditControl14").Text
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
            editControl.Text
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
            ("EditControl14", "Подбирают при регулировании"),
            ("EditControl15", ""),
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
            if settingsKB2S is not None:
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
                ("EditControl14", "Подбирают при регулировании"),
                ("EditControl15", "Исключён из ПЭ"),
            )
            page1.getControl("CheckBox10").State = 1
        else:
            page1.getControl("CheckBox10").State = 0
        for control, value in defaultValues:
            page1.getControl(control).Text = value
