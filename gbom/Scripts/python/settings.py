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
    if sys.platform.startswith("linux"):
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
    checkModelSet.Width = 140
    checkModelSet.Height = 16
    checkModelSet.PositionX = 5
    checkModelSet.PositionY = dialogModel.Height - checkModelSet.Height - 4
    checkModelSet.Name = "CheckBoxSet"
    checkModelSet.State = int(config.getboolean("settings", "set view options"))
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
    checkModel00.State = int(config.getboolean("doc", "add units"))
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
    checkModel01.State = int(config.getboolean("doc", "space before units"))
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
    checkModel02.State = int(config.getboolean("doc", "separate group for each doc"))
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
    checkModel04.State = int(config.getboolean("doc", "every group has title"))
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
    checkModel010.State = int(config.getboolean("doc", "only components have position numbers"))
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
    checkModel09.State = int(config.getboolean("doc", "reserve position numbers"))
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
    checkModel05.State = int(config.getboolean("doc", "empty row after group title"))
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
    checkModel06.State = int(config.getboolean("doc", "append rev table"))
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
    checkModel07.State = int(config.getboolean("doc", "prohibit titles at bottom"))
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
    checkModel08.State = int(config.getboolean("doc", "prohibit empty rows at top"))
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
    checkModel011.State = int(config.getboolean("doc", "process repeated values"))
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
    checkModel012.State = int(config.getboolean("doc", "footprint only"))
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
    checkModel013.State = int(config.getboolean("doc", "split row by \\n"))
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

    labelModel16 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel16.PositionX = 0
    labelModel16.PositionY = labelModel10.Height * 2
    labelModel16.Width = labelModel10.Width
    labelModel16.Height = labelModel10.Height
    labelModel16.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel16.Name = "Label16"
    labelModel16.Label = "Код ОКП:"
    labelModel16.HelpText = """\
Значение поля с указанным именем
будет помещено в графу "Код ОКП"."""
    pageModel1.insertByName("Label16", labelModel16)

    editControlModel16 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel16.Width = tabsModel.Width - labelModel16.Width - 3
    editControlModel16.Height = labelModel16.Height
    editControlModel16.PositionX = labelModel16.Width
    editControlModel16.PositionY = labelModel16.PositionY
    editControlModel16.Name = "EditControl16"
    editControlModel16.Text = config.get("fields", "code")
    pageModel1.insertByName("EditControl16", editControlModel16)

    labelModel12 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel12.PositionX = 0
    labelModel12.PositionY = labelModel10.Height * 3
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
    labelModel17.PositionY = labelModel10.Height * 4
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
    labelModel13.PositionY = labelModel10.Height * 5
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
    labelModel15.PositionY = labelModel10.Height * 6
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
    checkModel10.State = int(config.getboolean("settings", "compatibility mode"))
    checkModel10.Label = "Режим совместимости с kicadbom2spec"
    checkModel10.HelpText = """\
Если отмечено, то при формировании
ведомости из файла настроек
приложения kicadbom2spec будут
использованы данные о разделителях
и словарь наименований групп."""
    pageModel1.insertByName("CheckBox10", checkModel10)

    # ------------------------------------------------------------------------
    # Sorting Tab Model
    # ------------------------------------------------------------------------

    pageModel4 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page4", pageModel4)
    pageModel4.Title = " Сортировка "

    # group frame

    groupBoxModel400 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlGroupBoxModel"
    )
    groupBoxModel400.PositionX = 0
    groupBoxModel400.PositionY = 5
    groupBoxModel400.Width = dialogModel.Width - 4
    groupBoxModel400.Height = 15 + 3 * editControlHeight + 16
    groupBoxModel400.Label = "Сортировка групп"
    pageModel4.insertByName("GroupBox400", groupBoxModel400)

    labelModel400 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel400.PositionX = 5
    labelModel400.PositionY = groupBoxModel400.PositionY + 10
    labelModel400.Width = int((dialogModel.Width - 10) * 0.05)
    labelModel400.Height = 15
    labelModel400.Align = 1 # center
    labelModel400.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel400.Name = "Label400"
    labelModel400.Label = "№"
    pageModel4.insertByName("Label400", labelModel400)

    labelModel401 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel401.PositionX = labelModel400.PositionX + labelModel400.Width
    labelModel401.PositionY = labelModel400.PositionY
    labelModel401.Width = int((dialogModel.Width - 10) * 0.35)
    labelModel401.Height = labelModel400.Height
    labelModel401.Align = 1 # center
    labelModel401.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel401.Name = "Label401"
    labelModel401.Label = "Поле или шаблон"
    pageModel4.insertByName("Label401", labelModel401)

    labelModel402 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel402.PositionX = labelModel401.PositionX + labelModel401.Width
    labelModel402.PositionY = labelModel400.PositionY
    labelModel402.Width = int((dialogModel.Width - 10) * 0.3)
    labelModel402.Height = labelModel400.Height
    labelModel402.Align = 1 # center
    labelModel402.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel402.Name = "Label402"
    labelModel402.Label = "Порядок"
    pageModel4.insertByName("Label402", labelModel402)

    labelModel403 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel403.PositionX = labelModel402.PositionX + labelModel402.Width
    labelModel403.PositionY = labelModel400.PositionY
    labelModel403.Width = int((dialogModel.Width - 10) * 0.3) - 3
    labelModel403.Height = labelModel400.Height
    labelModel403.Align = 1 # center
    labelModel403.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel403.Name = "Label403"
    labelModel403.Label = "Содержимое"
    pageModel4.insertByName("Label403", labelModel403)

    # group sort level 1

    labelModel410 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel410.PositionX = labelModel400.PositionX
    labelModel410.PositionY = labelModel400.PositionY + labelModel400.Height
    labelModel410.Width = labelModel400.Width
    labelModel410.Height = editControlHeight
    labelModel410.Align = 1 # center
    labelModel410.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel410.Name = "Label410"
    labelModel410.NoLabel = True
    pageModel4.insertByName("Label410", labelModel410)

    editControlModel411 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel411.Width = labelModel401.Width
    editControlModel411.Height = labelModel410.Height
    editControlModel411.PositionX = labelModel401.PositionX
    editControlModel411.PositionY = labelModel410.PositionY
    editControlModel411.Name = "EditControl411"
    pageModel4.insertByName("EditControl411", editControlModel411)

    listBoxModel412 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel412.Width = labelModel402.Width
    listBoxModel412.Height = labelModel410.Height
    listBoxModel412.PositionX = labelModel402.PositionX
    listBoxModel412.PositionY = labelModel410.PositionY
    listBoxModel412.Name = "ListBox412"
    listBoxModel412.Dropdown = True
    pageModel4.insertByName("ListBox412", listBoxModel412)

    listBoxModel413 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel413.Width = labelModel403.Width
    listBoxModel413.Height = labelModel410.Height
    listBoxModel413.PositionX = labelModel403.PositionX
    listBoxModel413.PositionY = labelModel410.PositionY
    listBoxModel413.Name = "ListBox413"
    listBoxModel413.Dropdown = True
    pageModel4.insertByName("ListBox413", listBoxModel413)

    # group sort level 2

    labelModel420 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel420.PositionX = labelModel400.PositionX
    labelModel420.PositionY = labelModel410.PositionY + labelModel410.Height
    labelModel420.Width = labelModel400.Width
    labelModel420.Height = editControlHeight
    labelModel420.Align = 1 # center
    labelModel420.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel420.Name = "Label420"
    labelModel420.NoLabel = True
    pageModel4.insertByName("Label420", labelModel420)

    editControlModel421 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel421.Width = labelModel401.Width
    editControlModel421.Height = labelModel420.Height
    editControlModel421.PositionX = labelModel401.PositionX
    editControlModel421.PositionY = labelModel420.PositionY
    editControlModel421.Name = "EditControl421"
    pageModel4.insertByName("EditControl421", editControlModel421)

    listBoxModel422 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel422.Width = labelModel402.Width
    listBoxModel422.Height = labelModel420.Height
    listBoxModel422.PositionX = labelModel402.PositionX
    listBoxModel422.PositionY = labelModel420.PositionY
    listBoxModel422.Name = "ListBox422"
    listBoxModel422.Dropdown = True
    pageModel4.insertByName("ListBox422", listBoxModel422)

    listBoxModel423 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel423.Width = labelModel403.Width
    listBoxModel423.Height = labelModel420.Height
    listBoxModel423.PositionX = labelModel403.PositionX
    listBoxModel423.PositionY = labelModel420.PositionY
    listBoxModel423.Name = "ListBox423"
    listBoxModel423.Dropdown = True
    pageModel4.insertByName("ListBox423", listBoxModel423)

    # group sort level 3

    labelModel430 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel430.PositionX = labelModel400.PositionX
    labelModel430.PositionY = labelModel420.PositionY + labelModel420.Height
    labelModel430.Width = labelModel400.Width
    labelModel430.Height = editControlHeight
    labelModel430.Align = 1 # center
    labelModel430.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel430.Name = "Label430"
    labelModel430.NoLabel = True
    pageModel4.insertByName("Label430", labelModel430)

    editControlModel431 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel431.Width = labelModel401.Width
    editControlModel431.Height = labelModel430.Height
    editControlModel431.PositionX = labelModel401.PositionX
    editControlModel431.PositionY = labelModel430.PositionY
    editControlModel431.Name = "EditControl431"
    pageModel4.insertByName("EditControl431", editControlModel431)

    listBoxModel432 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel432.Width = labelModel402.Width
    listBoxModel432.Height = labelModel430.Height
    listBoxModel432.PositionX = labelModel402.PositionX
    listBoxModel432.PositionY = labelModel430.PositionY
    listBoxModel432.Name = "ListBox432"
    listBoxModel432.Dropdown = True
    pageModel4.insertByName("ListBox432", listBoxModel432)

    listBoxModel433 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel433.Width = labelModel403.Width
    listBoxModel433.Height = labelModel430.Height
    listBoxModel433.PositionX = labelModel403.PositionX
    listBoxModel433.PositionY = labelModel430.PositionY
    listBoxModel433.Name = "ListBox433"
    listBoxModel433.Dropdown = True
    pageModel4.insertByName("ListBox433", listBoxModel433)

    # comp frame

    groupBoxModel401 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlGroupBoxModel"
    )
    groupBoxModel401.PositionX = 0
    groupBoxModel401.PositionY = groupBoxModel400.PositionY + groupBoxModel400.Height + 10
    groupBoxModel401.Width = groupBoxModel400.Width
    groupBoxModel401.Height = groupBoxModel400.Height
    groupBoxModel401.Label = "Сортировка компонентов"
    pageModel4.insertByName("GroupBox401", groupBoxModel401)

    labelModel404 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel404.PositionX = labelModel400.PositionX
    labelModel404.PositionY = groupBoxModel401.PositionY + 10
    labelModel404.Width = labelModel400.Width
    labelModel404.Height = labelModel400.Height
    labelModel404.Align = 1 # center
    labelModel404.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel404.Name = "Label404"
    labelModel404.Label = "№"
    pageModel4.insertByName("Label404", labelModel404)

    labelModel405 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel405.PositionX = labelModel401.PositionX
    labelModel405.PositionY = labelModel404.PositionY
    labelModel405.Width = labelModel401.Width
    labelModel405.Height = labelModel404.Height
    labelModel405.Align = 1 # center
    labelModel405.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel405.Name = "Label405"
    labelModel405.Label = "Поле или шаблон"
    pageModel4.insertByName("Label405", labelModel405)

    labelModel406 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel406.PositionX = labelModel402.PositionX
    labelModel406.PositionY = labelModel404.PositionY
    labelModel406.Width = labelModel402.Width
    labelModel406.Height = labelModel404.Height
    labelModel406.Align = 1 # center
    labelModel406.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel406.Name = "Label406"
    labelModel406.Label = "Порядок"
    pageModel4.insertByName("Label406", labelModel406)

    labelModel407 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel407.PositionX = labelModel403.PositionX
    labelModel407.PositionY = labelModel404.PositionY
    labelModel407.Width = labelModel403.Width
    labelModel407.Height = labelModel404.Height
    labelModel407.Align = 1 # center
    labelModel407.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel407.Name = "Label407"
    labelModel407.Label = "Содержимое"
    pageModel4.insertByName("Label407", labelModel407)

    # comp sort level 1

    labelModel440 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel440.PositionX = labelModel400.PositionX
    labelModel440.PositionY = labelModel404.PositionY + labelModel404.Height
    labelModel440.Width = labelModel400.Width
    labelModel440.Height = labelModel410.Height
    labelModel440.Align = 1 # center
    labelModel440.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel440.Name = "Label440"
    labelModel440.NoLabel = True
    pageModel4.insertByName("Label440", labelModel440)

    editControlModel441 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel441.Width = labelModel405.Width
    editControlModel441.Height = labelModel440.Height
    editControlModel441.PositionX = labelModel405.PositionX
    editControlModel441.PositionY = labelModel440.PositionY
    editControlModel441.Name = "EditControl441"
    pageModel4.insertByName("EditControl441", editControlModel441)

    listBoxModel442 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel442.Width = labelModel406.Width
    listBoxModel442.Height = labelModel440.Height
    listBoxModel442.PositionX = labelModel406.PositionX
    listBoxModel442.PositionY = labelModel440.PositionY
    listBoxModel442.Name = "ListBox442"
    listBoxModel442.Dropdown = True
    pageModel4.insertByName("ListBox442", listBoxModel442)

    listBoxModel443 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel443.Width = labelModel407.Width
    listBoxModel443.Height = labelModel440.Height
    listBoxModel443.PositionX = labelModel407.PositionX
    listBoxModel443.PositionY = labelModel440.PositionY
    listBoxModel443.Name = "ListBox443"
    listBoxModel443.Dropdown = True
    pageModel4.insertByName("ListBox443", listBoxModel443)

    # comp sort level 2

    labelModel450 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel450.PositionX = labelModel404.PositionX
    labelModel450.PositionY = labelModel440.PositionY + labelModel440.Height
    labelModel450.Width = labelModel404.Width
    labelModel450.Height = labelModel440.Height
    labelModel450.Align = 1 # center
    labelModel450.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel450.Name = "Label450"
    labelModel450.NoLabel = True
    pageModel4.insertByName("Label450", labelModel450)

    editControlModel451 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel451.Width = labelModel405.Width
    editControlModel451.Height = labelModel450.Height
    editControlModel451.PositionX = labelModel405.PositionX
    editControlModel451.PositionY = labelModel450.PositionY
    editControlModel451.Name = "EditControl451"
    pageModel4.insertByName("EditControl451", editControlModel451)

    listBoxModel452 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel452.Width = labelModel406.Width
    listBoxModel452.Height = labelModel450.Height
    listBoxModel452.PositionX = labelModel406.PositionX
    listBoxModel452.PositionY = labelModel450.PositionY
    listBoxModel452.Name = "ListBox452"
    listBoxModel452.Dropdown = True
    pageModel4.insertByName("ListBox452", listBoxModel452)

    listBoxModel453 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel453.Width = labelModel407.Width
    listBoxModel453.Height = labelModel450.Height
    listBoxModel453.PositionX = labelModel407.PositionX
    listBoxModel453.PositionY = labelModel450.PositionY
    listBoxModel453.Name = "ListBox453"
    listBoxModel453.Dropdown = True
    pageModel4.insertByName("ListBox453", listBoxModel453)

    # comp sort level 3

    labelModel460 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel460.PositionX = labelModel404.PositionX
    labelModel460.PositionY = labelModel450.PositionY + labelModel450.Height
    labelModel460.Width = labelModel404.Width
    labelModel460.Height = labelModel440.Height
    labelModel460.Align = 1 # center
    labelModel460.VerticalAlign = uno.Enum(
        "com.sun.star.style.VerticalAlignment",
        "MIDDLE"
    )
    labelModel460.Name = "Label460"
    labelModel460.NoLabel = True
    pageModel4.insertByName("Label460", labelModel460)

    editControlModel461 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel461.Width = labelModel405.Width
    editControlModel461.Height = labelModel460.Height
    editControlModel461.PositionX = labelModel405.PositionX
    editControlModel461.PositionY = labelModel460.PositionY
    editControlModel461.Name = "EditControl461"
    pageModel4.insertByName("EditControl461", editControlModel461)

    listBoxModel462 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel462.Width = labelModel406.Width
    listBoxModel462.Height = labelModel460.Height
    listBoxModel462.PositionX = labelModel406.PositionX
    listBoxModel462.PositionY = labelModel460.PositionY
    listBoxModel462.Name = "ListBox462"
    listBoxModel462.Dropdown = True
    pageModel4.insertByName("ListBox462", listBoxModel462)

    listBoxModel463 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel"
    )
    listBoxModel463.Width = labelModel407.Width
    listBoxModel463.Height = labelModel460.Height
    listBoxModel463.PositionX = labelModel407.PositionX
    listBoxModel463.PositionY = labelModel460.PositionY
    listBoxModel463.Name = "ListBox463"
    listBoxModel463.Dropdown = True
    pageModel4.insertByName("ListBox463", listBoxModel463)

    buttonModel400 = pageModel4.createInstance(
        "com.sun.star.awt.UnoControlButtonModel"
    )
    buttonModel400.Width = tabsModel.Width - 7
    buttonModel400.Height = 16
    buttonModel400.PositionX = 2
    buttonModel400.PositionY = dialogModel.Height - 67
    buttonModel400.Name = "Button400"
    buttonModel400.Label = "Установить значения по умолчанию"
    pageModel4.insertByName("Button400", buttonModel400)

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
    checkModel20.State = int(config.getboolean("stamp", "convert doc title"))
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
    checkModel21.State = int(config.getboolean("stamp", "convert doc id"))
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
    checkModel22.State = int(config.getboolean("stamp", "fill first usage"))
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
    checkModel23.State = int(config.getboolean("stamp", "place doc id to table title"))
    checkModel23.Label = "Поместить децимальный номер в заголовок графы \"Кол. на исполнение\""
    checkModel23.HelpText = """\
Если отмечено, при заполнении
основной надписи в заголовке
графы "Кол. на исполнение"
будет указано обозначение
документа.
В противном случае, графа
останется без изменений."""
    pageModel2.insertByName("CheckBox23", checkModel23)

    checkModel24 = pageModel2.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel24.PositionX = checkModel20.PositionX
    checkModel24.PositionY = checkModel20.PositionY + checkModel20.Height * 4
    checkModel24.Width = checkModel20.Width
    checkModel24.Height = checkModel20.Height
    checkModel24.Name = "CheckBox24"
    checkModel24.State = int(config.getboolean("stamp", "doc type is file name"))
    checkModel24.Label = "Использовать имя файла в качестве типа документа"
    checkModel24.HelpText = """\
Если отмечено, и активен параметр
"Преобразовать наименование документа",
то в качестве типа документа будет
указано имя файла."""
    pageModel2.insertByName("CheckBox24", checkModel24)

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
    Button400 = dialog.getControl("Tabs").getControl("Page4").getControl("Button400")
    Button400.addActionListener(Button400ActionListener(dialog))

    # ------------------------------------------------------------------------

    toolkit = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.Toolkit",
        context
    )
    dialog.createPeer(toolkit, None)
    dialog.Visible = True

    # Initialize sort controls in loop
    # Some properties cannot be set while dialog is not shown
    page4 = dialog.getControl("Tabs").getControl("Page4")
    for sortLevel in "123":
        # group
        page4.getControl("Label4{}0".format(sortLevel)).Text = sortLevel
        page4.getControl("EditControl4{}1".format(sortLevel)).Text = \
            config.get("group sort fields", sortLevel)
        page4.getControl("ListBox4{}2".format(sortLevel)).addItems(
            ("По возрастанию", "По убыванию"),
            0
        )
        page4.getControl("ListBox4{}2".format(sortLevel)).selectItem(
            config.get("group sort order", sortLevel),
            True
        )
        page4.getControl("ListBox4{}3".format(sortLevel)).addItems(
            ("Текст", "Число", "Текст+Число"),
            0
        )
        page4.getControl("ListBox4{}3".format(sortLevel)).selectItem(
            config.get("group sort data", sortLevel),
            True
        )
        # comp
        page4.getControl("Label4{}0".format(int(sortLevel) + 3)).Text = sortLevel
        page4.getControl("EditControl4{}1".format(int(sortLevel) + 3)).Text = \
            config.get("comp sort fields", sortLevel)
        page4.getControl("ListBox4{}2".format(int(sortLevel) + 3)).addItems(
            ("По возрастанию", "По убыванию"),
            0
        )
        page4.getControl("ListBox4{}2".format(int(sortLevel) + 3)).selectItem(
            config.get("comp sort order", sortLevel),
            True
        )
        page4.getControl("ListBox4{}3".format(int(sortLevel) + 3)).addItems(
            ("Текст", "Число", "Текст+Число"),
            0
        )
        page4.getControl("ListBox4{}3".format(int(sortLevel) + 3)).selectItem(
            config.get("comp sort data", sortLevel),
            True
        )

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
        page4 = self.dialog.getControl("Tabs").getControl("Page4")

        # --------------------------------------------------------------------
        # Оптимальный вид
        # --------------------------------------------------------------------
        config.setboolean("settings", "set view options",
            self.dialog.getControl("CheckBoxSet").State
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
        config.setboolean("doc", "add units",
            page0.getControl("CheckBox00").State
        )
        config.setboolean("doc", "space before units",
            page0.getControl("CheckBox01").State
        )
        config.setboolean("doc", "separate group for each doc",
            page0.getControl("CheckBox02").State
        )
        config.setboolean("doc", "every group has title",
            page0.getControl("CheckBox04").State
        )
        config.setboolean("doc", "only components have position numbers",
            page0.getControl("CheckBox010").State
        )
        config.setboolean("doc", "reserve position numbers",
            page0.getControl("CheckBox09").State
        )
        config.setboolean("doc", "empty row after group title",
            page0.getControl("CheckBox05").State
        )
        config.setboolean("doc", "append rev table",
            page0.getControl("CheckBox06").State
        )
        config.set("doc", "pages rev table",
            str(int(page0.getControl("EditControl04").Value))
        )
        config.setboolean("doc", "prohibit titles at bottom",
            page0.getControl("CheckBox07").State
        )
        config.setboolean("doc", "prohibit empty rows at top",
            page0.getControl("CheckBox08").State
        )
        config.setboolean("doc", "process repeated values",
            page0.getControl("CheckBox011").State
        )
        config.setboolean("doc", "footprint only",
            page0.getControl("CheckBox012").State
        )
        config.setboolean("doc", "split row by \\n",
            page0.getControl("CheckBox013").State
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
        config.set("fields", "code",
            page1.getControl("EditControl16").Text
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
        config.setboolean("settings", "compatibility mode",
            page1.getControl("CheckBox10").State
        )

        # --------------------------------------------------------------------
        # Сортировка
        # --------------------------------------------------------------------

        for sortLevel in "123":
            # group
            config.set("group sort fields", sortLevel,
                page4.getControl("EditControl4{}1".format(sortLevel)).Text
            )
            config.set("group sort order", sortLevel,
                page4.getControl("ListBox4{}2".format(sortLevel)).getSelectedItem()
            )
            config.set("group sort data", sortLevel,
                page4.getControl("ListBox4{}3".format(sortLevel)).getSelectedItem()
            )
            # comp
            config.set("comp sort fields", sortLevel,
                page4.getControl("EditControl4{}1".format(int(sortLevel) + 3)).Text
            )
            config.set("comp sort order", sortLevel,
                page4.getControl("ListBox4{}2".format(int(sortLevel) + 3)).getSelectedItem()
            )
            config.set("comp sort data", sortLevel,
                page4.getControl("ListBox4{}3".format(int(sortLevel) + 3)).getSelectedItem()
            )

        # --------------------------------------------------------------------
        # Основная надпись
        # --------------------------------------------------------------------

        config.setboolean("stamp", "convert doc title",
            page2.getControl("CheckBox20").State
        )
        config.setboolean("stamp", "convert doc id",
            page2.getControl("CheckBox21").State
        )
        config.setboolean("stamp", "fill first usage",
            page2.getControl("CheckBox22").State
        )
        config.setboolean("stamp", "place doc id to table title",
            page2.getControl("CheckBox23").State
        )
        config.setboolean("stamp", "doc type is file name",
            page2.getControl("CheckBox24").State
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
            ("EditControl16", ""),
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
                ("EditControl16", ""),
                ("EditControl17", ""),
            )
            page1.getControl("CheckBox10").State = 1
        else:
            page1.getControl("CheckBox10").State = 0
        for control, value in defaultValues:
            page1.getControl(control).Text = value


class Button400ActionListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        page4 = self.dialog.getControl("Tabs").getControl("Page4")
        # group
        page4.getControl("EditControl411").Text = "Обозначение"
        page4.getControl("ListBox412").selectItem("По возрастанию", True)
        page4.getControl("ListBox413").selectItem("Текст+Число", True)
        page4.getControl("EditControl421").Text = "Тип"
        page4.getControl("ListBox422").selectItem("По возрастанию", True)
        page4.getControl("ListBox423").selectItem("Текст", True)
        page4.getControl("EditControl431").Text = ""
        page4.getControl("ListBox432").selectItem("По возрастанию", True)
        page4.getControl("ListBox433").selectItem("Текст", True)
        # comp
        page4.getControl("EditControl441").Text = "Значение!"
        page4.getControl("ListBox442").selectItem("По возрастанию", True)
        page4.getControl("ListBox443").selectItem("Число", True)
        page4.getControl("EditControl451").Text = ""
        page4.getControl("ListBox452").selectItem("По возрастанию", True)
        page4.getControl("ListBox453").selectItem("Текст", True)
        page4.getControl("EditControl461").Text = ""
        page4.getControl("ListBox462").selectItem("По возрастанию", True)
        page4.getControl("ListBox463").selectItem("Текст", True)
