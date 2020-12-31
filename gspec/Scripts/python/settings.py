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
    dialogModel.Title = "Параметры спецификации"

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
    # Spec Tab Model
    # ------------------------------------------------------------------------

    pageModel0 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page0", pageModel0)
    pageModel0.Title = " Спецификация "

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
    radioButtonModel00.State = 1 if config.get("doc", "ref separator") == '-' else 0
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
    radioButtonModel01.State = 1 if config.get("doc", "ref separator") == '…' else 0
    pageModel0.insertByName("RadioButton01", radioButtonModel01)

    checkModel00 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel00.PositionX = 5
    checkModel00.PositionY = labelModel04.PositionY + labelModel04.Height + 5
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

    checkModel03 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel03.PositionX = 5
    checkModel03.PositionY = checkModel02.PositionY + checkModel00.Height
    checkModel03.Width = tabsModel.Width - 10
    checkModel03.Height = checkModel00.Height
    checkModel03.Name = "CheckBox03"
    checkModel03.State = int(config.getboolean("doc", "title with doc"))
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

    checkModel09 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel09.PositionX = 5
    checkModel09.PositionY = checkModel04.PositionY + checkModel00.Height
    checkModel09.Width = tabsModel.Width - 10
    checkModel09.Height = checkModel00.Height
    checkModel09.Name = "CheckBox09"
    checkModel09.State = int(config.getboolean("doc", "reserve position numbers"))
    checkModel09.Label = "Резервировать номера позиций"
    checkModel09.HelpText = """\
По умолчанию, позиции в спецификации
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

    checkModel010 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel010.PositionX = 5
    checkModel010.PositionY = checkModel08.PositionY + checkModel08.Height
    checkModel010.Width = tabsModel.Width - 10
    checkModel010.Height = checkModel00.Height
    checkModel010.Name = "CheckBox010"
    checkModel010.State = int(config.getboolean("doc", "footprint only"))
    checkModel010.Label = "\"Посад.место\" без наименования библиотеки"
    checkModel010.HelpText = """\
Если отмечено, то посадочное место
будет указано без наименования библиотеки."""
    pageModel0.insertByName("CheckBox010", checkModel010)

    checkModel011 = pageModel0.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel011.PositionX = 5
    checkModel011.PositionY = checkModel010.PositionY + checkModel010.Height
    checkModel011.Width = tabsModel.Width - 10
    checkModel011.Height = checkModel00.Height
    checkModel011.Name = "CheckBox011"
    checkModel011.State = int(config.getboolean("doc", "split row by \\n"))
    checkModel011.Label = "Обрабатывать \"\\n\" как переход на новую строку"
    checkModel011.HelpText = """\
Если отмечено, то комбинация символов
"\\n" будет обрабатываться как переход
на следующую строку таблицы."""
    pageModel0.insertByName("CheckBox011", checkModel011)

    # ------------------------------------------------------------------------
    # Sections Tab Model
    # ------------------------------------------------------------------------

    pageModel3 = tabsModel.createInstance(
        "com.sun.star.awt.UnoPageModel"
    )
    tabsModel.insertByName("Page3", pageModel3)
    pageModel3.Title = " Разделы "

    checkModel30 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel30.PositionX = 5
    checkModel30.PositionY = 5
    checkModel30.Width = tabsModel.Width - 10
    checkModel30.Height = 15
    checkModel30.Name = "CheckBox30"
    checkModel30.State = int(config.getboolean("sections", "documentation"))
    checkModel30.Label = "Документация"
    checkModel30.HelpText = """\
Если отмечено, то при формировании
спецификации будет создан раздел
"Документация"."""
    pageModel3.insertByName("CheckBox30", checkModel30)

    checkModel31 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel31.PositionX = checkModel30.PositionX + 10
    checkModel31.PositionY = checkModel30.PositionY + checkModel30.Height
    checkModel31.Width = checkModel30.Width
    checkModel31.Height = checkModel30.Height
    checkModel31.Name = "CheckBox31"
    checkModel31.State = int(config.getboolean("sections", "assembly drawing"))
    checkModel31.Label = "Сборочный чертёж"
    checkModel31.HelpText = """\
Если отмечено, то при формировании
спецификации в разделе "Документация"
будет указан сборочный чертёж."""
    pageModel3.insertByName("CheckBox31", checkModel31)

    checkModel32 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel32.PositionX = checkModel30.PositionX + 10
    checkModel32.PositionY = checkModel30.PositionY + checkModel30.Height * 2
    checkModel32.Width = checkModel30.Width
    checkModel32.Height = checkModel30.Height
    checkModel32.Name = "CheckBox32"
    checkModel32.State = int(config.getboolean("sections", "schematic"))
    checkModel32.Label = "Схема электрическая принципиальная"
    checkModel32.HelpText = """\
Если отмечено, то при формировании
спецификации в разделе "Документация"
будет указана принципиальная схема."""
    pageModel3.insertByName("CheckBox32", checkModel32)

    checkModel33 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel33.PositionX = checkModel30.PositionX + 10
    checkModel33.PositionY = checkModel30.PositionY + checkModel30.Height * 3
    checkModel33.Width = checkModel30.Width
    checkModel33.Height = checkModel30.Height
    checkModel33.Name = "CheckBox33"
    checkModel33.State = int(config.getboolean("sections", "index"))
    checkModel33.Label = "Перечень элементов"
    checkModel33.HelpText = """\
Если отмечено, то при формировании
спецификации в разделе "Документация"
будет указан перечень элементов."""
    pageModel3.insertByName("CheckBox33", checkModel33)

    checkModel310 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel310.PositionX = checkModel30.PositionX + 10
    checkModel310.PositionY = checkModel30.PositionY + checkModel30.Height * 4
    checkModel310.Width = checkModel30.Height
    checkModel310.Height = checkModel30.Height
    checkModel310.Name = "CheckBox310"
    checkModel310.State = int(config.getboolean("sections", "bom"))
    checkModel310.Label = ""
    checkModel310.HelpText = """\
Если отмечено, то при формировании
спецификации в разделе "Документация"
будет указана ведомость покупных изделий."""
    pageModel3.insertByName("CheckBox310", checkModel310)

    editControlModel311 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlEditModel"
    )
    editControlModel311.PositionX = checkModel310.PositionX + checkModel310.Width
    editControlModel311.PositionY = checkModel310.PositionY
    editControlModel311.Width = tabsModel.Width - editControlModel311.PositionX - 3
    editControlModel311.Height = checkModel310.Height
    editControlModel311.Name = "EditControl311"
    editControlModel311.Text = config.get("sections", "bom name")
    pageModel3.insertByName("EditControl311", editControlModel311)

    checkModel34 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel34.PositionX = checkModel30.PositionX
    checkModel34.PositionY = checkModel30.PositionY + checkModel30.Height * 5
    checkModel34.Width = checkModel30.Width
    checkModel34.Height = checkModel30.Height
    checkModel34.Name = "CheckBox34"
    checkModel34.State = int(config.getboolean("sections", "assembly units"))
    checkModel34.Label = "Сборочные единицы"
    checkModel34.HelpText = """\
Если отмечено, то при формировании
спецификации будет создан раздел
"Сборочные единицы"."""
    pageModel3.insertByName("CheckBox34", checkModel34)

    checkModel312 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel312.PositionX = checkModel30.PositionX + 10
    checkModel312.PositionY = checkModel30.PositionY + checkModel30.Height * 6
    checkModel312.Width = checkModel30.Width
    checkModel312.Height = checkModel30.Height
    checkModel312.Name = "CheckBox312"
    checkModel312.State = int(config.getboolean("sections", "assembly pcb"))
    checkModel312.Label = "Плата печатная"
    checkModel312.HelpText = """\
Если отмечено, то при формировании
спецификации в разделе "Сборочные единицы"
будет указана печатная плата."""
    pageModel3.insertByName("CheckBox312", checkModel312)

    checkModel35 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel35.PositionX = checkModel30.PositionX
    checkModel35.PositionY = checkModel30.PositionY + checkModel30.Height * 7
    checkModel35.Width = checkModel30.Width
    checkModel35.Height = checkModel30.Height
    checkModel35.Name = "CheckBox35"
    checkModel35.State = int(config.getboolean("sections", "details"))
    checkModel35.Label = "Детали"
    checkModel35.HelpText = """\
Если отмечено, то при формировании
спецификации будет создан раздел
"Детали"."""
    pageModel3.insertByName("CheckBox35", checkModel35)

    checkModel36 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel36.PositionX = checkModel30.PositionX + 10
    checkModel36.PositionY = checkModel30.PositionY + checkModel30.Height * 8
    checkModel36.Width = checkModel30.Width
    checkModel36.Height = checkModel30.Height
    checkModel36.Name = "CheckBox36"
    checkModel36.State = int(config.getboolean("sections", "pcb"))
    checkModel36.Label = "Плата печатная"
    checkModel36.HelpText = """\
Если отмечено, то при формировании
спецификации в разделе "Детали"
будет указана печатная плата."""
    pageModel3.insertByName("CheckBox36", checkModel36)

    checkModel37 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel37.PositionX = checkModel30.PositionX
    checkModel37.PositionY = checkModel30.PositionY + checkModel30.Height * 9
    checkModel37.Width = checkModel30.Width
    checkModel37.Height = checkModel30.Height
    checkModel37.Name = "CheckBox37"
    checkModel37.State = int(config.getboolean("sections", "standard parts"))
    checkModel37.Label = "Стандартные изделия"
    checkModel37.HelpText = """\
Если отмечено, то при формировании
спецификации будет создан раздел
"Стандартные изделия"."""
    pageModel3.insertByName("CheckBox37", checkModel37)

    checkModel38 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel38.PositionX = checkModel30.PositionX
    checkModel38.PositionY = checkModel30.PositionY + checkModel30.Height * 10
    checkModel38.Width = checkModel30.Width
    checkModel38.Height = checkModel30.Height
    checkModel38.Name = "CheckBox38"
    checkModel38.State = int(config.getboolean("sections", "other parts"))
    checkModel38.Label = "Прочие изделия"
    checkModel38.HelpText = """\
Если отмечено, то при формировании
спецификации будет создан раздел
"Прочие изделия"."""
    pageModel3.insertByName("CheckBox38", checkModel38)

    checkModel39 = pageModel3.createInstance(
        "com.sun.star.awt.UnoControlCheckBoxModel"
    )
    checkModel39.PositionX = checkModel30.PositionX
    checkModel39.PositionY = checkModel30.PositionY + checkModel30.Height * 11
    checkModel39.Width = checkModel30.Width
    checkModel39.Height = checkModel30.Height
    checkModel39.Name = "CheckBox39"
    checkModel39.State = int(config.getboolean("sections", "materials"))
    checkModel39.Label = "Материалы"
    checkModel39.HelpText = """\
Если отмечено, то при формировании
спецификации будет создан раздел
"Материалы"."""
    pageModel3.insertByName("CheckBox39", checkModel39)

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

    labelModel15 = pageModel1.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel"
    )
    labelModel15.PositionX = 0
    labelModel15.PositionY = labelModel10.Height * 4
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
спецификации."""
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
спецификации из файла настроек
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
    checkModel20.State = int(config.getboolean("stamp", "convert doc title"))
    checkModel20.Label = "Преобразовать наименование документа"
    checkModel20.HelpText = """\
Если отмечено, тип схемы в наименовании
документа будет удалён.
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
Если отмечено, тип схемы в
обозначении документа будет
удалён.
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
        config.setboolean("settings", "set view options",
            self.dialog.getControl("CheckBoxSet").State
        )

        # --------------------------------------------------------------------
        # Спецификация
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
        if page0.getControl("RadioButton00").State:
            config.set("doc", "ref separator", "-")
        if page0.getControl("RadioButton01").State:
            config.set("doc", "ref separator", "…")
        config.setboolean("doc", "add units",
            page0.getControl("CheckBox00").State
        )
        config.setboolean("doc", "space before units",
            page0.getControl("CheckBox01").State
        )
        config.setboolean("doc", "separate group for each doc",
            page0.getControl("CheckBox02").State
        )
        config.setboolean("doc", "title with doc",
            page0.getControl("CheckBox03").State
        )
        config.setboolean("doc", "every group has title",
            page0.getControl("CheckBox04").State
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
        config.setboolean("doc", "footprint only",
            page0.getControl("CheckBox010").State
        )
        config.setboolean("doc", "split row by \\n",
            page0.getControl("CheckBox011").State
        )

        # --------------------------------------------------------------------
        # Разделы
        # --------------------------------------------------------------------

        config.setboolean("sections", "documentation",
            page3.getControl("CheckBox30").State
        )
        config.setboolean("sections", "assembly drawing",
            page3.getControl("CheckBox31").State
        )
        config.setboolean("sections", "schematic",
            page3.getControl("CheckBox32").State
        )
        config.setboolean("sections", "index",
            page3.getControl("CheckBox33").State
        )
        config.setboolean("sections", "bom",
            page3.getControl("CheckBox310").State
        )
        config.set("sections", "bom name",
            page3.getControl("EditControl311").Text
        )
        config.setboolean("sections", "assembly units",
            page3.getControl("CheckBox34").State
        )
        config.setboolean("sections", "assembly pcb",
            page3.getControl("CheckBox312").State
        )
        config.setboolean("sections", "details",
            page3.getControl("CheckBox35").State
        )
        config.setboolean("sections", "pcb",
            page3.getControl("CheckBox36").State
        )
        config.setboolean("sections", "standard parts",
            page3.getControl("CheckBox37").State
        )
        config.setboolean("sections", "other parts",
            page3.getControl("CheckBox38").State
        )
        config.setboolean("sections", "materials",
            page3.getControl("CheckBox39").State
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
        config.set("fields", "excluded",
            page1.getControl("EditControl15").Text
        )
        config.setboolean("settings", "compatibility mode",
            page1.getControl("CheckBox10").State
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
            )
            page1.getControl("CheckBox10").State = 1
        else:
            page1.getControl("CheckBox10").State = 0
        for control, value in defaultValues:
            page1.getControl(control).Text = value
