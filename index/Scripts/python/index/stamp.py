import re
import common
import config
import textwidth

common.XSCRIPTCONTEXT = XSCRIPTCONTEXT
config.XSCRIPTCONTEXT = XSCRIPTCONTEXT

def syncCommonFields(*args):
    """Обновить значения общих граф.

    Обновить значения граф форматной рамки последующих листов, которые
    совпадают с графами форматной рамки и основной надписи первого листа.
    К таким графам относятся:
    - 2. Обозначение документа;
    - 19. Инв. № подл.;
    - 21. Взам. инв. №;
    - 22. Инв. № дубл.
    Необходимость в обновлении возникает при изменении значения графы
    на первом листе.
    На втором и последующих листах эти графы защищены от записи.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    for name in common.STAMP_COMMON_FIELDS:
        doc.getTextFrames().getByName("N." + name).setString(
            doc.getTextFrames().getByName("1.1." + name).getString()
    )

def setFirstPageFrameValue(name, value):
    """Установить значение поля форматной рамки первого листа.

    Имеется четыре стиля для первого листа и у каждого содержится
    свой набор полей. Они должны иметь одинаковые значения, чтобы
    при смене стиля значения сохранялись.

    Атрибуты:
    name -- имя текстового поля (без первых 4-х символов);
    value -- новое значение поля.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    for i in range(1, 5):
        fullName = "1.{}.{}".format(i, name)
        if doc.getTextFrames().hasByName(fullName):
            frame = doc.getTextFrames().getByName(fullName)
            if name in ("11 Разраб.", "11 Пров.", "11 Утв."):
                frame.setString("")
                cursor = frame.Text.createTextCursor()
                cursor.CharScaleWidth = textwidth.getWidthFactor("ФИО", value)
            frame.setString(value)

def clean(*args):
    """Очистить основную надпись.

    Удалить содержимое полей основной надписи и форматной рамки.

    """
    if common.isThreadWorking():
        return
    doc = XSCRIPTCONTEXT.getDocument()
    doc.lockControllers()
    for frame in doc.getTextFrames():
        if frame.Name.startswith("1.") \
            and not frame.Name.endswith("(наименование)") \
            and not frame.Name.endswith(".7 Лист") \
            and not frame.Name.endswith(".8 Листов"):
                frame.setString("")
                if frame.Name.endswith(".11 Разраб.") \
                    or frame.Name.endswith(".11 Пров.") \
                    or frame.Name.endswith(".11 Утв."):
                        cursor = frame.Text.createTextCursor()
                        cursor.CharScaleWidth = 100
    syncCommonFields()
    doc.unlockControllers()

def fill(*args):
    """Заполнить основную надпись.

    Считать данные из файла списка цепей и заполнить графы основной надписи.
    Имеющиеся данные не удаляются из основной надписи.
    Новое значение вносится только в том случае, если оно не пустое.

    """
    if common.isThreadWorking():
        return
    schematic = common.getSchematicData()
    if schematic is None:
        return
    doc = XSCRIPTCONTEXT.getDocument()
    doc.lockControllers()
    settings = config.load()
    # Наименование документа
    docTitle = schematic.title.replace('\\n', '\n')
    if settings.getboolean("stamp", "convert doc title"):
        suffix = "Перечень элементов"
        titleParts = docTitle.rsplit("Схема электрическая ", 1)
        schTypes = (
            "структурная",
            "функциональная",
            "принципиальная",
            "соединений",
            "подключения",
            "общая",
            "расположения"
        )
        if len(titleParts) > 1 and titleParts[1] in schTypes:
            # Оставить только наименование изделия
            docTitle = titleParts[0]
        if docTitle:
            # Только один перевод строки
            docTitle = docTitle.rstrip('\n') + '\n'
        docTitle = docTitle + suffix
    setFirstPageFrameValue("1 Наименование документа", docTitle)
    # Наименование организации
    setFirstPageFrameValue("9 Наименование организации", schematic.company)
    # Обозначение документа
    docId = schematic.number
    idParts = re.match(
        r"([А-ЯA-Z0-9]+(?:[\.\-][0-9]+)+\s?)(Э[1-7])",
        docId
    )
    if settings.getboolean("stamp", "convert doc id") \
        and idParts is not None:
            docId = 'П'.join(idParts.groups())
    setFirstPageFrameValue("2 Обозначение документа", docId)
    # Первое применение
    if settings.getboolean("stamp", "fill first usage") \
        and idParts is not None:
            setFirstPageFrameValue("25 Перв. примен.", idParts.group(1).strip())
    # Разработал
    setFirstPageFrameValue("11 Разраб.", schematic.developer)
    # Проверил
    setFirstPageFrameValue("11 Пров.", schematic.verifier)
    # Утвердил
    setFirstPageFrameValue("11 Утв.", schematic.approver)

    syncCommonFields()
    doc.unlockControllers()
