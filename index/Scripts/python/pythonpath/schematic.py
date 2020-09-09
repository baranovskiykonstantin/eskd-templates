"""Объектное представление схемы."""

import re
import sys

kicadnet = None
config = None

def init(scriptcontext):
    global kicadnet
    global config
    kicadnet = sys.modules["kicadnet" + scriptcontext.getDocument().RuntimeUID]
    config = sys.modules["config" + scriptcontext.getDocument().RuntimeUID]

REF_REGEXP = re.compile(r"([^0-9?]+)([0-9]+)")


class Component():
    """Данные о компоненте схемы."""

    def __init__(self, schematic):
        self.schematic = schematic
        self.reference = ""
        self.value = ""
        self.footprint = ""
        self.datasheet = ""
        self.description = ""
        self.fields = {}

    def getFieldValue(self, name):
        """Вернуть значение поля с указанным именем."""
        value = None
        if name == "Обозначение":
            value = self.reference
        elif name == "Значение":
            if config.getboolean("doc", "add units"):
                value = self.getValueWithUnits()
            else:
                value = self.value
        elif name == "Посад.место":
            if config.getboolean("doc", "footprint only"):
                value = self.getFieldValue("Посад.место!")
            else:
                value = self.footprint
        elif name == "Посад.место!":
            value = self.footprint
            if ':' in value:
                # Удалить наименование библиотеки включительно с двоеточием
                value = value[(value.index(':') + 1):]
        elif name == "Документация":
            value = self.datasheet
        elif name == "Описание":
            value = self.description
        elif name in self.fields:
            value = self.fields[name]
        if value:
            value = self.formatPattern(value)
        return value

    def getRefType(self, ref=None):
        """Вернуть буквенную часть обозначения."""
        if ref is None:
            ref = self.reference
        if not re.match(REF_REGEXP, ref):
            return None
        refType = re.search(REF_REGEXP, ref).group(1)
        return refType

    def getRefNumber(self, ref=None):
        """Вернуть цифровую часть обозначения."""
        if ref is None:
            ref = self.reference
        if not re.match(REF_REGEXP, ref):
            return None
        refNumber = re.search(REF_REGEXP, ref).group(2)
        return int(refNumber)

    def _convertSingularPlural(self, value, singular, plural):
        """Привести переданное значение к единственному либо множественному числу.

        Если параметр plural==True, то значение поля будет указано в
        множественном числе.
        Если параметр singular==True, то значение поля будет указано в
        единственном числе.
        Если значение поля имеет формат:
        значение 1 {значение 2}
        то "значение 1" воспринимается как значение поля в единственном числе,
        а "значение 2" - как значение в множественном числе.
        Если значение поля не соответствует указанному формату, то это значение
        будет использоваться полностью как в единственном, так и в
        множественном числе.

        Аргументы:
        value (str) -- значение поля, которое необходимо обработать;
        singular (boolean) -- привести к единственному числу;
        plural (boolean) -- привести к множественному числу.

        Возвращаемое значение (str) -- преобразованное значение.

        """

        if value and (singular or plural):
            valueSingularAndPlural = re.match(r"^(.+)\s\{(.+)\}$", value)
            if valueSingularAndPlural:
                if singular:
                    value = valueSingularAndPlural.group(1)
                elif plural:
                    value = valueSingularAndPlural.group(2)
            elif self.schematic.typeNamesDict:
                for item in iter(self.schematic.typeNamesDict.items()):
                    if value in item:
                        if singular:
                            value = item[0]
                        elif plural:
                            value = item[1]
        return value

    def getValueWithUnits(self):
        """Преобразовать значение к стандартному виду.

        Возвращаемое значение -- значение элемента, приведённое к
            стандартному виду, например:
            2u7 -> 2,7 мкФ

        """
        multipliersDict = {
            'G': 'Г',
            'M': 'М',
            'k': 'к',
            'm': 'м',
            'μ': 'мк',
            'u': 'мк',
            'U': 'мк',
            'n': 'н',
            'p': 'п'
        }
        numValue = ""
        separator = ""
        if config.getboolean("doc", "space before units"):
            separator = ' '
        multiplier = ""
        units = ""
        multipliers = set(list(multipliersDict.keys()) + list(multipliersDict.values()))
        # 2u7, 2н7, 4m7, 5k1 ...
        regexpr1 = re.compile(
            r"^(\d+)({})(\d+)$".format('|'.join(multipliers))
        )
        # 2.7 u, 2700p, 4.7 m, 470u, 5.1 k, 510 ...
        regexpr2 = re.compile(
            r"^(\d+(?:[\.,]\d+)?)\s*({})?$".format('|'.join(multipliers))
        )
        if self.getRefType().startswith('C') \
            and not self.value.endswith('Ф'):
                units = 'Ф'
                if re.match(r"^\d+$", self.value):
                    numValue = self.value
                    multiplier = 'п'
                elif re.match(r"^\d+[\.,]\d+$", self.value):
                    numValue = self.value
                    multiplier = "мк"
                else:
                    numValue = self.value.rstrip('F')
                    numValue = numValue.strip()
                    if re.match(regexpr1, numValue):
                        searchRes = re.search(regexpr1, numValue).groups()
                        numValue = "{},{}".format(searchRes[0], searchRes[2])
                        multiplier = searchRes[1]
                    elif re.match(regexpr2, numValue):
                        searchRes = re.search(regexpr2, numValue).groups()
                        numValue = searchRes[0]
                        multiplier = searchRes[1]
                    else:
                        numValue = ""
        elif self.getRefType().startswith('L') \
            and not self.value.endswith("Гн"):
                units = "Гн"
                numValue = self.value.rstrip('H')
                numValue = numValue.strip()
                if re.match(regexpr1, numValue):
                    searchRes = re.search(regexpr1, numValue).groups()
                    numValue = "{},{}".format(searchRes[0], searchRes[2])
                    multiplier = searchRes[1]
                elif re.match(regexpr2, numValue):
                    searchRes = re.search(regexpr2, numValue).groups()
                    numValue = searchRes[0]
                    if searchRes[1] is None:
                        multiplier = "мк"
                    else:
                        multiplier = searchRes[1]
                else:
                    numValue = ""
        elif self.getRefType().startswith('R') \
            and not self.value.endswith("Ом"):
                units = "Ом"
                numValue = self.value.rstrip('Ω')
                if numValue.endswith("Ohm") or numValue.endswith("ohm"):
                    numValue = numValue[:-3]
                numValue = numValue.strip()
                if re.match(r"R\d+", numValue):
                    numValue = numValue.replace('R', "0,")
                elif re.match(r"\d+R\d+", numValue):
                    numValue = numValue.replace('R', ',')
                elif re.match(regexpr1, numValue):
                    searchRes = re.search(regexpr1, numValue).groups()
                    numValue = "{},{}".format(searchRes[0], searchRes[2])
                    multiplier = searchRes[1]
                elif re.match(regexpr2, numValue):
                    searchRes = re.search(regexpr2, numValue).groups()
                    numValue = searchRes[0]
                    if searchRes[1] is not None:
                        multiplier = searchRes[1]
                else:
                    numValue = ""
        if numValue:
            # Перевести множитель на русский
            if multiplier in multipliersDict:
                multiplier = multipliersDict[multiplier]
            elif multiplier is None:
                multiplier = ''
            numValue = numValue.replace('.', ',')
            return numValue + separator + multiplier + units
        return self.value

    def formatPattern(self, pattern, check=False, singular=False, plural=False):
        """Преобразовать шаблон.

        Шаблон представляет собой строку текста, в которой конструкции типа:

        ${НаименованиеПоля}
        ${Префикс|НаименованиеПоля|Суффикс}

        будут преобразованы в текст вида:

        ЗначениеПоля
        ПрефиксЗначениеПоляСуффикс

        Например:
        "МЛТ-0,5-${Значение}${-|Класс точности|}-В" -> "МЛТ-0,5-4,7кОм-±5%-В"
        Если значение поля пусто или указанного поля нет в компоненте, то
        соответствующий элемент шаблона удаляется. Если, допустим, для
        приведённого выше примера, в компоненте нет поля "Класс точности", то
        результат будет следующим:
        "МЛТ-0,5-4,7кОм-В" (префикс '-' тоже отсутствует)

        Символы '{', '|', '}' имеют специальное назначение. Если в шаблоне
        требуется указать эти символы, то их нужно экранировать символом
        обратной косой черты ' \ ', например:
        "Обозначение компонента ${\{|Обозначение|\}} в фигурных скобках."
        Но спец. символы вне конструкции ${} экранировать не нужно:
        "Обозначение компонента {${Обозначение}} в фигурных скобках."

        Если параметр check==True, то вместо преобразования строки будет
        выполнена проверка - является ли переданная строка шаблоном. При первом
        обнаружении конструкции ${} будет возвращено значение True, при
        отсутствии такой конструкции - False.

        Аргументы:
        pattern (str) -- строка текста, которую следует обработать как шаблон;
        check (boolean) -- проверить шаблон без преобразования;
        singular (boolean) -- привести к единственному числу;
        plural (boolean) -- привести к множественному числу.

        Возвращаемое значение (str) -- преобразованное значение.

        """
        out = ""
        prefix = ""
        fieldName = ""
        suffix = ""
        temp = ""

        # Флаг, указывающий на то, что спец.символ нужно обработать как обычный
        ignore = False
        # Флаг, указывающий на обрабатываемую часть подстановки.
        substitution = ""

        def resetSubstitution():
            nonlocal out, temp, substitution, prefix, fieldName, suffix
            out += temp
            substitution = temp = ""
            prefix = fieldName = suffix = ""

        for char in pattern:
            if char == '\\' and substitution and not ignore:
                    ignore = True
                    temp += char
                    continue
            elif substitution:
                temp += char
                if substitution == "beginning":
                    if char == '{' and not ignore:
                        substitution = "prefix"
                    else:
                        out += temp
                        substitution = temp = ""
                elif char == '{' and not ignore:
                    # Конструкция ${} имеет неверный формат:
                    # ${...{
                    #      ^
                    # открывающаяся фигурная скобка внутри подстановки.
                    resetSubstitution()
                elif char == '|' and substitution == "prefix" and not ignore:
                    substitution = "fieldName"
                elif char == '|' and substitution == "fieldName" and not ignore:
                    substitution = "suffix"
                elif char == '|' and substitution == "suffix" and not ignore:
                    # Конструкция ${} имеет неверный формат:
                    # ${prefix|fieldName|suffix|
                    #                          ^
                    # третья вертикальная черта внутри подстановки.
                    resetSubstitution()
                elif char == "}" and not ignore:
                    if substitution == "fieldName":
                        # Конструкция ${} имеет неверный формат:
                        # ${prefix|fieldName}
                        #                   ^
                        # одна вертикальная черта в подстановке. Должно быть
                        # либо две (для пефикса/суффикса), либо не быть вовсе.
                        resetSubstitution()
                    else:
                        if substitution == "prefix":
                            # Если по завершении конструкции ${} имеется только
                            # префикс, значит найдена сокращённая конструкция
                            # (без префикса/суффикса).
                            fieldName = prefix
                            prefix = ""
                        if check:
                            return True
                        fieldValue = self.getFieldValue(fieldName)
                        if fieldValue:
                            fieldValue = self._convertSingularPlural(fieldValue, singular, plural)
                            out += prefix + fieldValue + suffix
                    substitution = temp = prefix = fieldName = suffix = ""
                elif substitution == "prefix":
                    prefix += char
                elif substitution == "fieldName":
                    fieldName += char
                elif substitution == "suffix":
                    suffix += char
            elif char == '$':
                substitution = "beginning"
                temp += char
            else:
                out += char
            ignore = False
        if substitution:
            # Конструкция ${} неожиданно закончилась.
            resetSubstitution()
        if check:
            return False
        return out

    def getIndexValue(self, name, singular=False, plural=False):
        """Вернуть преобразованное значение для перечня.

        Вернуть приведённое к конечному виду значение одного из полей,
        используемых при построении перечня.

        Аргументы:
        name (str) -- название требуемого значения; может быть одним из:
            "type", "name", "doc", "comment";
        singular (boolean) -- привести к единственному числу;
        plural (boolean) -- привести к множественному числу.

        Возвращаемое значение (str) -- итоговое значение.

        """
        if name not in ("type", "name", "doc", "comment"):
            return ""
        fieldName = config.get("fields", name)
        value = ""
        if self.formatPattern(fieldName, check=True):
            value = self.formatPattern(fieldName, singular=singular, plural=plural)
        else:
            value = self.getFieldValue(fieldName)
            value = self._convertSingularPlural(value, singular, plural)
        if name == "name" and not value:
            if config.getboolean("doc", "add units"):
                value = self.getValueWithUnits()
            else:
                value = self.value
        if value is None:
            value = ""
        return value


class CompRange(Component):
    """Множество компонентов с одинаковыми параметрами.

    Этот класс описывает множество компонентов перечня
    элементов, которые имеют одинаковые тип, наименование, документ,
    примечание, буквенную часть обозначения и следуют последовательно.

    """

    def __init__(self, schematic, comp=None):
        Component.__init__(self, schematic)
        self._refRange = []
        if comp is not None:
            self._refRange.append(comp.reference)
            self.reference = comp.reference
            self.value = comp.value
            self.footprint = comp.footprint
            self.datasheet = comp.datasheet
            self.description = comp.description
            self.fields = comp.fields

    def __iter__(self):
        for ref in self._refRange:
            yield ref

    def __len__(self):
        return len(self._refRange)

    def append(self, comp):
        """Добавить новый компонент.

        Добавить компонент в множество одинаковых компонентов.
        Если компонент отличается от имеющихся, то он не будет добавлен.

        Аргументы:
        comp (Component) -- компонент, который необходимо добавить.

        Возвращаемые значения (boolean) -- True - если компонент был добавлен,
            False - в противном случае.

        """
        if not self._refRange:
            self.__init__(self.schematic, comp)
            return True
        if self.getRefType() == comp.getRefType() \
            and self.getIndexValue("type") == comp.getIndexValue("type") \
            and self.getIndexValue("name") == comp.getIndexValue("name") \
            and self.getIndexValue("doc") == comp.getIndexValue("doc") \
            and self.getIndexValue("comment") == comp.getIndexValue("comment"):
                self._refRange.append(comp.reference)
                return True
        return False

    def getRefRangeString(self):
        """Вернуть перечень обозначений множества одинаковых компонентов."""
        refStr = ""
        adjustable = False
        adjustableField = config.get("fields", "adjustable")
        if self.getFieldValue(adjustableField) is not None:
            adjustable = True
        if len(self._refRange) > 1:
            # "VD1, VD2", "C8-C11", "R7, R9-R14", "C8*-C11*" ...
            prevType = self.getRefType(self._refRange[0])
            prevNumber = self.getRefNumber(self._refRange[0])
            counter = 0
            separator = ", "
            refStr = prevType + str(prevNumber)
            if adjustable:
                refStr += '*'
            for nextRef in self._refRange[1:]:
                currentType = self.getRefType(nextRef)
                currentNumber = self.getRefNumber(nextRef)
                if currentType == prevType \
                    and currentNumber == (prevNumber + 1):
                        prevNumber = currentNumber
                        counter += 1
                        if counter > 1:
                            separator = config.get("doc", "ref separator")
                        continue
                else:
                    if counter > 0:
                        refStr += separator + prevType + str(prevNumber)
                        if adjustable:
                            refStr += '*'
                    separator = ', '
                    refStr += separator + currentType + str(currentNumber)
                    if adjustable:
                        refStr += '*'
                    prevType = currentType
                    prevNumber = currentNumber
                    counter = 0
            if counter > 0:
                refStr += separator + prevType + str(prevNumber)
                if adjustable:
                    refStr += '*'
        else:
            # "R5"; "VT13" ...
            refStr = self.reference
            if adjustable:
                refStr += '*'
        return refStr


class CompGroup():
    """Группа компонентов.

    Группой считается множество CompRange, которые имеют однотипные
    обозначения (например: R, C, DA и т.д.) и имеют одинаковый "Тип".

    Если установлен параметр "concatenate same name groups", то
    группы, идущие подряд и имеющие одинаковое наименование типа, но
    отличающиеся обозначением - будут объединены. Например, компоненты типа
    "Разъём (Разъёмы)", но с обозначениями "XP..." и "XS...", по умолчанию
    формируют две отдельные группы с одинаковым заголовком. Но, с помощью выше
    указанного параметра, эти группы могут быть объединены в одну.

    """

    def __init__(self, schematic, compRange=None):
        self.schematic = schematic
        self._compRanges = []
        if compRange is not None:
            self._compRanges.append(compRange)

    def __iter__(self):
        for compRange in self._compRanges:
            yield compRange

    def __getitem__(self, key):
        return self._compRanges[key]

    def __len__(self):
        return len(self._compRanges)

    def append(self, compRange):
        """Добавить множество компонентов в группу.

        Аргументы:
        compRange (CompRange) -- множество компонентов, которое необходимо
            добавить в группу.

        Возвращаемые значения (boolean) -- True - если множество было
            добавлено, False - в противном случае.

        """
        if not self._compRanges:
            self._compRanges.append(compRange)
            return True
        skipRefType = config.getboolean("doc", "concatenate same name groups")
        lastCompRange = self._compRanges[-1]
        if (lastCompRange.getRefType() == compRange.getRefType() or skipRefType) \
            and lastCompRange.getIndexValue("type") == compRange.getIndexValue("type"):
                self._compRanges.append(compRange)
                return True
        return False

    @staticmethod
    def _shortenName(name):
        """Сократить имя до наименее возможного.

        Сократить имя до первого пробела, дефиса или конца строки.

        Аргументы:
        name (str) -- имя для сокращения.

        Возвращаемое значение (str) -- сокращённое имя.

        """
        index = len(name)
        if " " in name:
            index = name.index(" ")
        if "-" in name:
            index = min(index, name.index("-"))
        name = name[:index]
        name = name.rstrip(" -")
        return name

    def getTitle(self):
        """Вернуть заголовок группы компонентов.

        По умолчанию, заголовком является "Тип" компонента.
        Если установлен параметр "title with doc", то после "Типа", через
        пробел, будет указан "Документ".
        Если "Документы" компонентов группы отличаются, то заголовок будет
        состоять из нескольких строк, каждая из которых служит для отдельного
        документа. При этом перед каждым документом будет указана часть
        наименования для идентификации компонентов.

        Возвращаемое значение (list) -- список строк заголовка.

        """
        if len(self) == 0:
            return []

        currentType = self._compRanges[0].getIndexValue("type", plural=True)

        if not config.getboolean("doc", "title with doc"):
            return [currentType]

        # Список уникальных пар Наименование-Документ
        nameDocList = []
        for compRange in self:
            currentName = compRange.getIndexValue("name")
            currentShortestName = self._shortenName(currentName)
            currentDoc = compRange.getIndexValue("doc")
            if not currentDoc:
                # Если имеются компоненты, в которых документ не указан,
                # то в заголовке для них будет указан только тип.
                currentName = ""
            for i in range(len(nameDocList)):
                savedName = nameDocList[i][0]
                savedShortestName = self._shortenName(savedName)
                savedDoc = nameDocList[i][1]
                if savedDoc == currentDoc \
                        and savedShortestName == currentShortestName:
                    break
            else:
                nameDocList.append([currentName, currentDoc])

        # Максимально сократить наименования, оставив только часть
        # достаточную для идентификации.
        for i in range(len(nameDocList)):
            name = nameDocList[i][0]
            doc = nameDocList[i][1]
            nameParts = re.findall(r"([-\s]?[^-\s]+)", name)
            if len(nameParts) > 1:
                for j in range(1, len(nameParts)):
                    shortName = "".join(nameParts[:j])
                    if [shortName, doc] not in nameDocList:
                        nameDocList[i][0] = shortName
                        break

        # Сформировать наименование
        if not nameDocList:
            return [currentType]
        firstDoc = nameDocList[0][-1]
        for name, doc in nameDocList:
            if doc != firstDoc:
                break
        else:
            # У всех компонентов один документ
            title = currentType
            if title:
                title += ' '
            title += firstDoc
            return [title]
        groupNames = []
        nameDocList.sort(key=lambda nameDoc: nameDoc[0])
        for nameDoc in nameDocList:
            name = currentType
            if nameDoc[0]:
                if name:
                    name += ' '
                name += nameDoc[0]
            if nameDoc[1]:
                if name:
                    name += ' '
                name += nameDoc[1]
            groupNames.append(name)

        return groupNames


class Schematic():
    """Данные о схеме и компонентах."""

    def __init__(self, netlistName):
        self.title = ""
        self.number = ""
        self.company = ""
        self.developer = ""
        self.verifier = ""
        self.inspector = ""
        self.approver = ""
        self.components = []

        self.typeNamesDict = {}
        if config.getboolean("settings", "compatibility mode"):
            # KB2S - kicadbom2spec
            settingsKB2S = config.loadFromKicadbom2spec()
            if settingsKB2S is not None:
                if settingsKB2S.has_section('group names singular'):
                    for index in settingsKB2S.options('group names singular'):
                        if settingsKB2S.has_option('group names plural', index):
                            singular = settingsKB2S.get('group names singular', index)
                            plural = settingsKB2S.get('group names plural', index)
                            self.typeNamesDict[singular] = plural

        netlist = kicadnet.Netlist(netlistName)
        for sheet in netlist.items("sheet"):
            if sheet.attributes["name"] == "/":
                title_block = netlist.find("title_block", sheet)
                for item in title_block.items:
                    if item.name == "title":
                        self.title = item.text if item.text is not None else ""
                    elif item.name == "company":
                        self.company = item.text if item.text is not None else ""
                    elif item.name == "comment":
                        if item.attributes["number"] == "1":
                            self.number = item.attributes["value"]
                        elif item.attributes["number"] == "2":
                            self.developer = item.attributes["value"]
                        elif item.attributes["number"] == "3":
                            self.verifier = item.attributes["value"]
                        elif item.attributes["number"] == "4":
                            self.approver = item.attributes["value"]
                        elif item.attributes["number"] == "6":
                            self.inspector = item.attributes["value"]
                break
        for comp in netlist.items("comp"):
            component = Component(self)
            component.reference = comp.attributes["ref"]
            for item in comp.items:
                if item.name == "value":
                    component.value = item.text if item.text is not None and item.text != "~" else ""
                elif item.name == "footprint":
                    component.footprint = item.text if item.text is not None and item.text != "~" else ""
                elif item.name == "datasheet":
                    component.datasheet = item.text if item.text is not None and item.text != "~" else ""
                elif item.name == "libsource":
                    if "description" in item.attributes:
                        component.description = item.attributes["description"]
                elif item.name == "fields":
                    for field in item.items:
                        fieldName = field.attributes["name"]
                        component.fields[fieldName] = field.text if field.text is not None and field.text != "~" else ""
            self.components.append(component)

    def getGroupedComponents(self):
        """Вернуть компоненты, сгруппированные по обозначению и типу."""
        sortedComponents = sorted(
            self.components,
            key=lambda comp: comp.getRefNumber()
        )
        sortedComponents = sorted(
            sortedComponents,
            key=lambda comp: comp.getRefType()
        )
        groups = []
        compGroup = CompGroup(self)
        compRange = CompRange(self)
        excludedField = config.get("fields", "excluded")
        for comp in sortedComponents:
            if excludedField and excludedField in comp.fields:
                continue
            if not compRange.append(comp):
                if not compGroup.append(compRange):
                    groups.append(compGroup)
                    compGroup = CompGroup(self, compRange)
                compRange = CompRange(self, comp)
        if len(compRange) > 0:
            if not compGroup.append(compRange):
                groups.append(compGroup)
                compGroup = CompGroup(self, compRange)
        if len(compGroup) > 0:
            groups.append(compGroup)
        return groups
