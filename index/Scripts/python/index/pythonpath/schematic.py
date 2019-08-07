"""Объектное представление схемы."""

import re
from config import loadFromKicadbom2spec

REF_REGEXP = re.compile(r"([^0-9?]+)([0-9]+)")


class Component():
    """Данные о компоненте схемы."""

    def __init__(self, schematic):
        self.schematic = schematic
        self.reference = ""
        self.value = ""
        self.footprint = ""
        self.datasheet = ""
        self.fields = {}

    def getFieldValue(self, name):
        """Вернуть значение поля с указанным именем."""
        value = None
        if name == "Обозначение":
            value = self.reference
        elif name == "Значение":
            if self.schematic.settings.getboolean("index", "add units"):
                value = self.getValueWithUnits()
            else:
                value = self.value
        elif name == "Посад.место":
            value = self.footprint
        elif name == "Документация":
            value = self.datasheet
        elif name in self.fields:
            value = self.formatPattern(self.fields[name])
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

    def _getTypeSingularAndPlural(self):
        """Вернуть тип элемента в единственном и множественном числе."""
        typeValue = self.getIndexValue("type")
        singularAndPlural = re.match(r"^([^\s]+)\s*\(([^\s]+)\)$", typeValue)
        if singularAndPlural:
            return singularAndPlural.groups()
        elif self.schematic.typeNamesDict:
            for item in iter(self.schematic.typeNamesDict.items()):
                if typeValue in item:
                    return item
        return (typeValue, typeValue)

    def getTypeSingular(self):
        """Вернуть тип элемента в единственном числе."""
        return self._getTypeSingularAndPlural()[0]

    def getTypePlural(self):
        """Вернуть тип элемента в множественном числе."""
        return self._getTypeSingularAndPlural()[1]

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
        if self.schematic.settings.getboolean("index", "space before units"):
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
            numValue = numValue.replace('.', ',')
            return numValue + separator + multiplier + units
        return self.value

    def formatPattern(self, pattern):
        """Преобразовать шаблон.

        Шаблон представляет собой строку текста, в которой конструкции типа:
        {Префикс|НаименованиеПоля|Суффикс}
        будут преобразованы в записи вида:
        ПрефиксЗначениеПоляСуффикс

        Например:
        "МЛТ-0,5-{|Значение|}{-|Класс точности|}-В" -> "МЛТ-0,5-4,7кОм-±5%-В"
        Если значение поля пусто или указанного поля нет в компоненте, то
        соответствующий элемент шаблона игнорируется. Если, допустим, для
        приведённого выше примера, в компоненте нет поля "Класс точности", то
        результат будет следующим:
        "МЛТ-0,5-4,7кОм-В" (префикс '-' тоже отсутствует)

        Символы '{', '|', '}' имеют специальное назначение. Если в шаблоне
        требуется указать один из этих символов, то их нужно экранировать
        символом обратной косой черты ' \ ', например:
        "Обозначение компонента \{ {|Обозначение|} \} в фигурных скобках"

        """
        out = ""
        prefix = ""
        fieldName = ""
        suffix = ""
        escapeFlag = False
        prefixFlag = False
        fieldFlag = False
        suffixFlag = False

        for char in pattern:
            if char == '\\' and not escapeFlag:
                escapeFlag = True
                continue
            elif char == '{' and not escapeFlag:
                if prefixFlag or fieldFlag or suffixFlag:
                    # Шаблон имеет неверный формат
                    return pattern
                prefixFlag = True
            elif char == "|" and prefixFlag and not escapeFlag:
                prefixFlag = False
                fieldFlag = True
            elif char == "|" and fieldFlag and not escapeFlag:
                fieldFlag = False
                suffixFlag = True
            elif char == "|" and suffixFlag and not escapeFlag:
                # Шаблон имеет неверный формат
                return pattern
            elif char == "}" and not escapeFlag:
                if not suffixFlag:
                    # Шаблон имеет неверный формат
                    return pattern
                suffixFlag = False
                fieldValue = self.getFieldValue(fieldName)
                if fieldValue:
                    out += prefix + fieldValue + suffix
                prefix = fieldName = suffix = ""
            elif prefixFlag:
                prefix += char
            elif fieldFlag:
                fieldName += char
            elif suffixFlag:
                suffix += char
            else:
                out += char
            escapeFlag = False
        if prefixFlag or fieldFlag or suffixFlag:
            # Шаблон имеет неверный формат
            return pattern
        return out

    def getIndexValue(self, name):
        """Вернуть преобразованное значение для перечня.

        Вернуть приведённое к конечному виду значение одного из полей,
        используемых при построении перечня.

        Аргументы:
        name (str) -- название требуемого значения; может быть одним из:
            "type", "name", "doc", "comment".

        Возвращаемое значение (str) -- итоговое значение.

        """
        if name not in ("type", "name", "doc", "comment"):
            return ""
        fieldName = self.schematic.settings.get("fields", name)
        isPattern = re.match(
            r".*(?<!\\)\{.*(?<!\\)\|.*(?<!\\)\|.*(?<!\\)\}.*",
            fieldName
        )
        value = ""
        if isPattern:
            value = self.formatPattern(fieldName)
        else:
            value = self.getFieldValue(fieldName)
        if name == "name" and not value:
            return self.value
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
        adjustableField = self.schematic.settings.get(
            "fields",
            "adjustable"
        )
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
                            separator = '-'
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
        skipRefType = self.schematic.settings.getboolean(
            "index",
            "concatenate same name groups"
        )
        lastCompRange = self._compRanges[-1]
        if (lastCompRange.getRefType() == compRange.getRefType() or skipRefType) \
            and lastCompRange.getIndexValue("type") == compRange.getIndexValue("type"):
                self._compRanges.append(compRange)
                return True
        return False

    @staticmethod
    def _strCommon(str1, str2):
        """Определить общее начало двух строк.

        Вернуть подстроку, с которой начинаются обе указанные строки.

        Аргументы:
        str1 (str) -- первая строка;
        str2 (str) -- вторая строка.

        Возвращаемое значение (str) -- общее начало двух строк.

        """
        for i in range(len(str1)):
            if i == len(str2) or str1[i] != str2[i]:
                return str1[:i]
        return str1

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

        currentType = self._compRanges[0].getTypePlural()

        if not currentType:
            return []

        if not self.schematic.settings.getboolean("index", "title with doc"):
            return [currentType]

        currentName = self._compRanges[0].getIndexValue("name")
        currentDoc = self._compRanges[0].getIndexValue("doc")

        # Список уникальных пар Наименование-Документ
        nameDocList = []
        for compRange in self:
            currentName = compRange.getIndexValue("name")
            currentDoc = compRange.getIndexValue("doc")
            if not currentDoc:
                # Если имеются компоненты, в которых документ не указан,
                # то в заголовке для них будет указан только тип.
                currentName = ""
            for i in range(len(nameDocList)):
                savedName = nameDocList[i][0]
                savedDoc = nameDocList[i][1]
                if savedDoc == currentDoc:
                    commonName = self._strCommon(savedName, currentName)
                    commonName = commonName.rstrip(" -")
                    if commonName:
                        # Оставить только общую часть наименования
                        nameDocList[i][0] = commonName
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
        groupNames = []
        nameDocList.sort(key=lambda nameDoc: nameDoc[0])
        for nameDoc in nameDocList:
            name = currentType
            if nameDoc[0]:
                name += ' ' + nameDoc[0]
            if nameDoc[1]:
                name += ' ' + nameDoc[1]
            groupNames.append(name)

        return groupNames


class Schematic():
    """Данные о схеме и компонентах."""

    def __init__(self, settings):
        self.settings = settings
        self.title = ""
        self.number = ""
        self.company = ""
        self.developer = ""
        self.verifier = ""
        self.approver = ""
        self.components = []

        self.typeNamesDict = {}
        if self.settings.getboolean("settings", "compatibility mode"):
            # KB2S - kicadbom2spec
            settingsKB2S = loadFromKicadbom2spec()
            if settingsKB2S is not None:
                if settingsKB2S.has_section('group names singular'):
                    for index in settingsKB2S.options('group names singular'):
                        if settingsKB2S.has_option('group names plural', index):
                            singular = settingsKB2S.get('group names singular', index)
                            plural = settingsKB2S.get('group names plural', index)
                            self.typeNamesDict[singular] = plural

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
        for comp in sortedComponents:
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
