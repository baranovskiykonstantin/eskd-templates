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
    multipliers = set(list(multipliersDict.keys()) + list(multipliersDict.values()))
    # 2u7, 2н7, 4m7, 5k1 ...
    regexpr1 = re.compile(
        r"^(\d+)({})(\d+)$".format('|'.join(multipliers))
    )
    # 2.7 u, 2700p, 4.7 m, 470u, 5.1 k, 510 ...
    regexpr2 = re.compile(
        r"^(\d+(?:[\.,]\d+)?)\s*({})?$".format('|'.join(multipliers))
    )

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
            if config.getboolean("spec", "add units"):
                value = self.getValueWithUnits()
            else:
                value = self.value
        elif name == "Посад.место":
            if config.getboolean("spec", "footprint only"):
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
        numValue = ""
        separator = ""
        if config.getboolean("spec", "space before units"):
            separator = ' '
        multiplier = ""
        units = ""
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
                    if re.match(Component.regexpr1, numValue):
                        searchRes = re.search(Component.regexpr1, numValue).groups()
                        numValue = "{},{}".format(searchRes[0], searchRes[2])
                        multiplier = searchRes[1]
                    elif re.match(Component.regexpr2, numValue):
                        searchRes = re.search(Component.regexpr2, numValue).groups()
                        numValue = searchRes[0]
                        multiplier = searchRes[1]
                    else:
                        numValue = ""
        elif self.getRefType().startswith('L') \
            and not self.value.endswith("Гн"):
                units = "Гн"
                numValue = self.value.rstrip('H')
                numValue = numValue.strip()
                if re.match(Component.regexpr1, numValue):
                    searchRes = re.search(Component.regexpr1, numValue).groups()
                    numValue = "{},{}".format(searchRes[0], searchRes[2])
                    multiplier = searchRes[1]
                elif re.match(Component.regexpr2, numValue):
                    searchRes = re.search(Component.regexpr2, numValue).groups()
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
                elif re.match(Component.regexpr1, numValue):
                    searchRes = re.search(Component.regexpr1, numValue).groups()
                    numValue = "{},{}".format(searchRes[0], searchRes[2])
                    multiplier = searchRes[1]
                elif re.match(Component.regexpr2, numValue):
                    searchRes = re.search(Component.regexpr2, numValue).groups()
                    numValue = searchRes[0]
                    if searchRes[1] is not None:
                        multiplier = searchRes[1]
                else:
                    numValue = ""
        if numValue:
            # Перевести множитель на русский
            if multiplier in Component.multipliersDict:
                multiplier = Component.multipliersDict[multiplier]
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

    def getSpecValue(self, name, singular=False, plural=False):
        """Вернуть преобразованное значение для спецификации.

        Вернуть приведённое к конечному виду значение одного из полей,
        используемых при построении спецификации.

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
            if config.getboolean("spec", "add units"):
                value = self.getValueWithUnits()
            else:
                value = self.value
        if value is None:
            value = ""
        return value

    def getExpandedValue(self):
        """Вернуть значение без множителя.

        Если компонент имеет значение физического характера (сопротивление,
        ёмкость, индуктивность), то будет возвращено абсолютное значение с
        учётом указанного множителя, например:
        1к5 => 1500
        0u33 => 0.00000033
        120 => 120
        и т.п.

        Возвращаемое значение (float) -- абсолютное значение.

        """
        extValue = float("inf")
        multiplierValues = {
            'G': 1e9,
            'Г': 1e9,
            'M': 1e6,
            'М': 1e6,
            'k': 1e3,
            'к': 1e3,
            'm': 1e-3,
            'м': 1e-3,
            'μ': 1e-6,
            'u': 1e-6,
            'U': 1e-6,
            'мк': 1e-6,
            'n': 1e-9,
            'н': 1e-9,
            'p': 1e-12,
            'п': 1e-12,
            None: 1
        }
        if self.getRefType().startswith('C'):
            value = self.value
            value = value.rstrip('F')
            value = value.rstrip('Ф')
            value = value.strip()
            if re.match(r"^\d+$", value):
                extValue = float(value) * 1e-12
            elif re.match(r"^\d+[\.,]\d+$", value):
                extValue = float(value.replace(',', '.')) * 1e-6
            elif re.match(Component.regexpr1, value):
                searchRes = re.search(Component.regexpr1, value).groups()
                numValue = "{}.{}".format(searchRes[0], searchRes[2])
                multiplier = multiplierValues[searchRes[1]]
                extValue = float(numValue) * multiplier
            elif re.match(Component.regexpr2, value):
                searchRes = re.search(Component.regexpr2, value).groups()
                numValue = searchRes[0]
                multiplier = multiplierValues[searchRes[1]]
                extValue = float(numValue.replace(',', '.')) * multiplier
        elif self.getRefType().startswith('L'):
            value = self.value
            value = value.rstrip('H')
            value = value.replace("Гн", "")
            value = value.strip()
            if re.match(r"^\d+(?:[\.,]\d+)?$", value):
                extValue = float(value.replace(',', '.')) * 1e-6
            elif re.match(Component.regexpr1, value):
                searchRes = re.search(Component.regexpr1, value).groups()
                numValue = "{}.{}".format(searchRes[0], searchRes[2])
                multiplier = multiplierValues[searchRes[1]]
                extValue = float(numValue) * multiplier
            elif re.match(Component.regexpr2, value):
                searchRes = re.search(Component.regexpr2, value).groups()
                numValue = searchRes[0]
                multiplier = multiplierValues[searchRes[1]]
                extValue = float(numValue.replace(',', '.')) * multiplier
        elif self.getRefType().startswith('R'):
            value = self.value
            value = value.rstrip('Ω')
            value = value.replace("Ом", "")
            value = value.replace("ohm", "")
            value = value.replace("Ohm", "")
            value = value.strip()
            if re.match(r"R\d+", value):
                numValue = value.replace('R', "0.")
                extValue = float(numValue)
            elif re.match(r"\d+R\d+", value):
                numValue = value.replace('R', ".")
                extValue = float(numValue)
            elif re.match(Component.regexpr1, value):
                searchRes = re.search(Component.regexpr1, value).groups()
                numValue = "{}.{}".format(searchRes[0], searchRes[2])
                multiplier = multiplierValues[searchRes[1]]
                extValue = float(numValue) * multiplier
            elif re.match(Component.regexpr2, value):
                searchRes = re.search(Component.regexpr2, value).groups()
                numValue = searchRes[0]
                if searchRes[1] is not None:
                    multiplier = multiplierValues[searchRes[1]]
                else:
                    multiplier = 1
                extValue = float(numValue.replace(',', '.')) * multiplier

        return extValue


class CompRange(Component):
    """Множество компонентов с одинаковыми параметрами.

    Этот класс описывает множество компонентов спецификации, которые
    имеют одинаковые тип, наименование, документ и примечание
    (отличаются только обозначением).

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
        if self.getSpecValue("type") == comp.getSpecValue("type") \
            and self.getSpecValue("name") == comp.getSpecValue("name") \
            and self.getSpecValue("doc") == comp.getSpecValue("doc") \
            and self.getSpecValue("comment") == comp.getSpecValue("comment"):
                self._refRange.append(comp.reference)
                return True
        return False

    def getRefRangeString(self):
        """Вернуть перечень обозначений множества одинаковых компонентов."""
        refStr = ""
        if len(self._refRange) > 1:
            # "VD1, VD2", "C8-C11", "R7, R9-R14" ...
            sortedRanges = sorted(
                self._refRange,
                key=lambda ref: self.getRefNumber(ref)
            )
            sortedRanges = sorted(
                sortedRanges,
                key=lambda ref: self.getRefType(ref)
            )
            prevType = self.getRefType(sortedRanges[0])
            prevNumber = self.getRefNumber(sortedRanges[0])
            counter = 0
            separator = ", "
            refStr = prevType + str(prevNumber)
            for nextRef in sortedRanges[1:]:
                currentType = self.getRefType(nextRef)
                currentNumber = self.getRefNumber(nextRef)
                if currentType == prevType \
                    and currentNumber == (prevNumber + 1):
                        prevNumber = currentNumber
                        counter += 1
                        if counter > 1:
                            separator = config.get("spec", "ref separator")
                        continue
                else:
                    if counter > 0:
                        refStr += separator + prevType + str(prevNumber)
                    separator = ', '
                    refStr += separator + currentType + str(currentNumber)
                    prevType = currentType
                    prevNumber = currentNumber
                    counter = 0
            if counter > 0:
                refStr += separator + prevType + str(prevNumber)
        else:
            # "R5"; "VT13" ...
            refStr = self.reference
        return refStr


class CompGroup():
    """Группа компонентов.

    Группой считается множество CompRange, которые имеют одинаковый тип.
    Если установлен параметр "separate group for each doc", то компоненты
    будут разбиваться на группы не только по типу, но и по документу.

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

    def sort(self, key=None):
        self._compRanges.sort(key=key)

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
        lastCompRange = self._compRanges[-1]
        if lastCompRange.getSpecValue("type") == compRange.getSpecValue("type"):
            if config.getboolean("spec", "separate group for each doc"):
                if lastCompRange.getSpecValue("doc") == compRange.getSpecValue("doc"):
                    # Если тип и документ не указаны, формировать группы
                    # на основе буквенной части обозначения.
                    if compRange.getSpecValue("type") \
                        or compRange.getSpecValue("doc") \
                        or lastCompRange.getRefType() == compRange.getRefType():
                            self._compRanges.append(compRange)
                            return True
            else:
                # Если тип не указан, формировать группы на основе
                # буквенной части обозначения.
                if compRange.getSpecValue("type") \
                    or lastCompRange.getRefType() == compRange.getRefType():
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

        currentType = self._compRanges[0].getSpecValue("type", plural=True)

        if not config.getboolean("spec", "title with doc"):
            return [currentType]

        currentName = self._compRanges[0].getSpecValue("name")
        currentDoc = self._compRanges[0].getSpecValue("doc")

        # Список уникальных пар Наименование-Документ
        nameDocList = []
        for compRange in self:
            currentName = compRange.getSpecValue("name")
            currentDoc = compRange.getSpecValue("doc")
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
                elif item.name == "fields":
                    for field in item.items:
                        fieldName = field.attributes["name"]
                        component.fields[fieldName] = field.text if field.text is not None and field.text != "~" else ""
            self.components.append(component)

    def getGroupedComponents(self):
        """Вернуть компоненты, сгруппированные по типу."""
        sortedComponents = sorted(
            self.components,
            key=lambda comp: comp.getSpecValue("name")
        )
        sortedComponents = sorted(
            sortedComponents,
            key=lambda comp: comp.getSpecValue("type")
        )
        sortedComponents = sorted(
            sortedComponents,
            key=lambda comp: "" if comp.getSpecValue("type") else comp.getRefType()
        ) # Компоненты без типа сортировать оп буквенной части обозначения
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

        # Группы компонентов должны быть отсортированы по буквенной части
        # обозначений.
        # Если группы имеют одинаковые буквенные обозначения - сортировать
        # по наименованию группы (тип или тип+документ).
        # Внутри группы, элементы перечисляются в порядке возрастания значения.
        for index in range(len(groups)):
            groups[index].sort(
                key=lambda compRange: compRange.getExpandedValue()
            )
        groups.sort(
            key=lambda group: group.getTitle()[:1]
        )
        groups.sort(
            key=lambda group: group[0].getRefType()
        )

        return groups
