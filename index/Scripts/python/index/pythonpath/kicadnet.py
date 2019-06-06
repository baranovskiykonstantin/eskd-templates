"""Объектное представление списка цепей KiCad."""

import html


class ParseException(Exception):
    """Ошибка при разборе структуры файла списка цепей."""

    def __init__(self, line, pos, message):
        Exception.__init__(self)
        self.value = "Строка {}, позиция {}:\n{}".format(
            line,
            pos,
            message
        )

    def __str__(self):
        return self.value


class NetlistItem():
    """Элемент списка цепей."""

    def __init__(self, parent, name, attributes=None, items=None, text=None):
        """Создать элемент списка цепей.

        Каждый элемент:
        - должен иметь родителя или None -- если элемент корневой;
        - должен иметь имя;
        - может иметь атрибуты;
        - может иметь дочерние (вложенные) элементы;
        - может содержать значение в виде строки текста.

        Аргументы:
        parent (NetlistItem) -- родительский элемент;
        name (str) -- имя элемента;
        attributes (dict of str) -- словарь атрибутов ("имя": "значение");
        items (list of NetlistItem) -- массив дочерних элементов;
        text (str) -- текстовое значение элемента.

        """
        self.parent = parent
        self.name = name
        self.attributes = {} if attributes is None else attributes
        self.items = [] if items is None else items
        self.text = text


class Netlist():
    """Список цепей."""

    def __init__(self, fileName):
        """Считать список цепей.

        Загрузить содержимое файла списка цепей KiCad (*.net, *.xml)
        и построить его объектное представление.

        Атрибуты:
        fileName (str) -- полное имя файла списка цепей.
        data (NetlistItem) -- объектное представление списка цепей.

        """
        self.fileName = fileName
        self.data = None
        self._content = ""
        self._index = 0
        self._line = 1
        self._pos = 1
        with open(fileName, encoding="utf-8") as netlist:
            if self.fileName.endswith(".net"):
                self._content = netlist.read()
                self.data = self._parseNetItem(None)
            elif self.fileName.endswith(".xml"):
                netlist.readline() # Пропустить первую строку (заголовок)
                self._content = netlist.read()
                self.data = self._parseXmlItem(None)
            else:
                self._error("Формат файла не поддерживается.")

    def _error(self, message):
        raise ParseException(
            self._line,
            self._pos,
            message
        )

    def _hasChar(self):
        return self._index < len(self._content)

    def _getChar(self, offset=0):
        return self._content[self._index+offset]

    def _nextChar(self, offset=1):
        if offset > 1:
            self._nextChar(offset-1)
        if self._hasChar() and self._getChar() == '\n':
            self._line += 1
            self._pos = 0
        self._index += 1
        self._pos += 1

    def _parseNetText(self):
        if not self._hasChar():
            return None
        text = ""
        quoted = False
        if self._getChar() == '"':
            quoted = True
            self._nextChar()
        if quoted:
            while self._hasChar():
                character = self._getChar()
                if character == '\n':
                    self._error(
                        "Значение неожиданно закончилось " \
                        "(должно заканчиваться символом '\"')!"
                    )
                previous = self._getChar(-1)
                self._nextChar()
                if character == '"' and previous != '\\':
                    text = text.replace("\\\"", "\"")
                    text = text.replace("\\\\", "\\")
                    break
                text += character
            else:
                self._error(
                    "Значение неожиданно закончилось " \
                    "(должно заканчиваться символом '\"')!"
                )
        else:
            while self._hasChar():
                character = self._getChar()
                if character in " ()\n":
                    break
                text += character
                self._nextChar()
            else:
                self._error("Значение неожиданно закончилось!")
        return text

    def _parseNetItem(self, parent):
        if not self._hasChar():
            return None
        if self._getChar() != '(':
            self._error("Элемент должен начинаться символом '('!")
        self._nextChar()
        name = ""
        while self._hasChar():
            character = self._getChar()
            if character in " ()\n":
                break
            name += character
            self._nextChar()
        else:
            self._error("Элемент неожиданно закончился!")
        if name == "":
            self._error("Элемент не имеет имени!")
        item = NetlistItem(parent, name)
        isAttribute = True
        while self._hasChar():
            character = self._getChar()
            if character == ' ':
                self._nextChar()
            elif character == '\n':
                isAttribute = False
                self._nextChar()
            elif character == '(':
                subitem = self._parseNetItem(item)
                if isAttribute:
                    item.attributes[subitem.name] = subitem.text
                else:
                    item.items.append(subitem)
            elif character == ')':
                self._nextChar()
                break
            else:
                text = self._parseNetText()
                if item.text is not None:
                    self._error(
                        "У элемента обнаружено второе значение " \
                        "(не может быть больше одного)!"
                    )
                item.text = text
        else:
            self._error(
                "Элемент неожиданно закончился " \
                "(должен заканчиваться символом ')')!"
            )
        return item

    @staticmethod
    def _formatNetText(text):
        if text == "" \
            or ' ' in text \
            or '(' in text \
            or ')' in text \
            or '"' in text:
                text = text.replace("\\", "\\\\")
                text = text.replace("\"", "\\\"")
                text = '"{}"'.format(text)
        return text

    def _formatNetItem(self, item):
        output = '(' + item.name
        for attrName in item.attributes:
            attrValue = item.attributes[attrName]
            attrValue = self._formatNetText(attrValue)
            output += " ({} {})".format(attrName, attrValue)
        if item.items:
            output += '\n'
            for subitem in item.items:
                childText = self._formatNetItem(subitem)
                for line in childText.splitlines():
                    output += "  {}\n".format(line)
        if item.text is not None:
            output += ' ' + self._formatNetText(item.text)
        else:
            output = output.rstrip('\n')
        output += ')'
        return output

    def _parseXmlAttribute(self):
        if not self._hasChar():
            return None
        name = ""
        while self._hasChar():
            character = self._getChar()
            self._nextChar()
            if character == '=':
                break
            elif not character.isalnum():
                self._error(
                    "В имени атрибута содержится недопустимый символ '{}'!".format(character)
                )
            name += character
        else:
            self._error("Элемент неожиданно закончился!")
        if name == "":
            self._error("Атрибут не имеет имени!")
        value = ""
        if self._getChar() != '"':
            self._error("Значение должно начинаться символом '\"')!")
        self._nextChar()
        while self._hasChar():
            character = self._getChar()
            self._nextChar()
            if character == '"':
                break
            value += character
        else:
            self._error(
                "Значение неожиданно закончилось " \
                "(должно заканчиваться символом '\"')!"
            )
        value = html.unescape(value)
        return (name, value)

    def _parseXmlItem(self, parent):
        if not self._hasChar():
            return None
        if self._getChar() != '<':
            self._error("Элемент должен начинаться символом '<'!")
        self._nextChar()
        name = ""
        while self._hasChar():
            character = self._getChar()
            if character in "/> ":
                break
            name += character
            self._nextChar()
        else:
            self._error("Элемент неожиданно закончился!")
        if name == "":
            self._error("Элемент не имеет имени!")
        item = NetlistItem(parent, name)
        # Атрибуты
        while self._hasChar():
            character = self._getChar()
            if character == ' ':
                self._nextChar()
            elif character == '>':
                self._nextChar()
                break
            elif character == '/':
                if self._getChar(+1) != '>':
                    self._error(
                        "Недопустимая последовательность символов " \
                        "(после '/' ожидался символ '>')!"
                    )
                self._nextChar(2)
                return item
            else:
                attrName, attrValue = self._parseXmlAttribute()
                item.attributes[attrName] = attrValue
        else:
            self._error("Элемент неожиданно закончился!")
        # Дочерние элементы
        closingTag = "</{}>".format(name)
        if self._getChar() == '\n':
            self._nextChar()
            while self._hasChar():
                character = self._getChar()
                if character in " \n":
                    self._nextChar()
                elif self._content[self._index:].startswith(closingTag):
                    self._nextChar(len(closingTag))
                    break
                elif character == '<':
                    subitem = self._parseXmlItem(item)
                    item.items.append(subitem)
                else:
                    self._error(
                        "Обнаружен недопустимый символ '{}'!".format(character)
                    )
            else:
                self._error("Элемент неожиданно закончился!")
        # Значение
        else:
            text = ""
            while self._hasChar():
                character = self._getChar()
                if self._content[self._index:].startswith(closingTag):
                    self._nextChar(len(closingTag))
                    break
                text += character
                self._nextChar()
            else:
                self._error("Элемент неожиданно закончился!")
            if text:
                item.text = text
        return item

    def _formatXmlItem(self, item):
        output = '<' + item.name
        for attrName in item.attributes:
            attrValue = item.attributes[attrName]
            attrValue = html.escape(attrValue)
            output += ' {}="{}"'.format(attrName, attrValue)
        if not item.text and not item.items:
            output += "/>"
            return output
        output += '>'
        if item.items:
            output += '\n'
            for subitem in item.items:
                childText = self._formatXmlItem(subitem)
                for line in childText.splitlines():
                    output += "  {}\n".format(line)
        if item.text:
            output += item.text
        output += "</{}>".format(item.name)
        return output

    def find(self, name, item=None):
        """Найти элемент списка цепей с указанным именем.

        Будет возвращён первый найденный элемент с указанным именем (порядок
        элементов соответствует тому, который имеется в файле списка цепей).
        Если элемент найти не удастся -- будет возвращено значение None.

        Аргументы:
        name (str) -- имя элемента.

        """
        if item is None:
            item = self.data
        if item.name == name:
            return item
        for subitem in item.items:
            foundItem = self.find(name, subitem)
            if foundItem is not None:
                return foundItem
        return None

    def items(self, name, item=None):
        """Перебор элементов списка цепей с указанным именем.

        Будет возвращён итератор, возвращающий элементы с указанным именем
        (порядок элементов соответствует тому, который имеется в файле списка
        цепей).

        Аргументы:
        name (str) -- имя элемента.

        """
        if item is None:
            item = self.data
        if item.name == name:
            yield item
        else:
            for subitem in item.items:
                for nextItem in self.items(name, subitem):
                    yield nextItem

    def save(self, fileName=None):
        """Записать данные списка цепей в файл.

        Аргументы:
        fileName (str) -- имя файла для записи.

        """
        if fileName is None:
            fileName = self.fileName
        with open(fileName, 'w', encoding='utf-8') as netlist:
            if fileName.endswith(".net"):
                netlist.write(self._formatNetItem(self.data))
            else:
                netlist.write('<?xml version="1.0" encoding="UTF-8"?>\n')
                netlist.write(self._formatXmlItem(self.data))
