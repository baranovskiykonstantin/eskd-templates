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

    def __init__(self, parent, name, items=None, text=None):
        """Создать элемент списка цепей.

        Каждый элемент:
        - должен иметь родителя или None -- если элемент корневой;
        - должен иметь имя;
        - может иметь дочерние элементы;
        - может содержать одну или несколько строк текста.

        Аргументы:
        parent (NetlistItem) -- родительский элемент;
        name (str) -- имя элемента;
        items (list of NetlistItem) -- массив дочерних элементов;
        text (str/list of str) -- текстовое значение элемента. Если значение
            состоит из нескольких строк, то оно будет представлено в виде
            списка строк (например для tstamps).

        """
        self.parent = parent
        self.name = name
        self.items = [] if items is None else items
        self.text = text

    def getText(self, name):
        """Получить текст элемента списка цепей с указанным именем.

        Будет возвращён текст дочернего элемента с указанным именем.
        Если элемента с указанным именем нет -- будет возвращена пустая строка.

        Аргументы:
        name (str) -- имя элемента.

        """
        for item in self.items:
            if item.name == name:
                if type(item.text) is str:
                    return item.text
                elif type(item.text) is list:
                    return ' '.join(item.text)
                else:
                    break
        return ""


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
        self._reset()
        with open(fileName, encoding="utf-8") as netlist:
            if self.fileName.endswith(".net"):
                self._content = netlist.read()
                self.data = self._parseNetItem(None)
                self._reset()
            elif self.fileName.endswith(".xml"):
                netlist.readline() # Пропустить первую строку (заголовок)
                self._content = netlist.read()
                self.data = self._parseXmlItem(None)
                self._reset()
            else:
                self._error("Формат файла не поддерживается.")

    def _reset(self):
        self._content = ""
        self._index = 0
        self._line = 1
        self._pos = 1

    def _error(self, message):
        raise ParseException(
            self._line,
            self._pos,
            message
        )

    def _hasChar(self):
        return self._index < len(self._content)

    def _getChar(self, offset=0):
        return self._content[self._index + offset]

    def _nextChar(self, offset=1):
        if offset > 1:
            self._nextChar(offset - 1)
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
                if character in " \t()\n":
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
            if character in " \t()\n":
                break
            name += character
            self._nextChar()
        else:
            self._error("Элемент неожиданно закончился!")
        if name == "":
            self._error("Элемент не имеет имени!")
        item = NetlistItem(parent, name)
        while self._hasChar():
            character = self._getChar()
            if character in ' \t\n':
                self._nextChar()
            elif character == '(':
                subitem = self._parseNetItem(item)
                item.items.append(subitem)
            elif character == ')':
                self._nextChar()
                break
            else:
                text = self._parseNetText()
                if item.text is None:
                    item.text = text
                else:
                    if type(item.text) is not list:
                        item.text = [item.text]
                    item.text.append(text)
        else:
            self._error(
                "Элемент неожиданно закончился " \
                "(должен заканчиваться символом ')')!"
            )
        return item

    @staticmethod
    def _formatNetText(text):
        text = text.replace("\\", "\\\\")
        text = text.replace("\"", "\\\"")
        text = '"{}"'.format(text)
        return text

    def _formatNetItem(self, item):
        output = '(' + item.name
        if item.items:
            output += '\n'
            for subitem in item.items:
                childText = self._formatNetItem(subitem)
                for line in childText.splitlines():
                    output += "\t{}\n".format(line)
        if item.text is not None:
            output = output.rstrip('\n')
            if type(item.text) is list:
                for val in item.text:
                    output += ' ' + self._formatNetText(val)
            else:
                output += ' ' + self._formatNetText(item.text)
        output += ')'
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

