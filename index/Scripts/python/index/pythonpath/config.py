"""Параметры работы.

Модуль предоставляет средства для загрузки и сохранения параметров.
Параметры хранятся в текстовом файле внутри odt-документа.

Внимание!
Глобальную переменную XSCRIPTCONTEXT обязательно нужно установить
после импорта модуля.

"""

import os
import sys
import configparser
import tempfile
import uno

XSCRIPTCONTEXT = None

DEFAULTS = {
    "index": {
        "source": "",
        "add units": "yes",
        "space before units": "no",
        "concatenate same name groups": "no",
        "title with doc": "no",
        "every group has title": "no",
        "empty row after group title": "no",
        "empty rows between diff ref": 1,
        "empty rows between diff type": 0,
        "extreme width factor": 80,
        "append rev table": "no",
        "pages rev table": 3,
    },
    "fields": {
        "type": "Тип",
        "name": "Наименование",
        "doc": "Документ",
        "comment": "Примечание",
        "adjustable": "Подбирают при регулировании",
    },
    "stamp": {
        "convert doc title": "yes",
        "convert doc id": "yes",
        "fill first usage": "yes",
    },
    "settings": {
        "pos x": "100",
        "pos y": "100",
        "set view options": "yes",
        "compatibility mode": "no",
    }
}

def load():
    """Загрузить настройки.

    Считать параметры работы из файла.

    Возвращаемое значение -- ConfigParser

    """
    doc = XSCRIPTCONTEXT.getDocument()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    config = configparser.ConfigParser(dict_type=dict)
    config.read_dict(DEFAULTS)
    fileAccess = ctx.ServiceManager.createInstance(
        "com.sun.star.ucb.SimpleFileAccess"
    )
    configFileUrl = "vnd.sun.star.tdoc:/{}/Scripts/python/index/settings.ini".format(doc.RuntimeUID)
    if fileAccess.exists(configFileUrl):
        fileStream = fileAccess.openFileRead(configFileUrl)
        configInput = ctx.ServiceManager.createInstance(
            "com.sun.star.io.TextInputStream"
        )
        configInput.setInputStream(fileStream)
        configInput.setEncoding("UTF-8")
        configString = configInput.readString((), False)
        config.read_string(configString, source=configFileUrl)
        configInput.closeInput()
    return config

def save(config):
    """Сохранить настройки.

    Записать параметры работы в файл.

    Аргументы:
    config (ConfigParser) -- параметры для сохранения.

    """
    doc = XSCRIPTCONTEXT.getDocument()
    serviceManager = XSCRIPTCONTEXT.getComponentContext().ServiceManager
    fileAccess = serviceManager.createInstance(
        "com.sun.star.ucb.SimpleFileAccess"
    )
    configPathUrl = "vnd.sun.star.tdoc:/{}/Scripts/python/index/".format(doc.RuntimeUID)
    if not fileAccess.exists(configPathUrl):
        fileAccess.createFolder(configPathUrl)
    configFileUrl = configPathUrl + "settings.ini"
    tempFile = tempfile.NamedTemporaryFile(
        mode="wt",
        encoding="UTF-8",
        delete=False
    )
    tempFileUrl = uno.systemPathToFileUrl(tempFile.name)
    with tempFile:
        config.write(tempFile)
    fileAccess.copy(tempFileUrl, configFileUrl)
    fileAccess.kill(tempFileUrl)

def loadFromKicadbom2spec():
    """Загрузить настройки kicadbom2spec.

    Считать параметры приложения kicadbom2spec.

    Возвращаемое значение -- ConfigParser или None в случае ошибки.

    """
    configPath = ""
    if sys.platform == "win32":
        configPath = os.path.join(
            os.environ["APPDATA"],
            "kicadbom2spec",
            "settings.ini"
        )
    else:
        configPath = os.path.join(
            os.path.expanduser("~/.config"),
            "kicadbom2spec",
            "settings.ini"
        )
    if os.path.isfile(configPath):
        settings = configparser.ConfigParser()
        settings.read(configPath)
        return settings
    return None
