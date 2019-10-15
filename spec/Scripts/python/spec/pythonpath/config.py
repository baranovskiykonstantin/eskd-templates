"""Параметры работы.

Модуль предоставляет средства для загрузки и сохранения параметров.
Параметры хранятся в текстовом файле внутри odt-документа.

"""

import os
import sys
from configparser import ConfigParser
import tempfile
import uno

# Глобальная переменная XSCRIPTCONTEXT устанавливается в listener.py:init()
XSCRIPTCONTEXT = None

SETTINGS = ConfigParser()

def load():
    """Загрузить настройки.

    Считать параметры работы из файла.

    """
    global SETTINGS
    doc = XSCRIPTCONTEXT.getDocument()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    fileAccess = ctx.ServiceManager.createInstance(
        "com.sun.star.ucb.SimpleFileAccess"
    )
    configFileUrl = "vnd.sun.star.tdoc:/{}/Scripts/python/spec/settings.ini".format(doc.RuntimeUID)
    if fileAccess.exists(configFileUrl):
        fileStream = fileAccess.openFileRead(configFileUrl)
        configInput = ctx.ServiceManager.createInstance(
            "com.sun.star.io.TextInputStream"
        )
        configInput.setInputStream(fileStream)
        configInput.setEncoding("UTF-8")
        configString = configInput.readString((), False)
        SETTINGS.read_string(configString, source=configFileUrl)
        configInput.closeInput()
    else:
        SETTINGS.read_dict(
            {
                "spec": {
                    "source": "",
                    "add units": "yes",
                    "space before units": "no",
                    "separate group for each doc": "no",
                    "title with doc": "no",
                    "every group has title": "no",
                    "reserve position numbers": "no",
                    "empty row after group title": "no",
                    "empty rows between diff type": 1,
                    "prohibit titles at bottom": "no",
                    "prohibit empty rows at top": "no",
                    "extreme width factor": 80,
                    "append rev table": "no",
                    "pages rev table": 3,
                },
                "sections": {
                    "documentation": "yes",
                    "assembly": "no",
                    "schematic": "yes",
                    "index": "yes",
                    "details": "yes",
                    "pcb": "yes",
                    "standard parts": "no",
                    "other parts": "yes",
                    "materials": "no",
                },
                "fields": {
                    "type": "Тип",
                    "name": "Наименование",
                    "doc": "Документ",
                    "comment": "Примечание",
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
        )

def save():
    """Сохранить настройки.

    Записать параметры работы в файл.

    """
    global SETTINGS
    doc = XSCRIPTCONTEXT.getDocument()
    serviceManager = XSCRIPTCONTEXT.getComponentContext().ServiceManager
    fileAccess = serviceManager.createInstance(
        "com.sun.star.ucb.SimpleFileAccess"
    )
    configPathUrl = "vnd.sun.star.tdoc:/{}/Scripts/python/spec/".format(doc.RuntimeUID)
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
        SETTINGS.write(tempFile)
    fileAccess.copy(tempFileUrl, configFileUrl)
    fileAccess.kill(tempFileUrl)

def get(section, option):
    """Получить значение параметра "option" из раздела "section"."""
    global SETTINGS
    return SETTINGS.get(section, option)

def getboolean(section, option):
    """Получить булево значение параметра "option" из раздела "section"."""
    global SETTINGS
    return SETTINGS.getboolean(section, option)

def getint(section, option):
    """Получить целочисленное значение параметра "option" из раздела "section"."""
    global SETTINGS
    return SETTINGS.getint(section, option)

def set(section, option, value):
    """Установить значение "value" параметру "option" из раздела "section"."""
    global SETTINGS
    return SETTINGS.set(section, option, value)

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
        settings = ConfigParser()
        settings.read(configPath)
        return settings
    return None
