import tempfile
import uno

def show(*args):
    """Показать справочное руководство."""
    context = XSCRIPTCONTEXT.getComponentContext()
    shell = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.system.SystemShellExecute",
        context
    )
    fileAccess = context.ServiceManager.createInstance(
        "com.sun.star.ucb.SimpleFileAccess"
    )
    tempFile = tempfile.NamedTemporaryFile(
        delete=False,
        prefix="help-",
        suffix=".html"
    )
    tempFileUrl = uno.systemPathToFileUrl(tempFile.name)
    helpFileUrl = "vnd.sun.star.tdoc:/{}/Scripts/python/doc/help.html".format(
        XSCRIPTCONTEXT.getDocument().RuntimeUID
    )
    tempFile.close()
    fileAccess.copy(helpFileUrl, tempFileUrl)
    shell.execute(
        "file://" + tempFile.name,
        "",
        0
    )
