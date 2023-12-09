from pywinauto.application import Application
import win32gui
import sys
from pywinauto import Desktop
top_windows = Desktop(backend="uia").windows() # or backend="win32" by default

# walk the returned list of wrappers
for w in top_windows:
    print(w.window_text())
    print(w.rectangle())

exit()
    
#app = Application().start("notepad.exe")
#app = Application(backend="uia").start('notepad.exe')
#app = Application(backend="uia").connect(path="notepad.exe")

#app.UntitledNotepad.menu_select("Help->About Notepad")
#app.AboutNotepad.OK.click()
#app.UntitledNotepad.Edit.type_keys("pywinauto Works!", with_spaces = True)

#from pywinauto import application
#app = application.Application()
#app.start("Notepad.exe")
#dlg_spec = app.window()
#dlg_spec.move_window(x=10, y=1, width=1500, height=1500, repaint=True)
#dlg_spec.move_window(x=500, y=500, width=500, height=500, repaint=True)



#from pywinauto import Application

#app = Application(backend="win32").start("C:\\Program Files\\Mozilla Firefox\\firefox.exe")
#app = Application(backend="win32").connect(title='Mozilla Firefox', found_index=0, timeout=5)


def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
        print ( hex( hwnd ), win32gui.GetWindowText( hwnd ) )

win32gui.EnumWindows( winEnumHandler, None )