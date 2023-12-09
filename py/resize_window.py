from pywinauto.application import Application
import win32gui
import sys
from time import sleep


sleep(20)

for i in range(1,len(sys.argv)):
    _arg = sys.argv[i]
    (k,v) = _arg.split("=")
    locals()[k] = v
    
if "launch" not in locals().keys():
    launch=False

if "execproc" not in locals().keys():
    launch=False
    
if "args" not in locals().keys():
    launch=False
    
if launch == "True":
    app = Application(backend="win32").start(execproc + " " + args)
    sleep(0.4)
    app = Application(backend="win32").connect(title=proctitle, found_index=0, timeout=5)
else:
    app = Application(backend="win32").connect(title=proctitle, found_index=0, timeout=5)

dlg_spec = app.window(title=proctitle, found_index=0)
dlg_spec.move_window(x=int(x), y=int(y), width=int(width), height=int(height), repaint=True)

exit()