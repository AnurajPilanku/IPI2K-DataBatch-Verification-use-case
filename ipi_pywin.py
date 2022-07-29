#Anuraj Pilanku
#ipi2k rpa automation
#Accessing the contextMenu

import pywinauto,time
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
import os


try:
    os.startfile(r"\\pmcsnttest64\DropBox\Cognizant\MEP_Verifier")  # (r"C:\Users\2040664\anuraj\ipi2k")
    time.sleep(15)

    app = Application(backend="uia").connect(title="MEP_Verifier")  # (title="ipi2k")
    app.MEP_Verifier.set_focus()
    fileItem = app.MEP_Verifier.ItemsView.get_item('mep_verifier_ipi2k_daily.exe')
    fileItem.set_focus()
    fileItem.right_click_input()
    app.ContextMenu["Open"].click_input()
    time.sleep(28)
    print("success")
except:
    app = Application(backend="uia").connect(title="MEP_Verifier")  # (title="ipi2k")
    app.Minimize()
    app.Restore()
    app.MEP_Verifier.set_focus()
    fileItem = app.MEP_Verifier.ItemsView.get_item('mep_verifier_ipi2k_daily.exe')
    fileItem.right_click_input()
    app.ContextMenu["Open"].click_input()
    time.sleep(25)
    print("success")
