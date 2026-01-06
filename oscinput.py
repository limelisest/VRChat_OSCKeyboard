import sys
from main_ui import Ui_MainWindow
from PyQt6.QtWidgets import QMessageBox, QMainWindow, QApplication
from pythonosc.udp_client import SimpleUDPClient

 
import win32gui
from ctypes import HRESULT
from ctypes.wintypes import HWND
from comtypes import IUnknown, GUID, COMMETHOD
import comtypes.client
import os

def popup_keyboard(a):
    toggle_tabtip()
    a.accept()
 
 
class ITipInvocation(IUnknown):
    _iid_ = GUID("{37c994e7-432b-4834-a2f7-dce1f13b834b}")
    _methods_ = [COMMETHOD([], HRESULT, "Toggle", (['in'], HWND, "hwndDesktop"))]
 
 
def toggle_tabtip():
    try:
        comtypes.CoInitialize()
        ctsdk = comtypes.client.CreateObject("{4ce576fa-83dc-4F88-951c-9d0782b4e376}", interface=ITipInvocation)
        ctsdk.Toggle(win32gui.GetDesktopWindow())
        comtypes.CoUninitialize()
    except OSError as e:
        os.system("C:\\PROGRA~1\\COMMON~1\\MICROS~1\\ink\\tabtip.exe")


class OSCInputWindows(QMainWindow, Ui_MainWindow):
    def __init__(self, parent = None):
        QMainWindow.__init__(self,parent=parent)
        self.setupUi(self)
        self.bt_open.clicked.connect(self.btnOpenKeyBoard)
        self.bt_close.clicked.connect(self.btnCloseKeyBoard)
        self.tx_input.returnPressed.connect(self.btnCloseKeyBoard)
        self.tx_input.mousePressEvent=popup_keyboard
        self.tx_input.textChanged.connect(self.textChange)
        self.ip="127.0.0.1"
        self.port=9000
        self.OSCClient=SimpleUDPClient(self.ip,self.port)

    def textChange(self):
        msg=self.tx_input.text()
        if msg == '':
            pass
        else:
            self.OSCClient.send_message('/chatbox/typing',True)

    def btnOpenKeyBoard(self):
        toggle_tabtip()
        self.tx_input.setFocus()


    def btnCloseKeyBoard(self):
        msg=self.tx_input.text()
        if msg == '':
            pass
        else:
            try:
                self.OSCClient.send_message('/chatbox/input',[msg,True,True])
                self.OSCClient.send_message('/chatbox/typing',False)

                self.tx_input.setText("")
                toggle_tabtip()
            
            except TypeError:
                self.OSCClient=SimpleUDPClient(self.ip,self.port)
                QMessageBox.information(self,"错误",'无法连接到VRChat,请再试一次')

if __name__ == '__main__':
    app=QApplication(sys.argv)
    OSCInputWindows=OSCInputWindows()
    OSCInputWindows.show()
    sys.exit(app.exec())

