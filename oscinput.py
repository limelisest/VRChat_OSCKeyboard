import sys

from yarg import get
from main_ui import Ui_MainWindow
from PyQt6.QtWidgets import QMessageBox, QMainWindow, QApplication
from PyQt6.QtCore import QTimer

import win32gui
from ctypes import HRESULT
from ctypes.wintypes import HWND
from comtypes import IUnknown, GUID, COMMETHOD
import comtypes.client
import os

from pythonosc.udp_client import SimpleUDPClient
import time
from getpower_win import get_battery_status


def popup_keyboard(a):
    toggle_tabtip()
    a.accept()


class ITipInvocation(IUnknown):
    _iid_ = GUID("{37c994e7-432b-4834-a2f7-dce1f13b834b}")
    _methods_ = [COMMETHOD([], HRESULT, "Toggle", (["in"], HWND, "hwndDesktop"))]


def toggle_tabtip():
    try:
        comtypes.CoInitialize()
        ctsdk = comtypes.client.CreateObject(
            "{4ce576fa-83dc-4F88-951c-9d0782b4e376}", interface=ITipInvocation
        )
        ctsdk.Toggle(win32gui.GetDesktopWindow())
        comtypes.CoUninitialize()
    except OSError as e:
        os.system("C:\\PROGRA~1\\COMMON~1\\MICROS~1\\ink\\tabtip.exe")


class OSCInputWindows(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent=parent)
        self.setupUi(self)
        self.bt_open.clicked.connect(self.btnOpenKeyBoard)
        self.bt_close.clicked.connect(self.btnCloseKeyBoard)
        self.tx_input.returnPressed.connect(self.btnCloseKeyBoard)
        self.tx_input.mousePressEvent = self.input_mouse_enter
        self.tx_input.textChanged.connect(self.textChange)

        self.ip = "127.0.0.1"
        self.port = 9000
        self.OSCClient = SimpleUDPClient(self.ip, self.port)

        self.last_send_time = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_battery)  # 绑定函数
        self.timer.start(5000) #发送间隔ms

        self.check_battery() #先运行一次

    def send_message(self, msg):
        if msg == "":
            pass
        else:
            try:
                self.OSCClient.send_message("/chatbox/input", [msg, True, True])
                self.OSCClient.send_message("/chatbox/typing", False)
                print(f"OSC send:{msg}")
                self.last_send_time = time.time()
            except TypeError:
                self.OSCClient = SimpleUDPClient(self.ip, self.port)
                QMessageBox.information(self, "错误", "无法连接到VRChat,请再试一次")

    def input_mouse_enter(self,a):
        if self.cb_autoOpenKeyboard.isChecked():
            toggle_tabtip()
        a.accept()

    def textChange(self):
        msg = self.tx_input.text()
        if msg == "":
            pass
        else:
            self.OSCClient.send_message("/chatbox/typing", True)

    def btnOpenKeyBoard(self):
        toggle_tabtip()
        self.tx_input.setFocus()

    def btnCloseKeyBoard(self):
        if self.cb_autoOpenKeyboard.isChecked():
            toggle_tabtip()
        msg = self.tx_input.text()
        self.send_message(msg)
        self.tx_input.setText("")

    def check_battery(self):
        changeing, battery = get_battery_status()
        self.cb_sendBettaryState.setText(f"自动发送电量状态 [{battery}%,{'充电中'if changeing == 1 else '未充电'}]")
        if time.time() - self.last_send_time < 15:  # 如果上次发送时间不超过15秒
            return
        msg = f"SteamDeck电量:{battery}%"
        if changeing:
            msg = msg + " 充电中"
        self.send_message(msg)




if __name__ == "__main__":
    app = QApplication(sys.argv)
    OSCInputWindows = OSCInputWindows()
    OSCInputWindows.show()
    sys.exit(app.exec())
