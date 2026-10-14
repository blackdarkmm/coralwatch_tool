import sys
import os
from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.uic import loadUi
from function import VarManager, Function, TimeConsumingFunction, ToDo
from PyQt6.QtGui import QRegularExpressionValidator, QIcon
from PyQt6.QtCore import QRegularExpression
import ctypes



try:
    myappid = "coralwatch_app"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception as e:
    print(f"無法設定 AppUserModelID: {e}")

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.var = VarManager(self)
        self.funcs = Function(self,self.var)
        self.tcfuncs = TimeConsumingFunction(self,self.var)
        self.todo = ToDo(self,self.var)
        ui_path = os.path.join(os.path.dirname(__file__), "main.ui")
        icon_path = os.path.join(os.path.dirname(__file__), "coralwatch.ico")
        loadUi(ui_path, self)
        self.setWindowIcon(QIcon(icon_path))

        self.setup_signals()
        self.advanced_uisetting()
    
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.tcfuncs.revise_picture()
        self.funcs.font_size()

    def trigger_auto_check(self, obj_name):
        if obj_name == "excel_source":
            if not self.excel_source.text():
                return
            else:
                self.todo.excel_opne()
        
        elif obj_name == "picture_source":
            if not self.picture_source.text():
                return
            else:
                self.todo.picture_next()

    def advanced_uisetting(self):
        self.first.setValidator(QRegularExpressionValidator(QRegularExpression('^[0-9]*$')))
        self.last.setValidator(QRegularExpressionValidator(QRegularExpression('^[0-9]*$')))
        self.lightest_color.setValidator(QRegularExpressionValidator(QRegularExpression("^[BCDE1-6]*$")))
        self.darkest_color.setValidator(QRegularExpressionValidator(QRegularExpression("^[BCDE1-6]*$")))

    def setup_signals(self):
        self.confirm.clicked.connect(self.funcs.confirm)
        self.picture_source_search.clicked.connect(self.todo.picture_click)
        self.excel_source_search.clicked.connect(self.todo.excel_click)
        self.branching.clicked.connect(lambda: self.todo.button_click("BR"))
        self.boulder.clicked.connect(lambda: self.todo.button_click("BO"))
        self.plate.clicked.connect(lambda: self.todo.button_click("PL"))
        self.soft.clicked.connect(lambda: self.todo.button_click("SO"))
        self.B1.clicked.connect(lambda: self.todo.button_click("B1"))
        self.B2.clicked.connect(lambda: self.todo.button_click("B2"))
        self.B3.clicked.connect(lambda: self.todo.button_click("B3"))
        self.B4.clicked.connect(lambda: self.todo.button_click("B4"))
        self.B5.clicked.connect(lambda: self.todo.button_click("B5"))
        self.B6.clicked.connect(lambda: self.todo.button_click("B6"))
        self.C1.clicked.connect(lambda: self.todo.button_click("C1"))
        self.C2.clicked.connect(lambda: self.todo.button_click("C2"))
        self.C3.clicked.connect(lambda: self.todo.button_click("C3"))
        self.C4.clicked.connect(lambda: self.todo.button_click("C4"))
        self.C5.clicked.connect(lambda: self.todo.button_click("C5"))
        self.C6.clicked.connect(lambda: self.todo.button_click("C6"))
        self.D1.clicked.connect(lambda: self.todo.button_click("D1"))
        self.D2.clicked.connect(lambda: self.todo.button_click("D2"))
        self.D3.clicked.connect(lambda: self.todo.button_click("D3"))
        self.D4.clicked.connect(lambda: self.todo.button_click("D4"))
        self.D5.clicked.connect(lambda: self.todo.button_click("D5"))
        self.D6.clicked.connect(lambda: self.todo.button_click("D6"))
        self.E1.clicked.connect(lambda: self.todo.button_click("E1"))
        self.E2.clicked.connect(lambda: self.todo.button_click("E2"))
        self.E3.clicked.connect(lambda: self.todo.button_click("E3"))
        self.E4.clicked.connect(lambda: self.todo.button_click("E4"))
        self.E5.clicked.connect(lambda: self.todo.button_click("E5"))
        self.E6.clicked.connect(lambda: self.todo.button_click("E6"))
        self.excel_source.textChanged.connect(lambda _: self.trigger_auto_check("excel_source"))
        self.picture_source.textChanged.connect(lambda _: self.trigger_auto_check("picture_source"))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    app.aboutToQuit.connect(window.todo.excel_close)
    sys.excepthook = window.todo.excel_close
    window.show()
    sys.exit(app.exec())