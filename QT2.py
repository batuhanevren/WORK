# -*- coding: utf-8 -*-
"""
Created on Mon Jan  1 10:52:45 2018

@author: batuhan
"""

import sys
from PyQt5.QtGui import QApplication, QDialog
from ui_imagedialog import Ui_ImageDialog

app = QApplication(sys.argv)
window = QDialog()
ui = Ui_ImageDialog()
ui.setupUi(window)

window.show()
sys.exit(app.exec_())