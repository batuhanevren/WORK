# -*- coding: utf-8 -*-
"""
Created on Thu Dec 28 23:26:22 2017

@author: batuhan
"""
import sys
from PyQt5 import QtWidgets


def window():
    app = QtWidgets.QApplication(sys.argv)
    w = QtWidgets.QWidget()
    w.show()
    w.setWindowTitle('PyQt5 Lesson 1')
    sys.exit(app.exec_())
window ()

