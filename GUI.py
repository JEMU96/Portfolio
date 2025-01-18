# -*- coding: utf-8 -*-
"""
INTERFAZ DE USUARIO SIMPLE EN LA QUE SE PUEDE
SELECCIONAR DE UNA LISTA DESPLEGABLE QUE 
PROGRAMA EJECUTAR.
"""

import sys
import importlib
import inspect
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QLabel, QWidget, 
                             QVBoxLayout, QHBoxLayout, QPushButton, 
                             QComboBox, QProgressBar, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt, QPropertyAnimation, QSequentialAnimationGroup, QPoint, QEvent
from PyQt5.QtGui import QIcon

# Buttons animation
class ShakingButton(QPushButton):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._originalPos = None

    def setOriginalPos(self):
        if self._originalPos is None:
            self._originalPos = self.pos()

    def shake(self):
        self.setOriginalPos()
        shake1 = QPropertyAnimation(self, b"pos")
        shake1.setStartValue(self._originalPos)
        shake1.setEndValue(self._originalPos + QPoint(-10, 0))
        shake1.setDuration(100)

        shake2 = QPropertyAnimation(self, b"pos")
        shake2.setStartValue(self._originalPos + QPoint(-10, 0))
        shake2.setEndValue(self._originalPos)
        shake2.setDuration(100)

        shake3 = QPropertyAnimation(self, b"pos")
        shake3.setStartValue(self._originalPos)
        shake3.setEndValue(self._originalPos + QPoint(10, 0))
        shake3.setDuration(100)

        shake4 = QPropertyAnimation(self, b"pos")
        shake4.setStartValue(self._originalPos + QPoint(10, 0))
        shake4.setEndValue(self._originalPos)
        shake4.setDuration(100)

        self._animationGroup = QSequentialAnimationGroup()
        self._animationGroup.addAnimation(shake1)
        self._animationGroup.addAnimation(shake2)
        self._animationGroup.addAnimation(shake3)
        self._animationGroup.addAnimation(shake4)
        self._animationGroup.start()

# Program selector class
class ProgramSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Program Selection")
        self.setGeometry(100, 100, 400, 300)  # x, y, width, height

        # Set the window icon
        self.setWindowIcon(QIcon("window_icon.png"))

        # Create a label
        self.helloMsg = QLabel("<b>Select Program to run.</b>", self)  # Use <b> for bold text
        self.helloMsg.setAlignment(Qt.AlignCenter)  # Center align the label text

        # Create a combo box (dropdown list)
        self.comboBox = QComboBox(self)
        values = ["", "IBM", "SAP", "WEB"]  
        self.comboBox.addItems(values)
        self.comboBox.currentTextChanged.connect(self.updateClassComboBox)

        # Create another combo box for classes
        self.classComboBox = QComboBox(self)

        # Create a progress bar
        self.progressBar = QProgressBar(self)
        self.progressBar.setAlignment(Qt.AlignCenter)

        # Create buttons
        self.continueBtn = ShakingButton("Continue!", self)
        self.continueBtn.setIcon(QIcon("continue_icon.png"))
        self.continueBtn.clicked.connect(self.onContinueClick)
        self.continueBtn.installEventFilter(self)

        self.cancelBtn = ShakingButton("Cancel", self)
        self.cancelBtn.setIcon(QIcon("cancel_icon.png"))
        self.cancelBtn.clicked.connect(self.onCancelClick)
        self.cancelBtn.installEventFilter(self)

        # Create a vertical layout and add the label, combo box, and progress bar
        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addWidget(self.helloMsg)
        vbox.addWidget(self.comboBox)
        vbox.addWidget(self.classComboBox)
        vbox.addWidget(self.progressBar)
        vbox.addStretch(1)

        # Create a horizontal layout for the buttons and add the buttons to it
        hbox_buttons = QHBoxLayout()
        hbox_buttons.addStretch(1)
        hbox_buttons.addWidget(self.continueBtn)
        hbox_buttons.addWidget(self.cancelBtn)
        hbox_buttons.addStretch(1)

        # Add the button layout to the vertical layout
        vbox.addLayout(hbox_buttons)

        # Create a horizontal layout to center everything
        hbox_main = QHBoxLayout()
        hbox_main.addStretch(1)
        hbox_main.addLayout(vbox)
        hbox_main.addStretch(1)

        # Set the main horizontal layout as the layout for the main window
        self.setLayout(hbox_main)

        self.show()
        self.raise_()
        self.activateWindow()

    def updateClassComboBox(self, module_name):
        self.classComboBox.clear()
        try:
            module = dynamic_import_module(module_name)
            functions = [name for name, obj in inspect.getmembers(module, inspect.isfunction) if obj.__module__ == module_name and not name.startswith('_')]
            self.classComboBox.addItems(functions)
            self.helloMsg.setText(f"<b>{module_name} loaded</b>")
        except ImportError:
            self.helloMsg.setText(f"<b>Failed to load {module_name}</b>")

    def onContinueClick(self):
        selected_module = self.comboBox.currentText()
        selected_function = self.classComboBox.currentText()
        self.helloMsg.setText(f"<b>Perfect! To new beginnings.<br>You've chosen {selected_function} from {selected_module}</b>")
        print(f"Module: {selected_module}, Function: {selected_function}")
        self.progressBar.setValue(0)  # Reset the progress bar
        Excel_open_read(self, selected_module, selected_function)

    def onCancelClick(self):
        currentText = self.helloMsg.text()
        if currentText == "<b>Why do you want to leave?</b>":
            self.close()
        else:
            self.helloMsg.setText("<b>Why do you want to leave?</b>") 

    def eventFilter(self, source, event):
        if event.type() == QEvent.Enter and isinstance(source, ShakingButton):
            source.shake()
        return super(ProgramSelector, self).eventFilter(source, event)

    def resizeEvent(self, event):
        self.continueBtn.setOriginalPos()
        self.cancelBtn.setOriginalPos()
        super().resizeEvent(event)

def dynamic_import_module(module_name):
    return importlib.import_module(module_name)

def show_warning(message):
    QMessageBox.warning(None, 'Warning', message)

def Excel_open_read(UI, selected_module, selected_function):
    # Create a file dialog to select the Excel file
    file_path, _ = QFileDialog.getOpenFileName(UI, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")

    if file_path:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name='Hoja1', header=None, dtype=str) # The excel is a template where only certain cells can be modified.
        # Process each row in the DataFrame
        total_rows = len(df)
        for index, row in enumerate(df.itertuples(index=False, name=None)):
            run_program(UI, selected_module, selected_function, *row)
            progress = int((index + 1) / total_rows * 100)
            UI.progressBar.setValue(progress)

def run_program(UI, selected_module_name, selected_function_name, *args):
    module = importlib.import_module(selected_module_name)
    func = getattr(module, selected_function_name, None)
    print(f"Module: {selected_module_name}, Function: {selected_function_name}")
    if func:
        func(*args)
    else:
        print(f"No function named '{selected_function_name}' found.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = ProgramSelector()
    sys.exit(app.exec_())
