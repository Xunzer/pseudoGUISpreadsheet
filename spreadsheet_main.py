from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidget, QTableWidgetItem, QItemDelegate, QLineEdit
import re
import sys
from collections import ChainMap
import math

# this regular expression matches cell references like "A1", "B2".
cell_references = re.compile(r'\b[A-Z][0-9]\b')

# function that generates the cell name by the column number and row number.
def cell_name(i, j):
    return f'{chr(ord("A")+j)}{i+1}'


class SpreadSheetDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super(SpreadSheetDelegate, self).__init__(parent)

    # define the editor widget (QLineEdit) for editing cells.
    def createEditor(self, parent, styleOption, index):
        editor = QLineEdit(parent)
        editor.editingFinished.connect(self.commitAndCloseEditor)
        return editor

    # commits the data from the editor to the model and closes the editor.
    def commitAndCloseEditor(self):
        editor = self.sender()
        self.commitData.emit(editor)
        self.closeEditor.emit(editor)

    # sets the editor's initial data from the model in the widget (updating purpose).
    def setEditorData(self, editor, index):
        editor.setText(index.model().data(index, Qt.ItemDataRole.EditRole))

    # writes the editor's data back to the model after entering new value in widget.
    def setModelData(self, editor, model, index):
        model.setData(index, editor.text())


class SpreadSheetItem(QTableWidgetItem):
    def __init__(self, siblings):
        super(SpreadSheetItem, self).__init__()
        self.siblings = siblings
        self.value = 0
        self.deps = set()
        self.reqs = set()
        self.error_shown = False 

    # retrieves the formula or content of the cell.
    def formula(self):
        return super().data(Qt.ItemDataRole.DisplayRole)

    # returns data based on the role requested.
    def data(self, role):
        if role == Qt.ItemDataRole.EditRole:
            return self.formula()
        if role == Qt.ItemDataRole.DisplayRole:
            return self.display()

        return super(SpreadSheetItem, self).data(role)

    # computes the value of the cell based on its formula.
    def calculate(self):
        formula = self.formula()

        if formula is None or formula == '':
            self.value = 0
            return

        # make sure lower case letters also match the cells
        formula = formula.upper()
        
        current_reqs = set(cell_references.findall(formula))

        name = cell_name(self.row(), self.column())

        # add this cell to the new requirement's dependents
        for r in current_reqs - self.reqs:
            self.siblings[r].deps.add(name)
        # add remove this cell from dependents no longer referenced
        for r in self.reqs - current_reqs:
            self.siblings[r].deps.remove(name)

        # look up the values of our required cells
        req_values = {r: self.siblings[r].value for r in current_reqs}
        # build an environment with these values and basic math functions
        environment = ChainMap(math.__dict__, req_values)
        previous_value = self.value
        # note that eval is DANGEROUS and should not be used in production
        try:
            self.value = eval(formula, {}, environment)
            self.error_shown = False  # Reset error flag
        except Exception as e:
            if not self.error_shown:
                # Show error message only once
                self.show_error_message(str(e))
                self.error_shown = True
            # Revert to previous value on error
            self.value = previous_value
        self.reqs = current_reqs

    def show_error_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setText("Calculation Error")
        msg_box.setInformativeText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec()

    # used to update all dependent cells.
    def propagate(self):
        for d in self.deps:
            self.siblings[d].calculate()
            self.siblings[d].propagate()

    # computes and returns the cell’s display value.
    def display(self):
        self.calculate()
        self.propagate()
        return str(self.value)


class SpreadSheet(QMainWindow):
    def __init__(self, rows, cols, parent=None):
        super(SpreadSheet, self).__init__(parent)

        self.rows = rows
        self.cols = cols
        self.cells = {}

        self.create_widgets()

    def create_widgets(self):
        table = self.table = QTableWidget(self.rows, self.cols, self)

        # creates a list of column headers based on the number of columns. 
        headers = [chr(ord('A') + j) for j in range(self.cols)]
        table.setHorizontalHeaderLabels(headers)

        # sets the custom delegate (SpreadSheetDelegate) for the table. 
        table.setItemDelegate(SpreadSheetDelegate(self))

        for i in range(self.rows):
            for j in range(self.cols):
                cell = SpreadSheetItem(self.cells)
                self.cells[cell_name(i, j)] = cell
                self.table.setItem(i, j, cell)

        # sets the QTableWidget as the central widget of the main window.
        self.setCentralWidget(table)

def main():
    app = QApplication(sys.argv)
    sheet = SpreadSheet(10, 10)
    sheet.resize(1040, 400)
    sheet.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()