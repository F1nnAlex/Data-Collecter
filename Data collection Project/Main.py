from PyQt5.QtWidgets import (QApplication, QPushButton, QGridLayout, QLabel, QLineEdit,
                             QHBoxLayout, QVBoxLayout, QTableWidget, QTableWidgetItem, QWidget)
import sys
import openpyxl


class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        # Place window closer to center and resize when program is run
        self.setGeometry(500, 500, 500, 300)
        # Change application name
        self.setWindowTitle("Data Collector")
        # If hovering over nothing add "Data" tooltip
        self.setToolTip("Data")

        self.input_lbl = QLabel(self)
        self.input_lbl.setText("Give us data...")

        self.input_box = QLineEdit(self)
        self.input_value = self.input_box.text()

        save_btn = QPushButton()
        save_btn.setText("Save")
        save_btn.clicked.connect(self.on_click)

        grid = QGridLayout()

        layout = QVBoxLayout()
        Vlayout = QVBoxLayout()
        self.setLayout(grid)

        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)
        Vlayout.addWidget(self.input_lbl)
        Vlayout.addWidget(self.input_box)
        Vlayout.addWidget(save_btn)

        grid.addLayout(layout, 0, 0)
        grid.addLayout(Vlayout, 0, 1)

        self.load_data()

    def on_click(self):
        self.add_row()
        self.load_data()

    def add_row(self):
        wb = openpyxl.load_workbook("Data.xlsx")
        sheet = wb.active
        sheet.append([self.input_box.text(), 1])
        self.input_box.clear()
        wb.save("Data.xlsx")

    def load_data(self):
        path = "Data.xlsx"
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.worksheets[0]

        # Set size of excel sheet
        self.table_widget.setRowCount(sheet.max_row)
        self.table_widget.setColumnCount(sheet.max_column)

        list_values = list(sheet.values)
        self.table_widget.setHorizontalHeaderLabels(list_values[0])

        row_index = 0
        for value_tuple in list_values[1:]:
            col_index = 0
            for value in value_tuple:
                self.table_widget.setItem(row_index,col_index,QTableWidgetItem(str(value)))
                col_index += 1
            row_index += 1



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    app.exec_()
