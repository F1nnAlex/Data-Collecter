from PyQt5.QtGui import QIntValidator
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

        save_btn = QPushButton()
        save_btn.setText("Save")
        save_btn.clicked.connect(self.on_click)

        self.del_input_box = QLineEdit(self)
        int_validator = QIntValidator()
        self.del_input_box.setValidator(int_validator)

        del_btn = QPushButton()
        del_btn.setText("Delete Row")
        del_btn.clicked.connect(self.on_click_del)

        grid = QGridLayout()

        layout = QVBoxLayout()
        Vlayout = QVBoxLayout()
        Vlayout2 = QVBoxLayout()
        self.setLayout(grid)

        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)
        Vlayout.addWidget(self.input_lbl)
        Vlayout.addWidget(self.input_box)
        Vlayout.addWidget(save_btn)
        Vlayout.addWidget(self.del_input_box)
        Vlayout.addWidget(del_btn)

        grid.addLayout(layout, 0, 0)
        grid.addLayout(Vlayout, 0, 1)

        self.load_data()

    def on_click_del(self):
        self.delete_yes(int(self.del_input_box.text()))
        self.del_input_box.clear()
        self.load_data()

    @staticmethod
    def delete_yes(num):
        wb = openpyxl.load_workbook("Data.xlsx")
        sheet = wb.active

        sheet.delete_rows(int(num) + 1)
        print("hello")
        wb.save("Data.xlsx")

    def on_click(self):
        self.add_row()
        self.input_box.clear()
        self.load_data()

    def add_row(self):
        wb = openpyxl.load_workbook("Data.xlsx")
        sheet = wb.active

        max_rows = sheet.max_row

        found = False

        # Loop through each row until the next row is empty
        for row_num in range(1, max_rows + 1):
            cell_value = sheet.cell(row=row_num, column=1).value
            if cell_value == self.input_box.text():
                found = True
                # If found, increment the value in the next column
                current_count = sheet.cell(row=row_num, column=2).value or 0
                sheet.cell(row=row_num, column=2).value = current_count + 1
                print("found")
                break

        if not found:
            # If the search string wasn't found and the loop didn't break, add it to the next empty row with count 1
            sheet.cell(row=max_rows + 1, column=1).value = self.input_box.text()
            sheet.cell(row=max_rows + 1, column=2).value = 1
            print("added")

        # Save the changes to the Excel file
        wb.save("Data.xlsx")

    def load_data(self):
        path = "Data.xlsx"
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.worksheets[0]

        # Set size of excel sheet
        self.table_widget.setRowCount(sheet.max_row -1)
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
