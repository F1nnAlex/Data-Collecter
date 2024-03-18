import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *

class my_window(QMainWindow):
    def __init__(self):
        super(my_window, self).__init__()
        # Place window closer to center and resize when program is run
        self.setGeometry(500, 500, 500, 300)
        # Change application name
        self.setWindowTitle("Data Collector")
        # If hovering over nothing add "Data" tooltip
        self.setToolTip("Data")
        self.initUI()

    def initUI(self):
        self.input_lbl = QtWidgets.QLabel(self)
        self.input_lbl.setText("Give us data...")
        self.input_lbl.move(100, 50)

        self.input_box = QtWidgets.QLineEdit(self)
        self.input_box.move(200, 50)
        self.input_box.resize(200,32)

        self.save_btn = QtWidgets.QPushButton(self)
        self.save_btn.setText("Save")
        self.save_btn.clicked.connect(self.clicked)
        self.save_btn.move(150, 80)

        self.result_lbl = QtWidgets.QLabel(self)
        self.result_lbl.setText("Saved Data")
        self.result_lbl.move(250, 50)
        self.result_lbl.resize(200, 200)

    def clicked(self):
        self.result_lbl.setText(("Your Data"))
# Create a fucntion for the window
def window():
    # Collect all arguements to run code
    app = QApplication(sys.argv)
    # Create a visual window
    win = my_window()
    # Makes window visible
    win.show()
    # Exits program if x is pressed
    sys.exit(app.exec_())


# Run "window" function
window()
